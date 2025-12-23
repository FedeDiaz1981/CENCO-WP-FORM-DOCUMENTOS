import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";

type BodyMap = Record<string, any>;

export class SpService {
  constructor(private spHttpClient: SPHttpClient, private siteUrl: string) {}

  private escListTitle(listTitle: string): string {
    return listTitle.replace(/'/g, "''");
  }

  private contains(haystack: string, needle: string): boolean {
    return haystack && needle ? haystack.indexOf(needle) !== -1 : false;
  }

  private getItemEntityType(listTitle: string): string {
    const safe = listTitle.replace(/ /g, "_x0020_");
    return `SP.Data.${safe}ListItem`;
  }

  private async safeReadText(res: SPHttpClientResponse): Promise<string> {
    try {
      return await res.text();
    } catch {
      try {
        const anyRes: any = res as any;
        if (anyRes?.nativeResponse?.clone) {
          const clone = anyRes.nativeResponse.clone();
          return await clone.text();
        }
      } catch {
        /* ignore */
      }
      return "";
    }
  }

  async getTiposFormulario(): Promise<any[]> {
    const url =
      `${this.siteUrl}/_api/web/lists/getbytitle('` +
      `${this.escListTitle("Tipo formulario")}')/items` +
      `?$select=Id,Title,orden,Template&$orderby=orden asc`;

    const res = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) throw new Error(await this.safeReadText(res));
    const j: any = await res.json();
    return Array.isArray(j) ? j : j?.value ?? j?.d?.results ?? [];
  }

  // ✅ CAMBIO: ahora recibe múltiples archivos
  async createFormulario(body: BodyMap, archivos?: File[]): Promise<number> {
    const listTitle = "Formularios";
    const listTitleEsc = this.escListTitle(listTitle);

    const norm = this.normalizeForRest(body);

    const verboseBody: any = {
      __metadata: { type: this.getItemEntityType(listTitle) },
    };

    for (const k in norm) {
      if (!Object.prototype.hasOwnProperty.call(norm, k)) continue;
      const v = (norm as any)[k];
      if (v === undefined || v === null || v === "") continue;
      verboseBody[k] = v;
    }

    const itemsUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${listTitleEsc}')/items`;
    const createRes = await this.postVerboseJson(itemsUrl, verboseBody);
    const created: any = await createRes.json();
    const id: number = created?.d?.Id ?? created?.Id ?? created?.id;
    if (!id) throw new Error("No se obtuvo el Id del ítem creado.");

    // ✅ adjuntar N archivos
    const files = Array.isArray(archivos) ? archivos : [];
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      if (!f) continue;
      await this.addAttachment(listTitle, id, f, true);
    }

    return id;
  }

  private normalizeForRest(input: BodyMap): BodyMap {
    const out: BodyMap = {};

    for (const k in input) {
      if (!Object.prototype.hasOwnProperty.call(input, k)) continue;
      const v = (input as any)[k];
      if (v === undefined || v === null || v === "") continue;

      // Multi-user field: acepta array y lo convierte a { results: [] }
      if (k.toLowerCase() === "usuarioregistradorid") {
        if (
          v &&
          typeof v === "object" &&
          Object.prototype.hasOwnProperty.call(v, "results")
        ) {
          // ✅ FIX dot-notation
          out.usuarioregistradorId = {
            results: (v as any).results as number[],
          };
        } else if (Object.prototype.toString.call(v) === "[object Array]") {
          // ✅ FIX dot-notation
          out.usuarioregistradorId = { results: v as number[] };
        }
        continue;
      }

      // Lookups numéricos terminan con Id
      const endsWithId = k.length >= 2 && k.slice(-2) === "Id";
      if (endsWithId && typeof v === "number") {
        out[k] = v;
        continue;
      }

      out[k] = v;
    }

    return out;
  }

  private async postVerboseJson(
    url: string,
    body: any
  ): Promise<SPHttpClientResponse> {
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": "",
      },
      body: JSON.stringify(body),
    };

    const res = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!res.ok) {
      const t = await this.safeReadText(res);
      throw new Error(t || `Error POST ${url}`);
    }

    return res;
  }

  private stripDiacriticsLegacy(s: string): string {
    if (!s) return "";
    const map: Array<[RegExp, string]> = [
      [/[áàäâãå]/gi, "a"],
      [/[éèëê]/gi, "e"],
      [/[íìïî]/gi, "i"],
      [/[óòöôõ]/gi, "o"],
      [/[úùüû]/gi, "u"],
      [/[ñ]/gi, "n"],
      [/[ç]/gi, "c"],
      [/[\u00E1\u00E0\u00E4\u00E2\u00E3\u00E5]/g, "a"],
      [/[\u00E9\u00E8\u00EB\u00EA]/g, "e"],
      [/[\u00ED\u00EC\u00EF\u00EE]/g, "i"],
      [/[\u00F3\u00F2\u00F6\u00F4\u00F5]/g, "o"],
      [/[\u00FA\u00F9\u00FC\u00FB]/g, "u"],
      [/[\u00F1]/g, "n"],
      [/[\u00E7]/g, "c"],
    ];
    let out = s;
    for (let i = 0; i < map.length; i++) out = out.replace(map[i][0], map[i][1]);
    return out;
  }

  private sanitizeFileName(original: string): string {
    const MAX_BASE = 80;
    const MAX_EXT = 16;

    const trimmed = (original || "").replace(/^\s+|\s+$/g, "");
    const dot = trimmed.lastIndexOf(".");
    const ext = dot > 0 ? trimmed.substring(dot) : "";
    const baseRaw = dot > 0 ? trimmed.substring(0, dot) : trimmed;

    const base1 = this.stripDiacriticsLegacy(baseRaw);
    const base2 = base1.replace(/['"#%&*:<>?/\\{|}~]/g, "_");
    const base3 = base2.replace(/\s+/g, " ");
    const base4 = base3.replace(/^[.\s]+|[.\s]+$/g, "") || "archivo";

    const safeBase =
      base4.length > MAX_BASE ? base4.substring(0, MAX_BASE) : base4;
    const safeExt = ext
      .replace(/['"#%&*:<>?/\\{|}~\s]/g, "")
      .substring(0, MAX_EXT);

    let finalName = safeExt
      ? safeBase + "." + safeExt.replace(/^\.+/, "")
      : safeBase;

    finalName = finalName.replace(/[.\s]+$/g, "");
    return finalName;
  }

  private async addAttachment(
    listTitle: string,
    itemId: number,
    file: File,
    overwrite: boolean = true
  ): Promise<void> {
    if (!file || file.size === 0) throw new Error("El archivo está vacío.");

    const listTitleEsc = this.escListTitle(listTitle);
    const base = `${this.siteUrl}/_api/web/lists/getbytitle('${listTitleEsc}')/items(${itemId})/AttachmentFiles`;

    const cleaned = this.sanitizeFileName(file.name);
    const safeODataName = cleaned.replace(/'/g, "''");

    // 1) borrar si existe (mantengo tu comportamiento; overwrite hoy no cambia reglas)
    try {
      const delUrl = `${base}/getByFileName('${safeODataName}')`;
      await this.spHttpClient.post(delUrl, SPHttpClient.configurations.v1, {
        headers: {
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE",
          Accept: "*/*",
        },
      });
    } catch {
      /* ignorar */
    }

    // 2) subir archivo
    const up = await this.spHttpClient.post(
      `${base}/add(FileName='${safeODataName}')`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "*/*",
          "Content-Type": "application/octet-stream",
        },
        body: file,
      }
    );

    if (!up.ok) {
      const msg = (await this.safeReadText(up)) || "";
      const lower = msg.toLowerCase();
      if (
        this.contains(lower, "already in use") ||
        this.contains(lower, "already exists") ||
        this.contains(lower, "body stream already read")
      ) {
        return;
      }
      throw new Error(msg || `Error adjuntando archivo: HTTP ${up.status}`);
    }
  }
}
