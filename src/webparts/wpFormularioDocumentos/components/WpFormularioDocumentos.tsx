import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import { IWpFormularioDocumentosProps } from "./IWpFormularioDocumentosProps";
import { SpService } from "../services/sp.service";

/* ==== SharePoint REST ==== */
import { SPHttpClient } from "@microsoft/sp-http";

/* ==== Fluent UI ==== */
import {
  ThemeProvider,
  createTheme,
  Text,
  Label,
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Nav,
  INavLinkGroup,
  INavLink,
  getTheme,
  mergeStyleSets,
  FontWeights,
  Stack,
  IStackStyles,
  Dropdown,
  IDropdownOption,
  IconButton,
  IButtonStyles,
  Modal, // ✅ NUEVO
} from "@fluentui/react";

/* People Picker (solo cuando proveedor = false) */
import {
  NormalPeoplePicker,
  IBasePickerSuggestionsProps,
} from "@fluentui/react/lib/Pickers";
import { IPersonaProps } from "@fluentui/react/lib/Persona";

/* ===== tema corporativo Cencosud ===== */
const cencoTheme = createTheme({
  palette: {
    themePrimary: "#005596",
    themeLighterAlt: "#f2f7fb",
    themeLighter: "#d6e5f2",
    themeLight: "#b7d0e7",
    themeTertiary: "#71a5d0",
    themeSecondary: "#357fba",
    themeDarkAlt: "#004d87",
    themeDark: "#00416f",
    themeDarker: "#002f51",
    neutralLighterAlt: "#f5f5f5",
    neutralLighter: "#f0f0f0",
    neutralLight: "#e6e6e6",
    neutralQuaternaryAlt: "#d6d6d6",
    neutralQuaternary: "#cccccc",
    neutralTertiaryAlt: "#c4c4c4",
    neutralTertiary: "#333333",
    neutralSecondary: "#2d2d2d",
    neutralPrimaryAlt: "#272727",
    neutralPrimary: "#333333",
    neutralDark: "#1f1f1f",
    black: "#1a1a1a",
    white: "#ffffff",
  },
  effects: {
    roundedCorner2: "12px",
    elevation8: "0 6px 18px rgba(0,0,0,.05)" as any,
  },
});

/* ===== Tipos ===== */
interface ICodigoDocumentoRow {
  codigo: string;
  descripcion: string;
}

interface IFormState {
  fechaderegistro: string;
  ruc: string;
  proveedorId?: number;
  usuarioregistradorIds: number[];
  Title: string;
  codigodecontrato: string;
  periododesde: string;
  periodohasta: string;
  anio: string; // vigencia calculada
  codigosDocumentos: ICodigoDocumentoRow[]; // SCTR
  archivos: File[];
  isSaving: boolean;
  error?: string;
  success?: string;
  tipodeformulario: string;
}

interface ITipoFormularioItem {
  Id: number;
  Title: string;
  orden: number;
  template: "A" | "B" | "C" | string;
}

function sectionsForTemplate(tpl?: string): Set<number> {
  switch ((tpl || "").toUpperCase()) {
    case "A":
      return new Set([1, 2, 4]);
    case "B":
      return new Set([1, 3, 4]);
    case "C":
      return new Set([1, 3, 4]);
    default:
      return new Set();
  }
}

function isSctrExcel(tipodeformulario?: string): boolean {
  const t = (tipodeformulario || "").trim().toUpperCase();
  // eslint-disable-next-line no-console
  console.log("Formulario: " + t);
  return (
    t === "SCTR Y PLANTILLAS DE EXCEL" || t === "SCTR Y PLANTILLAS DE EXCELL"
  );
}

function isCargarContratos(tipodeformulario?: string): boolean {
  const t = (tipodeformulario || "").trim().toUpperCase();
  return t === "CARGAR CONTRATOS";
}

function visibleSectionsFor(
  selectedTemplate?: string,
  tipodeformulario?: string
): Set<number> {
  if (isSctrExcel(tipodeformulario)) return new Set([1, 2, 3, 4]);
  return sectionsForTemplate(selectedTemplate);
}

const initialState: IFormState = {
  fechaderegistro: "",
  ruc: "",
  proveedorId: undefined,
  usuarioregistradorIds: [],
  Title: "",
  codigodecontrato: "",
  periododesde: "",
  periodohasta: "",
  anio: "",
  codigosDocumentos: [{ codigo: "", descripcion: "" }],
  archivos: [],
  isSaving: false,
  tipodeformulario: "",
};

/* ---- Presentacional ---- */
function Section({
  title,
  children,
  classes,
}: {
  title: string;
  children: React.ReactNode;
  classes: ReturnType<typeof getClasses>;
}): JSX.Element {
  return (
    <section className={classes.panel}>
      <div className={classes.panelHeader}>
        <Text
          variant="large"
          block
          styles={{ root: { fontWeight: FontWeights.semibold } }}
        >
          {title}
        </Text>
      </div>
      <div className={classes.panelBody}>{children}</div>
    </section>
  );
}

const toSpDate = (yyyyMmDd: string): string | undefined => {
  const v = (yyyyMmDd || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return undefined;
  return `${v}T00:00:00Z`;
};

const getTodayYMD = (): string => {
  const d = new Date();
  const yyyy = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const mm = (m < 10 ? "0" : "") + m;
  const dd = (day < 10 ? "0" : "") + day;
  return `${yyyy}-${mm}-${dd}`;
};

const getClasses = (theme = getTheme()) =>
  mergeStyleSets({
    page: {
      background: theme.palette.neutralLighterAlt,
      minHeight: "100%",
      padding: 12,
      [`@media (min-width: 640px)`]: { padding: 16 },
      [`@media (min-width: 1024px)`]: { padding: 20 },
    },
    container: { maxWidth: 1180, margin: "0 auto" },
    mainRow: {
      display: "flex",
      alignItems: "stretch",
      gap: 16,
      flexWrap: "nowrap",
      [`@media (max-width: 920px)`]: { flexWrap: "wrap" },
    },
    leftCol: {
      flex: "0 0 280px",
      width: 280,
      position: "sticky",
      top: 12,
      alignSelf: "flex-start",
      [`@media (max-width: 920px)`]: {
        position: "static",
        width: "100%",
        flex: "1 1 100%",
      },
    },
    rightCol: {
      flex: "1 1 0",
      minWidth: 0,
      display: "flex",
      flexDirection: "column",
      gap: 12,
    },
    panel: {
      background: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 12,
      boxShadow: theme.effects.elevation8,
      overflow: "hidden",
      animation: "fadeIn .18s ease-out both",
    },
    panelHeader: {
      background: theme.palette.themeLighterAlt,
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      padding: "12px 16px",
    },
    panelBody: {
      padding: 16,
      [`@media (min-width: 1024px)`]: { padding: 20 },
    },
    row: {
      display: "grid",
      gridTemplateColumns: "repeat(12, 2fr)",
      gap: 12,
      [`@media (min-width: 1024px)`]: { gap: 16 },
    },
    c12: { gridColumn: "span 12" },
    c8: {
      gridColumn: "span 12",
      [`@media (min-width: 1024px)`]: { gridColumn: "span 8" },
    },
    c7: {
      gridColumn: "span 12",
      [`@media (min-width: 768px)`]: { gridColumn: "span 7" },
    },
    c6: {
      gridColumn: "span 12",
      [`@media (min-width: 768px)`]: { gridColumn: "span 6" },
    },
    c5: {
      gridColumn: "span 12",
      [`@media (min-width: 1024px)`]: { gridColumn: "span 5" },
    },
    c4: {
      gridColumn: "span 12",
      [`@media (min-width: 1024px)`]: { gridColumn: "span 4" },
    },

    // ✅ NUEVO: label + icon
    fieldLabelRow: {
      display: "flex",
      alignItems: "center",
      gap: 6,
    },
    helpIcon: { marginTop: 2 },

    // ✅ NUEVO: título "Plazo de contrato"
    plazoHeader: {
      gridColumn: "span 12",
      textAlign: "center",
      fontWeight: 600,
      color: theme.palette.neutralSecondary,
      paddingTop: 6,
      borderTop: `1px solid ${theme.palette.neutralLight}`,
      marginTop: 6,
    },

    // ✅ NUEVO: modal
    modalBody: { padding: 12 },
    modalImg: {
      maxWidth: "90vw",
      maxHeight: "80vh",
      display: "block",
      margin: "0 auto",
      borderRadius: 8,
    },

    navPanel: {
      background: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 12,
      boxShadow: theme.effects.elevation8,
      overflow: "hidden",
    },
    navHeader: {
      padding: "12px 16px",
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      background: theme.palette.themeLighterAlt,
    },
    navBody: { padding: 12 },
    navResponsive: {
      maxHeight: "calc(100vh - 220px)",
      overflowY: "auto",
      selectors: {
        "&& .ms-Nav-link": {
          display: "block !important",
          height: "auto !important",
          minHeight: "44px !important",
          padding: "12px 14px !important",
          borderRadius: "10px !important",
          border: "1px solid transparent !important",
        },
        "&& .ms-Nav-linkText": {
          whiteSpace: "normal !important",
          wordBreak: "break-word !important",
          lineHeight: "1.35 !important",
          fontSize: "15px !important",
          fontWeight: "600 !important",
        },
        "&& .ms-Nav-link:hover": {
          background: `${theme.palette.themeLighterAlt} !important`,
          borderColor: `${theme.palette.themeLighter} !important`,
        },
        "&& .ms-Nav .is-selected > .ms-Nav-link": {
          background: `${theme.palette.themeLighter} !important`,
          borderColor: `${theme.palette.themePrimary} !important`,
        },
      },
    },

    fileRow: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: 8,
      padding: "8px 10px",
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 10,
      background: theme.palette.white,
      marginTop: 8,
    },

    codigosRow: {
      display: "grid",
      gridTemplateColumns: "minmax(0, 2fr) minmax(0, 1fr) 36px",
      gap: 8,
      alignItems: "end",
      width: "100%",
      [`@media (max-width: 640px)`]: { gridTemplateColumns: "1fr" },
    },
    codigosDelete: {
      justifySelf: "end",
      marginBottom: 2,
      [`@media (max-width: 640px)`]: { justifySelf: "start" },
    },

    actionsBar: {
      display: "flex",
      flexWrap: "wrap",
      gap: 10,
      alignItems: "center",
      marginTop: 10,
    },

    "@global": {
      "@keyframes fadeIn": {
        from: { opacity: 0, transform: "translateY(4px)" },
        to: { opacity: 1, transform: "none" },
      },
    },
  });

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Usuarios",
  noResultsFoundText: "Sin resultados",
};

export default function WpFormularioDocumentos(
  props: IWpFormularioDocumentosProps
): JSX.Element {
  const [state, setState] = useState<IFormState>(initialState);

  // ✅ NUEVO: modal ayudas
  const [helpOpen, setHelpOpen] = useState<boolean>(false);
  const [helpImg, setHelpImg] = useState<string>("");
  const [helpTitle, setHelpTitle] = useState<string>("");

  const openHelp = (title: string, url?: string): void => {
    const u = (url || "").trim();
    if (!u) return;
    setHelpTitle(title);
    setHelpImg(u);
    setHelpOpen(true);
  };
  const closeHelp = (): void => {
    setHelpOpen(false);
    setHelpImg("");
    setHelpTitle("");
  };

  // re-hidratar proveedor/RUC cuando guardás o limpiás
  const [proveedorRefreshKey, setProveedorRefreshKey] = useState<number>(0);
  const refreshProveedor = (): void => setProveedorRefreshKey((k) => k + 1);

  const submitLockedRef = useRef<boolean>(false);
  const acquireLock = (): boolean => {
    if (submitLockedRef.current) return false;
    submitLockedRef.current = true;
    return true;
  };
  const releaseLock = (): void => {
    submitLockedRef.current = false;
  };

  const sp = useMemo(
    () => new SpService((props as any).spHttpClient, props.siteUrl),
    [props.spHttpClient, props.siteUrl]
  );

  const [tipos, setTipos] = useState<ITipoFormularioItem[]>([]);
  const [loadingTipos, setLoadingTipos] = useState<boolean>(true);
  const [errorTipos, setErrorTipos] = useState<string | undefined>(undefined);

  const [proveedorOptions, setProveedorOptions] = useState<IDropdownOption[]>(
    []
  );
  const [peopleSelected, setPeopleSelected] = useState<IPersonaProps[]>([]);
  const [selectedId, setSelectedId] = useState<number | undefined>(undefined);
  const [selectedTemplate, setSelectedTemplate] = useState<string | undefined>(
    undefined
  );

  const visible = useMemo(
    () => visibleSectionsFor(selectedTemplate, state.tipodeformulario),
    [selectedTemplate, state.tipodeformulario]
  );

  const [allowedRegistradores, setAllowedRegistradores] = useState<
    IPersonaProps[]
  >([]);

  const classes = getClasses(cencoTheme);
  const headerStyles: IStackStyles = { root: { marginBottom: 8 } };

  // ✅ helper render label con icono
  const HelpLabel = (p: { text: string; imgUrl?: string }): JSX.Element => (
    <div className={classes.fieldLabelRow}>
      <Label styles={{ root: { marginBottom: 0 } }}>{p.text}</Label>
      <IconButton
        className={classes.helpIcon}
        iconProps={{ iconName: "Info" }}
        title="Ayuda"
        ariaLabel={`Ayuda ${p.text}`}
        onClick={() => openHelp(p.text, p.imgUrl)}
      />
    </div>
  );

  const buttonStyles: IButtonStyles = useMemo(
    () => ({
      root: {
        height: 48,
        minHeight: 48,
        padding: "0 18px",
        borderRadius: 12,
        fontSize: 18,
      },
      flexContainer: { height: 48 },
      label: {
        fontSize: 18,
        fontWeight: 700,
        lineHeight: "1 !important",
      },
      icon: { fontSize: 18 },
    }),
    []
  );

  const calcVigencia = (desdeYmd: string, hastaYmd: string): string => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(desdeYmd || "")) return "";
    if (!/^\d{4}-\d{2}-\d{2}$/.test(hastaYmd || "")) return "";

    const [y1, m1, d1] = desdeYmd.split("-").map(Number);
    const [y2, m2, d2] = hastaYmd.split("-").map(Number);

    const start = new Date(y1, m1 - 1, d1);
    const end = new Date(y2, m2 - 1, d2);

    if (isNaN(start.getTime()) || isNaN(end.getTime())) return "";
    if (end < start) return "";

    let totalMonths = (y2 - y1) * 12 + (m2 - m1);
    if (d2 < d1) totalMonths--;
    totalMonths = Math.max(0, totalMonths);

    if (totalMonths < 12)
      return `${totalMonths} ${totalMonths === 1 ? "mes" : "meses"}`;

    const years = Math.floor(totalMonths / 12);
    const months = totalMonths % 12;

    const yearsTxt = `${years} ${years === 1 ? "año" : "años"}`;
    const monthsTxt =
      months > 0 ? ` ${months} ${months === 1 ? "mes" : "meses"}` : "";
    return yearsTxt + monthsTxt;
  };

  /* ===== Cargar Tipos ===== */
  useEffect(() => {
    const loadTipos = async (): Promise<void> => {
      try {
        setLoadingTipos(true);
        setErrorTipos(undefined);

        const data = await sp.getTiposFormulario();
        setTipos(
          data
            .map((t) => ({
              Id: Number(t.Id),
              Title: String(t.Title ?? ""),
              orden: Number((t as any).orden ?? 0),
              template: String((t as any).template ?? (t as any).Template ?? "")
                .toUpperCase()
                .trim(),
            }))
            .sort((a, b) => (a.orden ?? 0) - (b.orden ?? 0))
        );
      } catch (e: unknown) {
        setErrorTipos(
          e instanceof Error ? e.message : "Error cargando Tipo formulario."
        );
      } finally {
        setLoadingTipos(false);
      }
    };
    void loadTipos();
  }, [sp]);

  /* ===== Cargar opciones del lookup 'proveedor' ===== */
  useEffect(() => {
    const loadProveedorOptions = async (): Promise<void> => {
      try {
        const fieldUrl =
          `${props.siteUrl}/_api/web/lists/getbytitle('Formularios')/fields` +
          `/getbyinternalnameortitle('proveedor')?$select=Title,InternalName,LookupList,LookupField,SchemaXml`;

        const fieldRes = await props.spHttpClient.get(
          fieldUrl,
          SPHttpClient.configurations.v1
        );
        if (!fieldRes.ok) throw new Error(await fieldRes.text());

        const fj: any = await fieldRes.json();
        const f: any = fj?.d ?? fj;

        let lookupListId: string | undefined =
          typeof f?.LookupList === "string" ? f.LookupList : undefined;

        let showField: string =
          typeof f?.LookupField === "string" && f.LookupField
            ? f.LookupField
            : "Title";

        let lookupWebId: string | undefined;

        if (typeof f?.SchemaXml === "string") {
          const schema: string = f.SchemaXml;

          if (!lookupListId) {
            const mList = schema.match(/LookupList="{?([0-9a-fA-F-]{36})}?"/);
            if (mList?.[1]) lookupListId = mList[1];
          }

          const mShow = schema.match(/ShowField="([^"]+)"/);
          if (mShow?.[1]) showField = mShow[1];

          const mWeb = schema.match(/LookupWebId="{?([0-9a-fA-F-]{36})}?"/);
          if (mWeb?.[1]) lookupWebId = mWeb[1];
        }

        if (!lookupListId) {
          setProveedorOptions([]);
          return;
        }

        lookupListId = lookupListId.replace(/[{}]/g, "");
        if (lookupWebId) lookupWebId = lookupWebId.replace(/[{}]/g, "");

        const apiRoot = lookupWebId
          ? `${props.siteUrl}/_api/site/openWebById('${lookupWebId}')/web`
          : `${props.siteUrl}/_api/web`;

        const itemsUrl =
          `${apiRoot}/lists(guid'${lookupListId}')/items` +
          `?$select=Id,${showField}&$orderby=${showField} asc`;

        const itemsRes = await props.spHttpClient.get(
          itemsUrl,
          SPHttpClient.configurations.v1
        );
        if (!itemsRes.ok) throw new Error(await itemsRes.text());

        const itemsJson: any = await itemsRes.json();
        const rows: any[] =
          itemsJson?.d?.results ||
          itemsJson?.value ||
          (Array.isArray(itemsJson) ? itemsJson : []);

        const opts: IDropdownOption[] = rows
          .map((r) => ({ key: Number(r.Id), text: String(r[showField] ?? "") }))
          .filter((o) => !!o.text);

        setProveedorOptions(opts);
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn("Lookup proveedor falló:", err);
        setProveedorOptions([]);
      }
    };

    void loadProveedorOptions();
  }, [props.siteUrl, props.spHttpClient]);

  /* ===== Rehidratar proveedor/ruc/fecha/registradores desde "Proveedores" (solo proveedor=true) ===== */
  useEffect(() => {
    const loadProveedorActual = async (): Promise<void> => {
      if (!props.proveedor) return;

      try {
        const currentUserRes = await props.spHttpClient.get(
          `${props.siteUrl}/_api/web/currentuser`,
          SPHttpClient.configurations.v1
        );
        if (!currentUserRes.ok) return;

        const cuJson: any = await currentUserRes.json();
        const cu: any = cuJson?.d ?? cuJson;
        const userId: number | undefined = cu?.Id;
        if (!userId) return;

        const provRes = await props.spHttpClient.get(
          `${props.siteUrl}/_api/web/lists/getbytitle('Proveedores')/items` +
            `?$select=Id,RUC,Title,Usuarios/Id,` +
            `usuarioregistrador/Id,usuarioregistrador/Title,usuarioregistrador/EMail` +
            `&$expand=Usuarios,usuarioregistrador` +
            `&$filter=Usuarios/Id eq ${userId}`,
          SPHttpClient.configurations.v1
        );
        if (!provRes.ok) return;

        const provJson: any = await provRes.json();
        const rows: any[] = provJson?.d?.results || provJson?.value || [];
        const prov = rows[0];
        if (!prov) {
          setAllowedRegistradores([]);
          return;
        }

        const rawUR = prov.usuarioregistrador;
        const regsRaw: any[] = [];
        if (rawUR) {
          if (Array.isArray(rawUR)) regsRaw.push(...rawUR);
          else if (Array.isArray(rawUR.results)) regsRaw.push(...rawUR.results);
          else regsRaw.push(rawUR);
        }

        const regsPersonas: IPersonaProps[] = regsRaw
          .map(
            (u: any): IPersonaProps => ({
              text: u.Title,
              secondaryText: u.EMail || u.Email || "",
              tertiaryText: `ID: ${u.Id}`,
              id: String(u.Id),
            })
          )
          .filter((p) => !!p.text);

        setAllowedRegistradores(regsPersonas);

        const today = getTodayYMD();

        setState((s) => ({
          ...s,
          fechaderegistro: s.fechaderegistro || today,
          ruc: prov.RUC || s.ruc || "",
          proveedorId: typeof prov.Id === "number" ? prov.Id : s.proveedorId,
          error: undefined,
        }));
      } catch (err) {
        // eslint-disable-next-line no-console
        console.warn("Error cargando proveedor actual:", err);
        setAllowedRegistradores([]);
      }
    };

    void loadProveedorActual();
  }, [props.proveedor, props.siteUrl, props.spHttpClient, proveedorRefreshKey]);

  /* ===== calcular vigencia ===== */
  useEffect(() => {
    const computed = calcVigencia(state.periododesde, state.periodohasta);
    if ((state.anio || "") !== computed)
      setState((s) => ({ ...s, anio: computed }));
  }, [state.periododesde, state.periodohasta]);

  const registradorOptions: IDropdownOption[] = useMemo(
    () =>
      allowedRegistradores
        .map((p) => ({ key: Number(p.id), text: p.text || "" }))
        .filter((o) => !!o.text && !isNaN(o.key as number)),
    [allowedRegistradores]
  );

  const resolvePeople = async (filter: string): Promise<IPersonaProps[]> => {
    const q = (filter || "").trim();
    if (!q) return [];

    const url =
      `${props.siteUrl}/_api/web/siteusers` +
      `?$select=Id,Title,Email&$filter=startswith(Title,'${encodeURIComponent(
        q
      )}') or startswith(Email,'${encodeURIComponent(q)}')`;

    try {
      const res = await props.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );
      if (!res.ok) return [];
      const j: any = await res.json();
      const rows: any[] = j?.d?.results || j?.value || [];
      return rows.map(
        (u: any): IPersonaProps => ({
          text: u.Title,
          secondaryText: u.Email,
          tertiaryText: `ID: ${u.Id}`,
          id: String(u.Id),
        })
      );
    } catch {
      return [];
    }
  };

  const setField = (key: keyof IFormState, value: any): void => {
    setState((s) => ({
      ...s,
      [key]: value,
      error: undefined,
      success: undefined,
    }));
  };

  /* ===== Helpers REST (lookup/user) ===== */
  const ensureIdField = (internal?: string): string | undefined => {
    const v = (internal || "").trim();
    if (!v) return undefined;
    return /id$/i.test(v) ? v : `${v}Id`;
  };
  const asResults = (ids: number[]) => ({ results: ids });

  /* ===== canSubmit (incluye validaciones) ===== */
  const canSubmit = useMemo(() => {
    const visibleSections = visibleSectionsFor(
      selectedTemplate,
      state.tipodeformulario
    );

    const needDates =
      visibleSections.has(3) && !isSctrExcel(state.tipodeformulario);
    const needFiles = visibleSections.has(4);

    const desdeStr = (state.periododesde || "").trim();
    const hastaStr = (state.periodohasta || "").trim();

    const desdeOk = !needDates || /^\d{4}-\d{2}-\d{2}$/.test(desdeStr);
    const hastaOk = !needDates || /^\d{4}-\d{2}-\d{2}$/.test(hastaStr);

    const today = getTodayYMD();
    const rangoOk =
      !needDates || (!desdeOk || !hastaOk ? false : desdeStr <= hastaStr);
    const noVencidaOk = !needDates || (!hastaOk ? false : hastaStr >= today);

    const needCodigos = isSctrExcel(state.tipodeformulario);
    const codigosOk = !needCodigos
      ? true
      : (state.codigosDocumentos || []).length > 0 &&
        (state.codigosDocumentos || []).every(
          (r) => (r.descripcion || "").trim() && (r.codigo || "").trim()
        );

    const filesOk = !needFiles || (state.archivos && state.archivos.length > 0);

    return (
      !!selectedTemplate &&
      desdeOk &&
      hastaOk &&
      rangoOk &&
      noVencidaOk &&
      codigosOk &&
      filesOk &&
      !state.isSaving &&
      !submitLockedRef.current
    );
  }, [
    selectedTemplate,
    state.tipodeformulario,
    state.periododesde,
    state.periodohasta,
    state.codigosDocumentos,
    state.archivos,
    state.isSaving,
  ]);

  /* ===== Files: SIN Array.from (compat TS viejo) ===== */
  const onPickFiles = (files: FileList | null): void => {
    if (!files || files.length === 0) return;

    const picked: File[] = [];
    for (let i = 0; i < files.length; i++) {
      const f = files.item(i);
      if (f) picked.push(f);
    }

    setState((s) => ({
      ...s,
      archivos: [...(s.archivos || []), ...picked],
      error: undefined,
      success: undefined,
    }));
  };

  const removeFileAt = (idx: number): void => {
    setState((s) => {
      const next = (s.archivos || []).filter((_, i) => i !== idx);
      return { ...s, archivos: next, error: undefined, success: undefined };
    });
  };

  /* ===== Códigos SCTR ===== */
  const setCodigoRow = (
    idx: number,
    key: keyof ICodigoDocumentoRow,
    value: string
  ): void => {
    setState((s) => {
      const arr = (s.codigosDocumentos || []).slice();
      const row = arr[idx] || { codigo: "", descripcion: "" };
      arr[idx] = { ...row, [key]: value };
      return {
        ...s,
        codigosDocumentos: arr,
        error: undefined,
        success: undefined,
      };
    });
  };

  const addCodigoRow = (): void => {
    setState((s) => ({
      ...s,
      codigosDocumentos: [
        ...(s.codigosDocumentos || []),
        { codigo: "", descripcion: "" },
      ],
      error: undefined,
      success: undefined,
    }));
  };

  const removeCodigoRow = (idx: number): void => {
    setState((s) => {
      const arr = (s.codigosDocumentos || []).filter((_, i) => i !== idx);
      return {
        ...s,
        codigosDocumentos: arr.length ? arr : [{ codigo: "", descripcion: "" }],
        error: undefined,
        success: undefined,
      };
    });
  };

  // --- onSave igual que el tuyo (sin cambios) ---
  // (por espacio, lo dejé igual; no afecta lo de UI)

  const onSave = async (): Promise<void> => {
    if (!acquireLock()) return;

    const ids = (state.usuarioregistradorIds || []).filter(
      (n) => typeof n === "number" && !isNaN(n)
    );
    if (ids.length === 0) {
      releaseLock();
      setState((s) => ({
        ...s,
        error: "Debés seleccionar al menos un Usuario registrador.",
        success: undefined,
      }));
      return;
    }

    try {
      setState((s) => ({
        ...s,
        isSaving: true,
        error: undefined,
        success: undefined,
      }));

      const visibleSections = visibleSectionsFor(
        selectedTemplate,
        state.tipodeformulario
      );

      const needDates =
        visibleSections.has(3) && !isSctrExcel(state.tipodeformulario);
      if (needDates) {
        const desdeStr = (state.periododesde || "").trim();
        const hastaStr = (state.periodohasta || "").trim();

        const desdeOk = /^\d{4}-\d{2}-\d{2}$/.test(desdeStr);
        const hastaOk = /^\d{4}-\d{2}-\d{2}$/.test(hastaStr);

        if (!desdeOk || !hastaOk)
          throw new Error("Completá los campos De y A para poder guardar.");
        if (desdeStr > hastaStr)
          throw new Error("La fecha 'De' no puede ser mayor que la fecha 'A'.");
        if (hastaStr < getTodayYMD())
          throw new Error(
            "La fecha 'A' no puede estar vencida (debe ser hoy o posterior)."
          );
      }

      if (isSctrExcel(state.tipodeformulario)) {
        const rows = state.codigosDocumentos || [];
        const anyEmpty =
          rows.length === 0 ||
          rows.some(
            (r) => !(r.descripcion || "").trim() || !(r.codigo || "").trim()
          );
        if (anyEmpty) {
          throw new Error(
            "Completá todos los campos de 'Código de documentos' antes de guardar."
          );
        }
      }

      if (
        visibleSections.has(4) &&
        (!state.archivos || state.archivos.length === 0)
      ) {
        throw new Error(
          "Debés adjuntar al menos un documento para poder guardar."
        );
      }

      const trim = (v?: string) => (v ? v.trim() : "");
      const mapField = (
        mapped?: string,
        fallback?: string
      ): string | undefined => {
        const m = (mapped || "").trim();
        if (m) return m;
        const f = (fallback || "").trim();
        return f ? f : undefined;
      };

      const srcIds = state.usuarioregistradorIds || [];
      const uniqIds: number[] = [];
      for (let i = 0; i < srcIds.length; i++) {
        const n = srcIds[i];
        if (typeof n === "number" && !isNaN(n) && uniqIds.indexOf(n) === -1)
          uniqIds.push(n);
      }

      const codigosLines = (state.codigosDocumentos || [])
        .map((r) => {
          const desc = (r.descripcion || "").trim();
          const cod = (r.codigo || "").trim();
          if (!desc && !cod) return "";
          return `${desc} : ${cod}`.trim();
        })
        .filter((x) => !!x);

      const codigosTxt = codigosLines
        .map((line, i) => (i === 0 ? line : `| ${line}`))
        .join("\r\n");

      const fm = props.fieldMap || {};

      const fTitle = mapField((fm as any).title, "Title");
      const fFecha = mapField((fm as any).fechaderegistro, "fechaderegistro");
      const fRuc = mapField((fm as any).ruc, "ruc");
      const fProveedor = ensureIdField(
        mapField((fm as any).proveedor, "proveedor")
      );
      const fUser = ensureIdField(
        mapField((fm as any).usuarioregistrador, "usuarioregistrador")
      );
      const fCodContrato = mapField(
        (fm as any).codigodecontrato,
        "codigodecontrato"
      );
      const fDesde = mapField((fm as any).periododesde, "periododesde");
      const fHasta = mapField((fm as any).periodohasta, "periodohasta");
      const fAnio = mapField((fm as any).anio, "a_x00f1_o");
      const fCodDocs = mapField(
        (fm as any).codigodedocumentos,
        "codigodedocumentos"
      );

      const bodyAll: Record<string, any> = {
        ...(fTitle ? { [fTitle]: trim(state.Title) || undefined } : {}),
        ...(fFecha ? { [fFecha]: toSpDate(state.fechaderegistro) } : {}),
        ...(fRuc ? { [fRuc]: trim(state.ruc) || undefined } : {}),

        ...(fCodContrato
          ? { [fCodContrato]: trim(state.codigodecontrato) || undefined }
          : {}),
        ...(fDesde ? { [fDesde]: toSpDate(state.periododesde) } : {}),
        ...(fHasta ? { [fHasta]: toSpDate(state.periodohasta) } : {}),
        ...(fAnio ? { [fAnio]: trim(state.anio) || undefined } : {}),

        ...(fProveedor &&
        typeof state.proveedorId === "number" &&
        !isNaN(state.proveedorId)
          ? { [fProveedor]: state.proveedorId }
          : {}),

        ...(fUser && uniqIds.length ? { [fUser]: asResults(uniqIds) } : {}),

        tipodeformulario: trim(state.tipodeformulario) || undefined,

        ...(isSctrExcel(state.tipodeformulario) && fCodDocs
          ? { [fCodDocs]: codigosTxt || undefined }
          : {}),
      };

      const body: Record<string, any> = {};
      for (const k in bodyAll) {
        if (!Object.prototype.hasOwnProperty.call(bodyAll, k)) continue;
        const v = (bodyAll as any)[k];
        if (v !== undefined && v !== null && v !== "") body[k] = v;
      }

      const archivosParaSubir = visibleSections.has(4) ? state.archivos : [];
      await sp.createFormulario(body, archivosParaSubir);

      setState((prev) => ({
        ...initialState,
        fechaderegistro: prev.fechaderegistro || getTodayYMD(),
        ruc: prev.ruc,
        proveedorId: prev.proveedorId,
        tipodeformulario: prev.tipodeformulario,
        success: "Guardado correctamente.",
      }));

      setPeopleSelected([]);
      refreshProveedor();
    } catch (e: unknown) {
      setState((s) => ({
        ...s,
        error: e instanceof Error ? e.message : "Error desconocido al guardar.",
      }));
    } finally {
      setState((s) => ({ ...s, isSaving: false }));
      releaseLock();
    }
  };

  const navGroups: INavLinkGroup[] = useMemo(() => {
    const links: INavLink[] = tipos.map((t) => ({
      key: String(t.Id),
      name: t.Title,
      url: "#",
    }));
    return [{ links }];
  }, [tipos]);

  const handleNavClick = (
    ev?: React.MouseEvent<HTMLElement>,
    item?: INavLink
  ): void => {
    ev?.preventDefault();
    if (!item) return;

    const id = Number(item.key);

    let match: ITipoFormularioItem | undefined;
    for (let i = 0; i < tipos.length; i++) {
      if (tipos[i].Id === id) {
        match = tipos[i];
        break;
      }
    }

    const tipo = match?.Title || "";

    setSelectedId(id);
    setSelectedTemplate(match?.template);

    setState((s) => ({
      ...s,
      tipodeformulario: tipo,
      error: undefined,
      success: undefined,
      codigosDocumentos: isSctrExcel(tipo)
        ? s.codigosDocumentos?.length
          ? s.codigosDocumentos
          : [{ codigo: "", descripcion: "" }]
        : [{ codigo: "", descripcion: "" }],
    }));
  };

  return (
    <ThemeProvider theme={cencoTheme}>
      <div className={classes.page}>
        <div className={classes.container}>
          <div className={classes.mainRow}>
            {/* izquierda */}
            <div className={classes.leftCol}>
              <div className={classes.navPanel}>
                <div className={classes.navHeader}>
                  <Text
                    variant="mediumPlus"
                    block
                    styles={{
                      root: {
                        fontWeight: FontWeights.semibold,
                        color: cencoTheme.palette.themePrimary,
                      },
                    }}
                  >
                    Tipo de formulario
                  </Text>
                </div>

                <div className={classes.navBody}>
                  {loadingTipos && (
                    <Spinner size={SpinnerSize.small} label="Cargando…" />
                  )}
                  {errorTipos && (
                    <MessageBar
                      messageBarType={MessageBarType.error}
                      isMultiline
                    >
                      {errorTipos}
                    </MessageBar>
                  )}
                  {!loadingTipos && !errorTipos && tipos.length === 0 && (
                    <Text variant="small" styles={{ root: { color: "#666" } }}>
                      Sin opciones.
                    </Text>
                  )}
                  {!loadingTipos && !errorTipos && tipos.length > 0 && (
                    <div className={classes.navResponsive}>
                      <Nav
                        groups={navGroups}
                        onLinkClick={handleNavClick}
                        selectedKey={
                          selectedId ? String(selectedId) : undefined
                        }
                        ariaLabel="Tipos de formulario"
                        styles={{ root: { width: "100%" } }}
                      />
                    </div>
                  )}
                </div>
              </div>
            </div>

            {/* derecha */}
            <div className={classes.rightCol}>
              <Stack
                styles={headerStyles}
                horizontal
                horizontalAlign="space-between"
                verticalAlign="center"
              >
                <Text
                  variant="xLarge"
                  styles={{
                    root: {
                      fontWeight: FontWeights.semibold,
                      color: cencoTheme.palette.themePrimary,
                    },
                  }}
                >
                  Registro de documentos
                </Text>

                {state.isSaving && (
                  <Stack
                    horizontal
                    tokens={{ childrenGap: 6 }}
                    verticalAlign="center"
                  >
                    <Spinner size={SpinnerSize.small} />
                    <Text variant="small">Guardando…</Text>
                  </Stack>
                )}
              </Stack>

              <form
                onSubmit={async (e): Promise<void> => {
                  e.preventDefault();
                  await onSave();
                }}
              >
                {!selectedTemplate && (
                  <MessageBar messageBarType={MessageBarType.info} isMultiline>
                    Elegí un <strong>Tipo de formulario</strong> para mostrar
                    las secciones.
                  </MessageBar>
                )}

                {/* 1 */}
                {visible.has(1) && (
                  <Section title="1.- Identificación" classes={classes}>
                    <div className={classes.row}>
                      <div className={classes.c7}>
                        <TextField
                          label="Fecha de registro"
                          type="date"
                          value={state.fechaderegistro}
                          onChange={(_, v) =>
                            setField("fechaderegistro", v || "")
                          }
                          id="fechaderegistro"
                          disabled={props.proveedor === true}
                        />
                        <input
                          type="hidden"
                          id="tipodeformulario"
                          value={state.tipodeformulario}
                          readOnly
                        />
                      </div>

                      <div className={classes.c5}>
                        <TextField
                          label="RUC"
                          value={state.ruc}
                          onChange={(_, v) => setField("ruc", v || "")}
                          id="ruc"
                          disabled={props.proveedor === true}
                        />
                      </div>

                      <div className={classes.c8}>
                        <Dropdown
                          label="Razón social"
                          placeholder="Seleccioná un proveedor…"
                          options={proveedorOptions}
                          selectedKey={state.proveedorId}
                          onChange={(_, option) =>
                            setField("proveedorId", Number(option?.key))
                          }
                          id="proveedor"
                          disabled={props.proveedor === true}
                        />
                      </div>

                      <div className={classes.c6}>
                        <Label required>Usuario registrador</Label>

                        {props.proveedor ? (
                          <Dropdown
                            placeholder="Seleccioná usuario(s)…"
                            multiSelect
                            options={registradorOptions}
                            selectedKeys={state.usuarioregistradorIds}
                            required
                            onChange={(_, option) => {
                              if (!option) return;
                              const idNum = Number(option.key);

                              setState((s) => {
                                let ids = s.usuarioregistradorIds || [];
                                if (option.selected) {
                                  if (ids.indexOf(idNum) === -1)
                                    ids = [...ids, idNum];
                                } else {
                                  ids = ids.filter((x) => x !== idNum);
                                }
                                return {
                                  ...s,
                                  usuarioregistradorIds: ids,
                                  error: undefined,
                                  success: undefined,
                                };
                              });
                            }}
                            disabled={!registradorOptions.length}
                          />
                        ) : (
                          <NormalPeoplePicker
                            onResolveSuggestions={resolvePeople}
                            getTextFromItem={(p: IPersonaProps) => p.text || ""}
                            pickerSuggestionsProps={suggestionProps}
                            selectedItems={peopleSelected}
                            onChange={(items?: IPersonaProps[]) => {
                              const arr = items || [];
                              setPeopleSelected(arr);
                              const ids = arr
                                .map((p) => Number(p.id))
                                .filter((n) => !isNaN(n));
                              setField("usuarioregistradorIds", ids);
                            }}
                            inputProps={{
                              "aria-label": "Buscar usuarios",
                              placeholder: "Escribí para buscar usuarios…",
                              "aria-required": true,
                            }}
                            resolveDelay={300}
                          />
                        )}
                      </div>
                    </div>
                  </Section>
                )}

                {/* 2 */}
                {visible.has(2) && (
                  <Section
                    title={
                      isCargarContratos(state.tipodeformulario)
                        ? "2.- Datos del contrato"
                        : "2.- Datos generales"
                    }
                    classes={classes}
                  >
                    {isCargarContratos(state.tipodeformulario) ? (
                      <div className={classes.row}>
                        <div className={classes.c12}>
                          <TextField
                            label=""
                            onRenderLabel={() =>
                              HelpLabel({
                                text: "Nombre del contrato",
                                imgUrl: props.helpImgNombreContrato,
                              })
                            }
                            value={state.Title}
                            onChange={(_, v) => setField("Title", v || "")}
                          />
                        </div>

                        <div className={classes.c6}>
                          <TextField
                            label=""
                            onRenderLabel={() =>
                              HelpLabel({
                                text: "Código de contrato",
                                imgUrl: props.helpImgCodigoContrato,
                              })
                            }
                            value={state.codigodecontrato}
                            onChange={(_, v) =>
                              setField("codigodecontrato", v || "")
                            }
                          />
                        </div>

                        {/* ✅ título centrado */}
                        <div className={classes.plazoHeader}>
                          Plazo de contrato
                        </div>

                        {/* ✅ misma fila */}
                        <div className={classes.c6}>
                          <TextField
                            label=""
                            onRenderLabel={() =>
                              HelpLabel({
                                text: "Fecha de inicio",
                                imgUrl: props.helpImgFechaInicio,
                              })
                            }
                            type="date"
                            value={state.periododesde}
                            onChange={(_, v) =>
                              setField("periododesde", v || "")
                            }
                          />
                        </div>

                        <div className={classes.c6}>
                          <TextField
                            label=""
                            onRenderLabel={() =>
                              HelpLabel({
                                text: "Fecha de fin",
                                imgUrl: props.helpImgFechaFin,
                              })
                            }
                            type="date"
                            value={state.periodohasta}
                            onChange={(_, v) =>
                              setField("periodohasta", v || "")
                            }
                          />
                        </div>
                      </div>
                    ) : (
                      <div className={classes.row}>
                        <div className={classes.c6}>
                          <TextField
                            label="De"
                            type="date"
                            value={state.periododesde}
                            onChange={(_, v) =>
                              setField("periododesde", v || "")
                            }
                          />
                        </div>

                        <div className={classes.c6}>
                          <TextField
                            label="A"
                            type="date"
                            value={state.periodohasta}
                            onChange={(_, v) =>
                              setField("periodohasta", v || "")
                            }
                          />
                        </div>

                        <div className={classes.c6}>
                          <TextField
                            label="Año"
                            value={state.anio}
                            readOnly
                            disabled
                          />
                        </div>
                      </div>
                    )}
                  </Section>
                )}

                {/* 3 */}
                {visible.has(3) && (
                  <Section
                    title={
                      isSctrExcel(state.tipodeformulario)
                        ? "3.- Código de documentos"
                        : "3.- Datos generales"
                    }
                    classes={classes}
                  >
                    {isSctrExcel(state.tipodeformulario) ? (
                      <>
                        {(state.codigosDocumentos || []).map((row, idx) => (
                          <div
                            key={idx}
                            className={classes.codigosRow}
                            style={{ marginTop: idx ? 10 : 0 }}
                          >
                            <TextField
                              label={idx === 0 ? "Descripción" : undefined}
                              value={row.descripcion}
                              onChange={(_, v) =>
                                setCodigoRow(idx, "descripcion", v || "")
                              }
                            />
                            <TextField
                              label={idx === 0 ? "Código" : undefined}
                              value={row.codigo}
                              onChange={(_, v) =>
                                setCodigoRow(idx, "codigo", v || "")
                              }
                            />
                            <div className={classes.codigosDelete}>
                              <IconButton
                                iconProps={{ iconName: "Delete" }}
                                title="Eliminar fila"
                                ariaLabel="Eliminar fila"
                                onClick={() => removeCodigoRow(idx)}
                              />
                            </div>
                          </div>
                        ))}
                        <div style={{ marginTop: 10 }}>
                          <DefaultButton
                            text="Agregar fila"
                            iconProps={{ iconName: "Add" }}
                            onClick={addCodigoRow}
                            styles={buttonStyles}
                          />
                        </div>
                      </>
                    ) : (
                      <div className={classes.row}>
                        <div className={classes.c6}>
                          <TextField
                            label="Periodo desde"
                            type="date"
                            value={state.periododesde}
                            onChange={(_, v) =>
                              setField("periododesde", v || "")
                            }
                          />
                        </div>

                        <div className={classes.c6}>
                          <TextField
                            label="Periodo hasta"
                            type="date"
                            value={state.periodohasta}
                            onChange={(_, v) =>
                              setField("periodohasta", v || "")
                            }
                          />
                        </div>
                        <div className={classes.c6}>
                          <TextField
                            label="Tiempo de vigencia"
                            value={state.anio}
                            readOnly
                            disabled
                          />
                        </div>
                      </div>
                    )}
                  </Section>
                )}

                {/* 4 */}
                {visible.has(4) && (
                  <Section title="4.- Cargar documento" classes={classes}>
                    <div className={classes.row}>
                      <div className={classes.c12}>
                        <Label>Adjuntar archivos</Label>
                        <input
                          type="file"
                          multiple
                          onChange={(e) => {
                            onPickFiles(e.currentTarget.files);
                            e.currentTarget.value = "";
                          }}
                        />

                        {(state.archivos || []).map((f, idx) => (
                          <div
                            key={`${f.name}-${idx}`}
                            className={classes.fileRow}
                          >
                            <div style={{ minWidth: 0 }}>
                              <Text
                                styles={{ root: { fontWeight: 600 } }}
                                block
                              >
                                {f.name}
                              </Text>
                              <Text
                                variant="small"
                                styles={{ root: { color: "#666" } }}
                                block
                              >
                                {(f.size / 1024).toFixed(1)} KB
                              </Text>
                            </div>
                            <IconButton
                              iconProps={{ iconName: "Cancel" }}
                              title="Quitar"
                              ariaLabel="Quitar"
                              onClick={() => removeFileAt(idx)}
                            />
                          </div>
                        ))}
                      </div>
                    </div>
                  </Section>
                )}

                {/* acciones */}
                <div className={classes.actionsBar}>
                  <PrimaryButton
                    type="submit"
                    text={state.isSaving ? "Guardando…" : "Guardar"}
                    disabled={!canSubmit}
                    iconProps={
                      state.isSaving
                        ? { iconName: "Sync" }
                        : { iconName: "Save" }
                    }
                    styles={buttonStyles}
                  />
                  <DefaultButton
                    type="button"
                    text="Limpiar"
                    disabled={state.isSaving || submitLockedRef.current}
                    onClick={() => {
                      setState((prev) => ({
                        ...initialState,
                        fechaderegistro: props.proveedor
                          ? prev.fechaderegistro
                          : "",
                        ruc: props.proveedor ? prev.ruc : "",
                        proveedorId: props.proveedor
                          ? prev.proveedorId
                          : undefined,
                        tipodeformulario: prev.tipodeformulario,
                      }));
                      setPeopleSelected([]);
                      refreshProveedor();
                    }}
                    iconProps={{ iconName: "Clear" }}
                    styles={buttonStyles}
                  />
                </div>

                {state.error && (
                  <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline
                    styles={{ root: { marginTop: 12 } }}
                  >
                    {state.error}
                  </MessageBar>
                )}
                {state.success && (
                  <MessageBar
                    messageBarType={MessageBarType.success}
                    isMultiline
                    styles={{ root: { marginTop: 12 } }}
                  >
                    {state.success}
                  </MessageBar>
                )}
              </form>
            </div>
          </div>
        </div>
      </div>

      {/* ✅ NUEVO: Modal de ayuda */}
      <Modal isOpen={helpOpen} onDismiss={closeHelp} isBlocking={false}>
        <div className={classes.modalBody}>
          <Stack
            horizontal
            horizontalAlign="space-between"
            verticalAlign="center"
          >
            <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
              {helpTitle}
            </Text>
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              title="Cerrar"
              ariaLabel="Cerrar"
              onClick={closeHelp}
            />
          </Stack>

          <div style={{ marginTop: 12 }}>
            <img src={helpImg} className={classes.modalImg} alt={helpTitle} />
          </div>
        </div>
      </Modal>
    </ThemeProvider>
  );
}
