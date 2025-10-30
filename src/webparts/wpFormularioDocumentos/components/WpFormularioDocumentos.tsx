import * as React from "react";
import { useEffect, useMemo, useRef, useState } from "react";
import { IWpFormularioDocumentosProps } from "./IWpFormularioDocumentosProps";
import { SpService } from "../services/sp.service";

/* ==== SharePoint REST ==== */
import { SPHttpClient } from "@microsoft/sp-http";

/* ==== Fluent UI ==== */
import {
  ThemeProvider,
  Text,
  TooltipHost,
  Label,
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Separator,
  Nav,
  INavLinkGroup,
  INavLink,
  getTheme,
  mergeStyleSets,
  FontWeights,
  Stack,
  IStackTokens,
  IStackStyles,
  Dropdown,
  IDropdownOption,
  IconButton,
  Dialog,
  DialogType,
  DialogFooter,
  Image,
  ImageFit,
  ITextFieldProps,
} from "@fluentui/react";

/* People Picker */
import {
  NormalPeoplePicker,
  IBasePickerSuggestionsProps,
} from "@fluentui/react/lib/Pickers";
import { IPersonaProps } from "@fluentui/react/lib/Persona";

/* ===== Tipos ===== */
interface IFormState {
  fechaderegistro: string; // DateOnly (input yyyy-mm-dd)
  ruc: string; // Text
  proveedorId?: number; // Lookup single -> proveedorId
  usuarioregistradorIds: number[]; // UserMulti -> usuarioregistradorId: []
  Title: string; // Text
  codigodecontrato: string; // Text
  periododesde: string; // DateOnly
  periodohasta: string; // DateOnly
  anio: string; // Año (internal a_x00f1_o)
  archivo?: File;
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
  archivo: undefined,
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
  // Mandamos UTC a medianoche
  return `${v}T00:00:00Z`;
};

/* ---- Estilos Fluent (responsive + Microsoft look) ---- */
const getClasses = (theme = getTheme()) =>
  mergeStyleSets({
    root: {
      padding: 12,
      [`@media (min-width: 640px)`]: { padding: 16 },
      [`@media (min-width: 1024px)`]: { padding: 20 },
    },

    /* layout principal */
    mainRow: {
      display: "flex",
      alignItems: "flex-start",
      gap: 16,
      flexWrap: "nowrap",
    },
    leftCol: {
      flex: "0 0 240px",
      width: 240,
      position: "sticky",
      top: 12,
      [`@media (max-width: 520px)`]: {
        position: "static",
        width: "100%",
        flex: "1 1 100%",
        marginBottom: 8,
      },
    },
    rightCol: {
      flex: "1 1 0",
      minWidth: 0,
    },

    /* panel */
    panel: {
      background: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 12,
      boxShadow: theme.effects.elevation8,
      overflow: "hidden",
      animation: "fadeIn .25s ease-out both",
    },
    panelHeader: {
      background: theme.palette.neutralLighterAlt,
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      padding: "12px 16px",
    },
    panelBody: {
      padding: 16,
      [`@media (min-width: 1024px)`]: { padding: 20 },
    },

    /* encabezado del formulario */
    formHeader: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      marginBottom: 12,
    },

    /* grid 12 cols */
    row: {
      display: "grid",
      gridTemplateColumns: "repeat(12, 1fr)",
      gap: 12,
      [`@media (min-width: 1024px)`]: { gap: 16 },
    },
    c12: { gridColumn: "span 12" },
    c8: {
      gridColumn: "span 12",
      [`@media (min-width: 1024px)`]: { gridColumn: "span 8" },
    },
    c6: {
      gridColumn: "span 12",
      [`@media (min-width: 768px)`]: { gridColumn: "span 6" },
    },
    c4: {
      gridColumn: "span 12",
      [`@media (min-width: 1024px)`]: { gridColumn: "span 4" },
    },

    /* nav panel */
    navPanel: {
      background: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 12,
      boxShadow: theme.effects.elevation8,
    },
    navHeader: {
      padding: "12px 16px",
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      background: theme.palette.neutralLighterAlt,
      borderTopLeftRadius: 12,
      borderTopRightRadius: 12,
    },
    navBody: { padding: 12 },

    /* ajustes para Nav */
    navResponsive: {
      maxHeight: "calc(100vh - 180px)",
      overflowY: "auto",

      ".ms-Nav-link": {
        display: "block !important",
        height: "auto !important",
        minHeight: 40,
        padding: "clamp(6px, 0.8vw, 10px) clamp(10px, 1vw, 14px)",
        borderRadius: 8,
        border: "1px solid transparent",
        lineHeight: "1.35 !important",
        boxSizing: "border-box",
      },
      ".ms-Nav-linkText": {
        display: "block !important",
        whiteSpace: "normal !important",
        wordBreak: "break-word",
        overflowWrap: "anywhere",
        hyphens: "auto",
        fontWeight: 500,
        fontSize: "clamp(12px, 1.2vw, 16px)",
        lineHeight: "1.35 !important",
      },
      ".ms-Nav-link:hover": {
        background: theme.palette.themeLighterAlt,
        borderColor: theme.palette.themeLighter,
      },
      ".ms-Nav .is-selected > .ms-Nav-link": {
        background: theme.palette.themeLighter,
        borderColor: theme.palette.themePrimary,
        boxShadow: "inset 0 0 0 1px rgba(0,0,0,.06)",
      },
    },

    "@global": {
      "@keyframes fadeIn": {
        from: { opacity: 0, transform: "translateY(6px)" },
        to: { opacity: 1, transform: "none" },
      },
    },
  });

/* ==== Sugerencias del PeoplePicker ==== */
const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: "Usuarios",
  noResultsFoundText: "Sin resultados",
};

/* ---- Componente principal ---- */
export default function WpFormularioDocumentos(
  props: IWpFormularioDocumentosProps
): JSX.Element {
  const [state, setState] = useState<IFormState>(initialState);

  /* ---------- Candado anti doble submit SIN Symbol ---------- */
  const submitLockedRef = useRef<boolean>(false);
  const acquireLock = (): boolean => {
    if (submitLockedRef.current) return false;
    submitLockedRef.current = true;
    return true;
  };
  const releaseLock = (): void => {
    submitLockedRef.current = false;
  };

  // Instancia del servicio (memoizada)
  const sp = useMemo(
    () => new SpService((props as any).spHttpClient, props.siteUrl),
    [props.spHttpClient, props.siteUrl]
  );

  // Tipos (Nav)
  const [tipos, setTipos] = useState<ITipoFormularioItem[]>([]);
  const [loadingTipos, setLoadingTipos] = useState<boolean>(true);
  const [errorTipos, setErrorTipos] = useState<string | undefined>(undefined);

  // Proveedores (lookup)
  const [proveedorOptions, setProveedorOptions] = useState<IDropdownOption[]>(
    []
  );

  // People Picker selección
  const [peopleSelected, setPeopleSelected] = useState<IPersonaProps[]>([]);

  // Selección de template
  const [selectedId, setSelectedId] = useState<number | undefined>(undefined);
  const [selectedTemplate, setSelectedTemplate] = useState<string | undefined>(
    undefined
  );
  const visible = sectionsForTemplate(selectedTemplate);

  const fileRef = useRef<HTMLInputElement>(null);

  /* ====== MODAL DE AYUDA ====== */
  const [helpOpen, setHelpOpen] = useState(false);
  const [helpTitle, setHelpTitle] = useState<string>("");
  const [helpUrl, setHelpUrl] = useState<string>("");

  const openHelp = (title: string, url: string): void => {
    setHelpTitle(title);
    setHelpUrl(url);
    setHelpOpen(true);
  };
  const closeHelp = (): void => setHelpOpen(false);

  const renderLabelConAyuda = (
    label: string,
    onClick: () => void
  ): ITextFieldProps["onRenderLabel"] => {
    return () => (
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <Label styles={{ root: { margin: 0 } }}>{label}</Label>
        <TooltipHost content="Ver ayuda">
          <IconButton
            aria-label={`Ayuda: ${label}`}
            iconProps={{ iconName: "Info" }}
            title="Ver ayuda"
            onClick={onClick}
            styles={{
              root: { height: 24, width: 24 },
              icon: { fontSize: 14 },
            }}
          />
        </TooltipHost>
      </div>
    );
  };

  /* ===== Cargar Tipos (Nav) ===== */
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

  /* ===== Cargar opciones del lookup 'proveedor' (Razón social) ===== */
  useEffect(() => {
    const loadProveedorOptions = async (): Promise<void> => {
      try {
        // 1) Traer definición del campo lookup
        const fieldUrl =
          `${props.siteUrl}/_api/web/lists/getbytitle('Formularios')/fields` +
          `/getbyinternalnameortitle('proveedor')?$select=Title,InternalName,LookupList,LookupField,SchemaXml`;

        // ⚠️ sin headers custom para no perder Accept
        const fieldRes = await props.spHttpClient.get(
          fieldUrl,
          SPHttpClient.configurations.v1
        );
        if (!fieldRes.ok) throw new Error(await fieldRes.text());
        const fj: any = await fieldRes.json();
        const f: any = fj?.d ?? fj;

        // 2) Extraer LookupList, LookupField y (si existe) LookupWebId
        let lookupListId: string | undefined =
          typeof f?.LookupList === "string" ? f.LookupList : undefined;
        let showField: string =
          typeof f?.LookupField === "string" && f.LookupField
            ? f.LookupField
            : "Title";
        let lookupWebId: string | undefined = undefined;

        // Fallback a SchemaXml para lo que falte
        if (typeof f?.SchemaXml === "string") {
          const schema: string = f.SchemaXml;
          if (!lookupListId) {
            const mList = schema.match(/LookupList="{?([0-9a-fA-F-]{36})}?"/);
            if (mList && mList[1]) lookupListId = mList[1];
          }
          const mShow = schema.match(/ShowField="([^"]+)"/);
          if (!showField && mShow && mShow[1]) showField = mShow[1];
          const mWeb = schema.match(/LookupWebId="{?([0-9a-fA-F-]{36})}?"/);
          if (mWeb && mWeb[1]) lookupWebId = mWeb[1];
        }

        if (!lookupListId) {
          setProveedorOptions([]);
          return;
        }

        // Normalizar GUIDs (sin llaves)
        lookupListId = lookupListId.replace(/[{}]/g, "");
        if (lookupWebId) lookupWebId = lookupWebId.replace(/[{}]/g, "");

        // 3) Construir raíz correcta del API según LookupWebId
        const apiRoot = lookupWebId
          ? `${props.siteUrl}/_api/site/openWebById('${lookupWebId}')/web`
          : `${props.siteUrl}/_api/web`;

        // 4) Traer ítems del lookup
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

        // 5) Mapear a opciones
        const opts: IDropdownOption[] = rows
          .map((r) => ({ key: Number(r.Id), text: String(r[showField] ?? "") }))
          .filter((o) => !!o.text);

        setProveedorOptions(opts);
      } catch (err) {
        console.warn("Lookup proveedor falló:", err);
        setProveedorOptions([]);
      }
    };

    void loadProveedorOptions();
  }, [props.siteUrl, props.spHttpClient]);

  /* ===== People Picker: búsqueda de usuarios (siteusers) ===== */
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
        (u): IPersonaProps => ({
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

  /* ===== Helpers de estado ===== */
  const setField = (key: keyof IFormState, value: any): void => {
    setState((s) => ({
      ...s,
      [key]: value,
      error: undefined,
      success: undefined,
    }));
  };

  const onSave = async (): Promise<void> => {
    // Candado anti doble submit
    if (!acquireLock()) return;

    try {
      setState((s) => ({
        ...s,
        isSaving: true,
        error: undefined,
        success: undefined,
      }));

      // --- Secciones visibles según la botonera/template ---
      const visibleSections = sectionsForTemplate(selectedTemplate);

      // --- Lista blanca por sección ---
      const allowed = [];
      if (visibleSections.has(1)) {
        allowed.push(
          "fechaderegistro",
          "ruc",
          "proveedorId",
          "usuarioregistradorId",
          "tipodeformulario"
        );
      }
      if (visibleSections.has(2)) {
        allowed.push(
          "Title",
          "codigodecontrato",
          "periododesde",
          "periodohasta"
        );
      }
      if (visibleSections.has(3)) {
        allowed.push("periododesde", "periodohasta", "a_x00f1_o");
      }
      // Sección 4: archivo (se maneja aparte abajo)

      // Normalizaciones rápidas
      const trim = (v?: string) => (v ? v.trim() : "");

      // IDs únicos sin Array.from ni Set (compatible ES5)
      const srcIds = state.usuarioregistradorIds || [];
      const uniqIds: number[] = [];
      for (let i = 0; i < srcIds.length; i++) {
        const n = srcIds[i];
        if (typeof n === "number" && !isNaN(n) && uniqIds.indexOf(n) === -1) {
          uniqIds.push(n);
        }
      }

      // Body completo con tipos correctos (SIN filtrar todavía)
      const bodyAll: Record<string, any> = {
        Title: trim(state.Title) || undefined,

        // Fechas: helper devuelve ISO o undefined
        fechaderegistro: toSpDate(state.fechaderegistro),
        periododesde: toSpDate(state.periododesde),
        periodohasta: toSpDate(state.periodohasta),

        // Textos
        ruc: trim(state.ruc) || undefined,
        codigodecontrato: trim(state.codigodecontrato) || undefined,

        // Lookup single (number)
        proveedorId:
          typeof state.proveedorId === "number" && !isNaN(state.proveedorId)
            ? state.proveedorId
            : undefined,

        // User multi (array de números)
        ...(uniqIds.length ? { usuarioregistradorId: uniqIds } : {}),

        // Tipo de formulario (string)
        tipodeformulario: trim(state.tipodeformulario) || undefined,

        // Año -> Edm.String
        a_x00f1_o: trim(state.anio) || undefined,
      };

      // Filtrar: enviar SOLO lo permitido por secciones visibles
      const body: Record<string, any> = {};
      for (const k in bodyAll) {
        if (!Object.prototype.hasOwnProperty.call(bodyAll, k)) continue;
        if (allowed.indexOf(k) === -1) continue; // filtro clave
        const v = (bodyAll as any)[k];
        if (v !== undefined && v !== null && v !== "") body[k] = v;
      }

      // Archivo: solo si la sección 4 está visible
      const archivoParaSubir = visibleSections.has(4)
        ? state.archivo
        : undefined;

      await sp.createFormulario(body, archivoParaSubir);

      setState({ ...initialState, success: "Guardado correctamente." });
      setPeopleSelected([]);
    } catch (e: unknown) {
      setState((s) => ({
        ...s,
        isSaving: false,
        error: e instanceof Error ? e.message : "Error desconocido al guardar.",
      }));
    } finally {
      setState((s) => ({ ...s, isSaving: false }));
      releaseLock();
    }
  };

  /* ===== Nav (Fluent) ===== */
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

    let match: ITipoFormularioItem | undefined = undefined;
    for (let i = 0; i < tipos.length; i++) {
      if (tipos[i].Id === id) {
        match = tipos[i];
        break;
      }
    }
    setSelectedId(id);
    setSelectedTemplate(match?.template);

    // <= NUEVO: guardar el texto del botón en el estado
    setField("tipodeformulario", match?.Title || "");
  };

  const fluentTheme =
    (props as any).fluentTheme || (props as any).theme || undefined;
  const classes = getClasses(fluentTheme || getTheme());

  const formActionsTokens: IStackTokens = { childrenGap: 8 };
  const headerStyles: IStackStyles = { root: { marginBottom: 8 } };

  return (
    <ThemeProvider theme={fluentTheme}>
      <div className={classes.root}>
        {/* ===== layout principal ===== */}
        <div className={classes.mainRow}>
          {/* izquierda: nav */}
          <div className={classes.leftCol}>
            <div className={classes.navPanel}>
              <div className={classes.navHeader}>
                <Text
                  variant="mediumPlus"
                  block
                  styles={{ root: { fontWeight: FontWeights.semibold } }}
                >
                  Tipo de formulario
                </Text>
              </div>
              <div className={classes.navBody}>
                {loadingTipos && (
                  <Spinner size={SpinnerSize.small} label="Cargando…" />
                )}
                {errorTipos && (
                  <MessageBar messageBarType={MessageBarType.error} isMultiline>
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
                      selectedKey={selectedId ? String(selectedId) : undefined}
                      ariaLabel="Tipos de formulario"
                      styles={{
                        root: { width: "100%" },
                        link: {
                          height: "auto",
                          minHeight: 40,
                          paddingTop: "clamp(6px, 0.8vw, 10px)",
                          paddingBottom: "clamp(6px, 0.8vw, 10px)",
                          lineHeight: 1.35,
                        },
                        linkText: {
                          whiteSpace: "normal",
                          wordBreak: "break-word",
                          fontSize: "clamp(12px, 1.2vw, 16px)",
                          lineHeight: 1.35,
                        },
                      }}
                      onRenderLink={(link?: INavLink) =>
                        link ? (
                          <TooltipHost content={link.name}>
                            <span className="ms-Nav-linkText">{link.name}</span>
                          </TooltipHost>
                        ) : null
                      }
                    />
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* derecha: formulario */}
          <div className={classes.rightCol}>
            <Stack
              styles={headerStyles}
              horizontal
              horizontalAlign="space-between"
              verticalAlign="center"
            >
              <Text
                variant="xLarge"
                styles={{ root: { fontWeight: FontWeights.semibold } }}
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
                  Elegí un <strong>Tipo de formulario</strong> para mostrar las
                  secciones.
                </MessageBar>
              )}

              {/* === Sección 1 === */}
              {visible.has(1) && (
                <Section title="1.- Identificación" classes={classes}>
                  <div className={classes.row}>
                    <div className={classes.c4}>
                      <TextField
                        label="Fecha de registro"
                        type="date"
                        value={state.fechaderegistro}
                        onChange={(_, v) =>
                          setField("fechaderegistro", v || "")
                        }
                        id="fechaderegistro"
                      />
                      <input
                        type="hidden"
                        id="tipodeformulario"
                        value={state.tipodeformulario}
                        readOnly
                      />
                    </div>
                    <div className={classes.c4}>
                      <TextField
                        label="RUC"
                        value={state.ruc}
                        onChange={(_, v) => setField("ruc", v || "")}
                        id="ruc"
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
                      />
                    </div>

                    <div className={classes.c6}>
                      <Label>Usuario registrador</Label>
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
                        }}
                        resolveDelay={300}
                      />
                    </div>
                  </div>
                </Section>
              )}

              {/* === Sección 2 === */}
              {visible.has(2) && (
                <Section title="2.- Datos del Contrato" classes={classes}>
                  <div className={classes.row}>
                    <div className={classes.c12}>
                      <TextField
                        label="Nombre del contrato"
                        onRenderLabel={renderLabelConAyuda(
                          "Nombre del contrato",
                          () =>
                            openHelp(
                              "Nombre del contrato",
                              "https://cnco.sharepoint.com/sites/DucumentosTrasportesPE/SiteAssets/Nombre_de_contrato.png"
                            )
                        )}
                        value={state.Title}
                        onChange={(_, v) => setField("Title", v || "")}
                        id="Title"
                      />
                    </div>
                    <div className={classes.c6}>
                      <TextField
                        label="Código de contrato"
                        onRenderLabel={renderLabelConAyuda(
                          "Código de contrato",
                          () =>
                            openHelp(
                              "Código de contrato",
                              "https://cnco.sharepoint.com/sites/DucumentosTrasportesPE/SiteAssets/codigo_contrato.png"
                            )
                        )}
                        value={state.codigodecontrato}
                        onChange={(_, v) =>
                          setField("codigodecontrato", v || "")
                        }
                        id="codigodecontrato"
                      />
                    </div>

                    <div className={classes.c12}>
                      <Separator>Plazo de contrato</Separator>
                      <div className={classes.row}>
                        <div className={classes.c6}>
                          <TextField
                            label="Fecha de inicio"
                            onRenderLabel={renderLabelConAyuda(
                              "Fecha de inicio",
                              () =>
                                openHelp(
                                  "Fecha de inicio",
                                  "https://cnco.sharepoint.com/sites/DucumentosTrasportesPE/SiteAssets/Fecha_contrato.png"
                                )
                            )}
                            type="date"
                            value={state.periododesde}
                            onChange={(_, v) =>
                              setField("periododesde", v || "")
                            }
                            id="periododesde"
                          />
                        </div>
                        <div className={classes.c6}>
                          <TextField
                            label="Fecha de término"
                            onRenderLabel={renderLabelConAyuda(
                              "Fecha de término",
                              () =>
                                openHelp(
                                  "Fecha de término",
                                  "https://cnco.sharepoint.com/sites/DucumentosTrasportesPE/SiteAssets/Fecha_contrato.png"
                                )
                            )}
                            type="date"
                            value={state.periodohasta}
                            onChange={(_, v) =>
                              setField("periodohasta", v || "")
                            }
                            id="periodohasta"
                          />
                        </div>
                      </div>
                    </div>
                  </div>
                </Section>
              )}

              {/* === Sección 3 === */}
              {visible.has(3) && (
                <Section title="2.- Datos generales" classes={classes}>
                  <div className={classes.row}>
                    <div className={classes.c12}>
                      <Separator>Periodo</Separator>
                      <div className={classes.row}>
                        <div className={classes.c6}>
                          <TextField
                            label="De"
                            type="date"
                            value={state.periododesde}
                            onChange={(_, v) =>
                              setField("periododesde", v || "")
                            }
                            id="periododesde_2"
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
                            id="periodohasta_2"
                          />
                        </div>
                      </div>
                    </div>
                    <div className={classes.c4}>
                      <TextField
                        label="Año"
                        type="number"
                        value={state.anio}
                        onChange={(_, v) => setField("anio", v || "")}
                        id="a_x00f1_o"
                      />
                    </div>
                  </div>
                </Section>
              )}

              {/* === Sección 4 === */}
              {visible.has(4) && (
                <Section title="3.- Cargar Documento" classes={classes}>
                  <div className={classes.row}>
                    <div className={classes.c8}>
                      <Label htmlFor="archivo">Archivo</Label>
                      <Stack
                        horizontal
                        tokens={{ childrenGap: 8 }}
                        verticalAlign="center"
                      >
                        <DefaultButton
                          text={
                            state.archivo
                              ? "Cambiar archivo"
                              : "Seleccionar archivo"
                          }
                          iconProps={{ iconName: "Upload" }}
                          onClick={() => fileRef.current?.click()}
                        />
                        <Text variant="small">
                          {state.archivo?.name ?? "Ningún archivo seleccionado"}
                        </Text>
                      </Stack>
                      <input
                        ref={fileRef}
                        id="archivo"
                        type="file"
                        style={{ display: "none" }}
                        onChange={(e) => {
                          const f =
                            e.target.files && e.target.files[0]
                              ? e.target.files[0]
                              : undefined;
                          setState((s) => ({
                            ...s,
                            archivo: f,
                            error: undefined,
                            success: undefined,
                          }));
                        }}
                      />
                      <Text
                        variant="small"
                        styles={{ root: { color: "#666" } }}
                      >
                        Opcional. Se adjunta al ítem de la lista.
                      </Text>
                    </div>
                  </div>
                </Section>
              )}

              {/* Acciones */}
              <Stack
                horizontal
                wrap
                tokens={formActionsTokens}
                style={{ marginTop: 8 }}
              >
                <PrimaryButton
                  type="submit"
                  text={state.isSaving ? "Guardando…" : "Guardar"}
                  disabled={state.isSaving || submitLockedRef.current}
                  iconProps={
                    state.isSaving ? { iconName: "Sync" } : { iconName: "Save" }
                  }
                />
                <DefaultButton
                  type="button"
                  text="Limpiar"
                  disabled={state.isSaving || submitLockedRef.current}
                  onClick={() => {
                    setState(initialState);
                    setPeopleSelected([]);
                  }}
                  iconProps={{ iconName: "Clear" }}
                />
              </Stack>

              {/* Mensajes */}
              {state.error && (
                <MessageBar
                  messageBarType={MessageBarType.success}
                  isMultiline
                  styles={{ root: { marginTop: 12 } }}
                >
                  El item se guardó exitosamente
                </MessageBar>
              )}
              {state.success && (
                <MessageBar
                  messageBarType={MessageBarType.success}
                  isMultiline
                  styles={{ root: { marginTop: 12 } }}
                >
                  El item se guardó exitosamente
                </MessageBar>
              )}
            </form>
          </div>
        </div>
      </div>

      {/* === MODAL DE AYUDA (images de Site Assets) === */}
      <Dialog
        hidden={!helpOpen}
        onDismiss={closeHelp}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: helpTitle,
        }}
        minWidth={480}
        maxWidth={900}
      >
        <Image
          src={helpUrl}
          alt={helpTitle}
          imageFit={ImageFit.contain}
          styles={{ root: { width: "100%", maxHeight: 600 } }}
        />
        <DialogFooter>
          <DefaultButton onClick={closeHelp} text="Cerrar" />
        </DialogFooter>
      </Dialog>
    </ThemeProvider>
  );
}
