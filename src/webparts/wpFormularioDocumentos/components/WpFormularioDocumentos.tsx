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
interface IFormState {
  fechaderegistro: string;
  ruc: string;
  proveedorId?: number;
  usuarioregistradorIds: number[];
  Title: string;
  codigodecontrato: string;
  periododesde: string;
  periodohasta: string;
  anio: string;
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
  return `${v}T00:00:00Z`;
};

/* ---- Estilos Fluent adaptados al tema ---- */
const getClasses = (theme = getTheme()) =>
  mergeStyleSets({
    root: {
      padding: 12,
      background: theme.palette.neutralLighterAlt,
      [`@media (min-width: 640px)`]: { padding: 16 },
      [`@media (min-width: 1024px)`]: { padding: 20 },
    },
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
    panel: {
      background: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 12,
      boxShadow: theme.effects.elevation8,
      overflow: "hidden",
      animation: "fadeIn .25s ease-out both",
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
    navPanel: {
      background: theme.palette.white,
      border: `1px solid ${theme.palette.neutralLight}`,
      borderRadius: 12,
      boxShadow: theme.effects.elevation8,
    },
    navHeader: {
      padding: "12px 16px",
      borderBottom: `1px solid ${theme.palette.neutralLight}`,
      background: theme.palette.themeLighterAlt,
      borderTopLeftRadius: 12,
      borderTopRightRadius: 12,
    },
    navBody: { padding: 12 },
    navResponsive: {
      maxHeight: "calc(100vh - 180px)",
      overflowY: "auto",
      ".ms-Nav-link": {
        display: "block !important",
        height: "auto !important",
        minHeight: 40,
        padding: "6px 12px",
        borderRadius: 8,
        border: "1px solid transparent",
      },
      ".ms-Nav-link:hover": {
        background: theme.palette.themeLighterAlt,
        borderColor: theme.palette.themeLighter,
      },
      ".ms-Nav .is-selected > .ms-Nav-link": {
        background: theme.palette.themeLighter,
        borderColor: theme.palette.themePrimary,
        boxShadow: "inset 0 0 0 1px rgba(0,0,0,.03)",
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

  const [proveedorOptions, setProveedorOptions] = useState<IDropdownOption[]>([]);
  const [peopleSelected, setPeopleSelected] = useState<IPersonaProps[]>([]);
  const [selectedId, setSelectedId] = useState<number | undefined>(undefined);
  const [selectedTemplate, setSelectedTemplate] = useState<string | undefined>(undefined);
  const visible = sectionsForTemplate(selectedTemplate);
  const fileRef = useRef<HTMLInputElement>(null);

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
        let lookupWebId: string | undefined = undefined;

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
        console.warn("Lookup proveedor falló:", err);
        setProveedorOptions([]);
      }
    };

    void loadProveedorOptions();
  }, [props.siteUrl, props.spHttpClient]);

  /* ===== People Picker ===== */
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

  const setField = (key: keyof IFormState, value: any): void => {
    setState((s) => ({
      ...s,
      [key]: value,
      error: undefined,
      success: undefined,
    }));
  };

  const onSave = async (): Promise<void> => {
    if (!acquireLock()) return;

    try {
      setState((s) => ({
        ...s,
        isSaving: true,
        error: undefined,
        success: undefined,
      }));

      const visibleSections = sectionsForTemplate(selectedTemplate);

      const allowed: string[] = [];
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

      const trim = (v?: string) => (v ? v.trim() : "");

      const srcIds = state.usuarioregistradorIds || [];
      const uniqIds: number[] = [];
      for (let i = 0; i < srcIds.length; i++) {
        const n = srcIds[i];
        if (typeof n === "number" && !isNaN(n) && uniqIds.indexOf(n) === -1) {
          uniqIds.push(n);
        }
      }

      const bodyAll: Record<string, any> = {
        Title: trim(state.Title) || undefined,
        fechaderegistro: toSpDate(state.fechaderegistro),
        periododesde: toSpDate(state.periododesde),
        periodohasta: toSpDate(state.periodohasta),
        ruc: trim(state.ruc) || undefined,
        codigodecontrato: trim(state.codigodecontrato) || undefined,
        proveedorId:
          typeof state.proveedorId === "number" && !isNaN(state.proveedorId)
            ? state.proveedorId
            : undefined,
        ...(uniqIds.length ? { usuarioregistradorId: uniqIds } : {}),
        tipodeformulario: trim(state.tipodeformulario) || undefined,
        a_x00f1_o: trim(state.anio) || undefined,
      };

      const body: Record<string, any> = {};
      for (const k in bodyAll) {
        if (!Object.prototype.hasOwnProperty.call(bodyAll, k)) continue;
        if (allowed.indexOf(k) === -1) continue;
        const v = (bodyAll as any)[k];
        if (v !== undefined && v !== null && v !== "") body[k] = v;
      }

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
    setField("tipodeformulario", match?.Title || "");
  };

  // usamos SIEMPRE el tema corporativo
  const classes = getClasses(cencoTheme);
  const formActionsTokens: IStackTokens = { childrenGap: 8 };
  const headerStyles: IStackStyles = { root: { marginBottom: 8 } };

  return (
    <ThemeProvider theme={cencoTheme}>
      <div className={classes.root}>
        <div className={classes.mainRow}>
          {/* izquierda */}
          <div className={classes.leftCol}>
            <div className={classes.navPanel}>
              <div className={classes.navHeader}>
                <Text
                  variant="mediumPlus"
                  block
                  styles={{ root: { fontWeight: FontWeights.semibold, color: cencoTheme.palette.themePrimary } }}
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
                styles={{ root: { fontWeight: FontWeights.semibold, color: cencoTheme.palette.themePrimary } }}
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
