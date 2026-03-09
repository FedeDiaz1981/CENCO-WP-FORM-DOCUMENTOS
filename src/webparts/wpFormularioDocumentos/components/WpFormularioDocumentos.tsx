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
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  Nav,
  INavLinkGroup,
  INavLink,
  getTheme,
  mergeStyleSets,
  FontWeights,
  Stack,
  Dropdown,
  IDropdownOption,
  Icon,
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
    themeLighter: "#deebf8",
    themeLight: "#c2daf1",
    themeTertiary: "#7eb2db",
    themeSecondary: "#2f7fc0",
    themeDarkAlt: "#004d87",
    themeDark: "#00406f",
    themeDarker: "#002f51",
    neutralLighterAlt: "#f4f9ff",
    neutralLighter: "#edf4fb",
    neutralLight: "#d7e5f3",
    neutralQuaternaryAlt: "#ccdaea",
    neutralQuaternary: "#c1d3e6",
    neutralTertiaryAlt: "#b5c7dc",
    neutralTertiary: "#333333",
    neutralSecondary: "#55687c",
    neutralPrimaryAlt: "#233140",
    neutralPrimary: "#1e2a36",
    neutralDark: "#1f1f1f",
    black: "#1a1a1a",
    white: "#ffffff",
  },
  effects: {
    roundedCorner2: "18px",
    elevation8: "0 12px 28px rgba(0,87,166,.12)" as any,
  },
});

const BRAND = {
  canvas: "#eef4fb",
  shell: "#f7fbff",
  ink: "#1e2a36",
  muted: "#617284",
  border: "#cad9ea",
  soft: "#e6f0fa",
};

const HERO_BG =
  "radial-gradient(circle at 14% 20%, rgba(255,255,255,.18) 0 82px, transparent 83px), radial-gradient(circle at 86% -8%, rgba(255,255,255,.18) 0 128px, transparent 130px), linear-gradient(135deg, #005596 0%, #0067b2 48%, #0072bc 100%)";

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

  // ✅ NUEVOS CAMPOS
  nombre?: string; // single line text
  restringido?: boolean; // yes/no
  mostrar?: boolean; // yes/no (si false => no mostrar)
  bloquear?: boolean; // yes/no (si true => no guardar)
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
        <div className={classes.sectionTitleRow}>
          <Icon iconName={getSectionIconName(title)} className={classes.sectionIcon} />
          <Text
            variant="large"
            block
            styles={{
              root: {
                fontWeight: FontWeights.semibold,
                color: cencoTheme.palette.white,
                lineHeight: 1.1,
              },
            }}
          >
            {title}
          </Text>
        </div>
      </div>
      <div className={classes.panelBody}>{children}</div>
    </section>
  );
}

function getSectionIconName(title: string): string {
  const normalized = (title || "").toLowerCase();

  if (normalized.indexOf("identific") !== -1) return "ContactCard";
  if (normalized.indexOf("contrato") !== -1) return "PageEdit";
  if (normalized.indexOf("codigo de documentos") !== -1) return "BulletedList";
  if (normalized.indexOf("cargar documento") !== -1) return "Attach";
  if (normalized.indexOf("datos generales") !== -1) return "PageList";
  return "Page";
}

const toSpDate = (yyyyMmDd: string): string | undefined => {
  const v = (yyyyMmDd || "").trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  if (!m) return undefined;

  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const local = new Date(y, mo - 1, d);
  if (
    isNaN(local.getTime()) ||
    local.getFullYear() !== y ||
    local.getMonth() !== mo - 1 ||
    local.getDate() !== d
  ) {
    return undefined;
  }

  // Use midday UTC to avoid day shifting when SharePoint applies timezone conversion.
  return `${v}T12:00:00Z`;
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

const normName = (s: string): string => (s || "").trim().toLowerCase();

const getClasses = (theme = getTheme()) =>
  mergeStyleSets({
    page: {
      background: `linear-gradient(180deg, ${BRAND.shell} 0%, ${BRAND.canvas} 100%)`,
      minHeight: "100%",
      padding: 16,
      [`@media (min-width: 640px)`]: { padding: 20 },
      [`@media (min-width: 1024px)`]: { padding: 24 },
    },
    container: { maxWidth: 1220, margin: "0 auto" },
    mainRow: {
      display: "flex",
      alignItems: "stretch",
      gap: 22,
      flexWrap: "nowrap",
      [`@media (max-width: 920px)`]: { flexWrap: "wrap" },
    },
    leftCol: {
      flex: "0 0 300px",
      width: 300,
      position: "sticky",
      top: 16,
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
      gap: 18,
    },
    headerCard: {
      background: HERO_BG,
      border: "1px solid rgba(255,255,255,.12)",
      borderRadius: 28,
      boxShadow: "0 24px 44px rgba(0,87,166,.22)",
      padding: "24px 26px",
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: 18,
      flexWrap: "wrap",
      overflow: "hidden",
      position: "relative",
      selectors: {
        "&::after": {
          content: "\"\"",
          position: "absolute",
          inset: "auto -58px -88px auto",
          width: 210,
          height: 210,
          borderRadius: "50%",
          background: "rgba(255,255,255,.08)",
          pointerEvents: "none",
        },
      },
    },
    headerTitleRow: {
      display: "flex",
      alignItems: "center",
      gap: 14,
      minWidth: 0,
      flex: "1 1 320px",
    },
    headerText: {
      display: "flex",
      flexDirection: "column",
      minWidth: 0,
      maxWidth: 520,
      paddingRight: 12,
    },
    headerIcon: {
      fontSize: 24,
      color: theme.palette.white,
      width: 56,
      height: 56,
      borderRadius: "50%",
      display: "inline-flex",
      alignItems: "center",
      justifyContent: "center",
      background: "rgba(255,255,255,.14)",
      border: "1px solid rgba(255,255,255,.22)",
      boxShadow: "0 10px 22px rgba(0,0,0,.12)",
    },
    headerTitle: {
      display: "block",
      fontSize: 28,
      fontWeight: 700,
      color: theme.palette.white,
      lineHeight: 1.12,
      whiteSpace: "normal",
      wordBreak: "break-word",
      selectors: {
        "@media (max-width: 640px)": {
          fontSize: 24,
        },
      },
    },
    headerHint: {
      display: "block",
      marginTop: 8,
      fontSize: 13,
      lineHeight: 1.35,
      color: "rgba(255,255,255,.82)",
    },
    panel: {
      background: "rgba(255,255,255,.94)",
      border: `1px solid ${BRAND.border}`,
      borderRadius: 24,
      boxShadow: "0 18px 36px rgba(0,87,166,.09)",
      overflow: "hidden",
      animation: "fadeIn .18s ease-out both",
    },
    panelHeader: {
      padding: "18px 22px 0",
    },
    sectionTitleRow: {
      display: "flex",
      alignItems: "center",
      gap: 10,
      width: "fit-content",
      padding: "8px 16px",
      borderRadius: 999,
      background: "linear-gradient(135deg, #005596 0%, #0072bc 100%)",
      boxShadow: "0 10px 18px rgba(0,87,166,.18)",
    },
    sectionIcon: {
      fontSize: 18,
      color: theme.palette.white,
    },
    panelBody: {
      padding: "18px 22px 22px",
      [`@media (min-width: 1024px)`]: { padding: "20px 24px 24px" },
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
      background: "rgba(255,255,255,.94)",
      border: `1px solid ${BRAND.border}`,
      borderRadius: 24,
      boxShadow: "0 18px 36px rgba(0,87,166,.09)",
      overflow: "hidden",
    },
    navHeader: {
      padding: "18px 20px",
      borderBottom: `1px solid ${BRAND.border}`,
      background: `linear-gradient(180deg, ${theme.palette.themeLighterAlt} 0%, ${BRAND.soft} 100%)`,
    },
    navTitleRow: {
      display: "flex",
      alignItems: "center",
      gap: 10,
    },
    navIcon: {
      fontSize: 18,
      color: theme.palette.themePrimary,
    },
    navBody: { padding: 16 },
    navResponsive: {
      maxHeight: "calc(100vh - 220px)",
      overflowY: "auto",
      selectors: {
        "&& .ms-Nav-link": {
          display: "block !important",
          height: "auto !important",
          minHeight: "48px !important",
          padding: "12px 14px !important",
          borderRadius: "16px !important",
          border: `1px solid ${BRAND.border} !important`,
          background: "rgba(255,255,255,.86) !important",
          boxShadow: "0 6px 14px rgba(0,87,166,.05) !important",
        },
        "&& .ms-Nav-linkText": {
          whiteSpace: "normal !important",
          wordBreak: "break-word !important",
          lineHeight: "1.35 !important",
          fontSize: "15px !important",
          fontWeight: "600 !important",
          color: `${BRAND.ink} !important`,
        },
        "&& .ms-Nav-link:hover": {
          background: `${theme.palette.white} !important`,
          borderColor: `#8bb8df !important`,
          transform: "translateY(-1px)",
        },
        "&& .ms-Nav .is-selected > .ms-Nav-link": {
          background: `${theme.palette.white} !important`,
          borderColor: `${theme.palette.themePrimary} !important`,
          boxShadow: "0 0 0 2px rgba(0,87,166,.12) inset !important",
        },
      },
    },

    fileRow: {
      display: "flex",
      alignItems: "center",
      justifyContent: "space-between",
      gap: 8,
      padding: "10px 12px",
      border: `1px solid ${BRAND.border}`,
      borderRadius: 18,
      background: "linear-gradient(180deg, rgba(255,255,255,.98) 0%, #f6fbff 100%)",
      marginTop: 8,
      boxShadow: "0 8px 18px rgba(0,87,166,.06)",
    },
    fileUploadRow: {
      display: "flex",
      alignItems: "center",
      gap: 12,
      flexWrap: "wrap",
    },
    hiddenFileInput: {
      display: "none",
    },
    fileHelpText: {
      color: BRAND.muted,
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
      gap: 12,
      alignItems: "center",
      marginTop: 12,
    },
    progressWrap: {
      marginTop: 12,
      background: "rgba(255,255,255,.96)",
      border: `1px solid ${BRAND.border}`,
      borderRadius: 20,
      boxShadow: "0 16px 30px rgba(0,87,166,.09)",
      padding: "14px 16px",
    },

    "@global": {
      "@keyframes fadeIn": {
        from: { opacity: 0, transform: "translateY(4px)" },
        to: { opacity: 1, transform: "none" },
      },
      ".cenco-doc-shell .ms-Label": {
        color: BRAND.ink,
        fontWeight: 600,
      },
      ".cenco-doc-shell .ms-TextField-fieldGroup, .cenco-doc-shell .ms-Dropdown-title, .cenco-doc-shell .ms-BasePicker-text": {
        minHeight: "46px",
        borderRadius: "18px",
        borderColor: `${BRAND.border} !important`,
        background: "#ffffff",
        boxShadow: "0 6px 16px rgba(0,87,166,.05)",
      },
      ".cenco-doc-shell .ms-Dropdown-title": {
        lineHeight: "44px",
      },
      ".cenco-doc-shell .ms-TextField-fieldGroup:hover, .cenco-doc-shell .ms-Dropdown-title:hover, .cenco-doc-shell .ms-BasePicker-text:hover": {
        borderColor: "#8bb8df !important",
      },
      ".cenco-doc-shell .ms-TextField-field": {
        fontSize: "14px",
      },
      ".cenco-doc-shell .ms-BasePicker-text input, .cenco-doc-shell .ms-DatePicker input": {
        fontSize: "14px",
      },
      ".cenco-doc-shell .ms-Nav-link:focus::after": {
        borderRadius: "16px",
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

  const selectedTipo = useMemo(() => {
    if (!selectedId) return undefined;
    for (let i = 0; i < tipos.length; i++)
      if (tipos[i].Id === selectedId) return tipos[i];
    return undefined;
  }, [tipos, selectedId]);

  const isBlocked = !!selectedTipo?.bloquear;
  const isRestricted = !!selectedTipo?.restringido;
  const restrictedName = (selectedTipo?.nombre || "").trim();

  const visible = useMemo(
    () => visibleSectionsFor(selectedTemplate, state.tipodeformulario),
    [selectedTemplate, state.tipodeformulario]
  );

  const [allowedRegistradores, setAllowedRegistradores] = useState<
    IPersonaProps[]
  >([]);

  const classes = getClasses(cencoTheme);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

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

  const primaryButtonStyles: IButtonStyles = useMemo(
    () => ({
      root: {
        height: 44,
        minHeight: 44,
        padding: "0 20px",
        borderRadius: 999,
        fontSize: 14,
        border: "none",
        background: "linear-gradient(135deg, #005596 0%, #0072bc 100%)",
        boxShadow: "0 12px 22px rgba(0,87,166,.2)",
      },
      rootHovered: {
        background: "linear-gradient(135deg, #004d87 0%, #0067b2 100%)",
        boxShadow: "0 14px 24px rgba(0,87,166,.24)",
      },
      rootPressed: {
        background: "linear-gradient(135deg, #00406f 0%, #005596 100%)",
      },
      flexContainer: { height: 44 },
      label: {
        fontSize: 14,
        fontWeight: 700,
        lineHeight: "1 !important",
        color: "#ffffff",
      },
      icon: { fontSize: 16, color: "#ffffff" },
    }),
    []
  );

  const secondaryButtonStyles: IButtonStyles = useMemo(
    () => ({
      root: {
        height: 44,
        minHeight: 44,
        padding: "0 20px",
        borderRadius: 999,
        fontSize: 14,
        border: `1px solid ${BRAND.border}`,
        background: "rgba(255,255,255,.94)",
        boxShadow: "0 8px 18px rgba(0,87,166,.08)",
      },
      rootHovered: {
        background: "#ffffff",
        borderColor: "#8bb8df",
      },
      flexContainer: { height: 44 },
      label: {
        fontSize: 14,
        fontWeight: 600,
        lineHeight: "1 !important",
        color: BRAND.ink,
      },
      icon: { fontSize: 16, color: BRAND.ink },
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

        const mapped: ITipoFormularioItem[] = data
          .map((t) => {
            const mostrarRaw = (t as any).mostrar;
            const bloquearRaw = (t as any).bloquear;
            const restringidoRaw = (t as any).restringido;

            return {
              Id: Number(t.Id),
              Title: String(t.Title ?? ""),
              orden: Number((t as any).orden ?? 0),
              template: String((t as any).template ?? (t as any).Template ?? "")
                .toUpperCase()
                .trim(),

              // ✅ map nuevos campos
              nombre: String((t as any).nombre ?? ""),
              // defaults: mostrar=true si viene undefined/null; bloquear/restringido=false
              mostrar: mostrarRaw === false ? false : true,
              bloquear: bloquearRaw === true,
              restringido: restringidoRaw === true,
            };
          })
          .sort((a, b) => (a.orden ?? 0) - (b.orden ?? 0))
          // ✅ mostrar=false => no aparece como opción
          .filter((x) => x.mostrar !== false);

        setTipos(mapped);
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

  const restrictedNameOk = useMemo(() => {
    if (!isRestricted) return true;
    const expected = normName(restrictedName);
    if (!expected) return true; // si no configuraron nombre, no bloqueamos
    if (!state.archivos || state.archivos.length !== 1) return false;
    return normName(state.archivos[0]?.name || "") === expected;
  }, [isRestricted, restrictedName, state.archivos]);

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

    const filesPresent =
      !needFiles || (state.archivos && state.archivos.length > 0);

    // ✅ restringido: exactamente 1 archivo + nombre esperado
    const restrictedOk = !needFiles
      ? true
      : !isRestricted
      ? true
      : restrictedNameOk;

    return (
      !!selectedTemplate &&
      !isBlocked &&
      desdeOk &&
      hastaOk &&
      rangoOk &&
      noVencidaOk &&
      codigosOk &&
      filesPresent &&
      restrictedOk &&
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
    isBlocked,
    isRestricted,
    restrictedNameOk,
  ]);

  /* ===== Files: SIN Array.from (compat TS viejo) ===== */
  const onPickFiles = (files: FileList | null): void => {
    if (!files || files.length === 0) return;

    if (isRestricted) {
      const f0 = files.item(0);
      if (!f0) return;

      // ✅ restringido: reemplaza y deja solo 1
      setState((s) => ({
        ...s,
        archivos: [f0],
        error: undefined,
        success: undefined,
      }));
      return;
    }

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

  const onSave = async (): Promise<void> => {
    // ✅ bloquear=true => no deja guardar
    if (isBlocked) {
      setState((s) => ({
        ...s,
        error: "Este formulario está bloqueado por el administrador",
        success: undefined,
      }));
      return;
    }

    // ✅ restringido: validación dura en onSave también
    if (isRestricted) {
      const expected = normName(restrictedName);
      if (expected && (!state.archivos || state.archivos.length !== 1)) {
        setState((s) => ({
          ...s,
          error: `Debés adjuntar 1 solo archivo: "${restrictedName}".`,
          success: undefined,
        }));
        return;
      }
      if (expected && normName(state.archivos?.[0]?.name || "") !== expected) {
        setState((s) => ({
          ...s,
          error: `El archivo debe llamarse exactamente "${restrictedName}".`,
          success: undefined,
        }));
        return;
      }
    }

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

      // ✅ restringido: si el nombre está configurado, lo exigimos también aquí
      if (visibleSections.has(4) && isRestricted) {
        const expected = normName(restrictedName);
        if (expected) {
          if (!state.archivos || state.archivos.length !== 1) {
            throw new Error(
              `Debés adjuntar 1 solo archivo: "${restrictedName}".`
            );
          }
          if (normName(state.archivos[0]?.name || "") !== expected) {
            throw new Error(
              `El archivo debe llamarse exactamente "${restrictedName}".`
            );
          }
        }
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

      // ✅ RESET TOTAL: como recién abierto (sin tipo seleccionado, sin mensajes)
      setState({ ...initialState });
      setPeopleSelected([]);
      setSelectedId(undefined);
      setSelectedTemplate(undefined);

      // ✅ si proveedor=true, re-hidrata como al abrir
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

    // ✅ si el tipo es restringido, limpiamos archivos para evitar arrastrar de otro tipo
    const willBeRestricted = !!match?.restringido;

    setState((s) => ({
      ...s,
      tipodeformulario: tipo,
      error: undefined,
      success: undefined,
      archivos: willBeRestricted ? [] : s.archivos,
      codigosDocumentos: isSctrExcel(tipo)
        ? s.codigosDocumentos?.length
          ? s.codigosDocumentos
          : [{ codigo: "", descripcion: "" }]
        : [{ codigo: "", descripcion: "" }],
    }));
  };

  return (
    <ThemeProvider theme={cencoTheme}>
      <div className={`${classes.page} cenco-doc-shell`}>
        <div className={classes.container}>
          <div className={classes.mainRow}>
            {/* izquierda */}
            <div className={classes.leftCol}>
              <div className={classes.navPanel}>
                <div className={classes.navHeader}>
                  <div className={classes.navTitleRow}>
                    <Icon iconName="BulletedList" className={classes.navIcon} />
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
                        selectedKey={selectedId ? String(selectedId) : undefined}
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
              <div className={classes.headerCard}>
                <div className={classes.headerTitleRow}>
                  <Icon iconName="Page" className={classes.headerIcon} />
                  <div className={classes.headerText}>
                    <div className={classes.headerTitle}>Documentos</div>
                  </div>
                </div>
              </div>

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

                {/* ✅ bloquear=true banner */}
                {selectedTemplate && isBlocked && (
                  <MessageBar
                    messageBarType={MessageBarType.warning}
                    isMultiline
                    styles={{ root: { marginTop: 12 } }}
                  >
                    Este formulario está bloqueado por el administrador
                  </MessageBar>
                )}

                {/* ✅ restringido hint */}
                {selectedTemplate && isRestricted && !!restrictedName && (
                  <MessageBar
                    messageBarType={MessageBarType.info}
                    isMultiline
                    styles={{ root: { marginTop: 12 } }}
                  >
                    Este formulario requiere un único archivo con nombre exacto:{" "}
                    <strong>{restrictedName}</strong>
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
                              const ids2 = arr
                                .map((p) => Number(p.id))
                                .filter((n) => !isNaN(n));
                              setField("usuarioregistradorIds", ids2);
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

                        <div className={classes.plazoHeader}>
                          Plazo de contrato
                        </div>

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
                            label={
                              isSctrExcel(state.tipodeformulario)
                                ? "Tiempo de vigencia"
                                : "Año"
                            }
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
                            styles={secondaryButtonStyles}
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
                          ref={fileInputRef}
                          className={classes.hiddenFileInput}
                          type="file"
                          multiple={!isRestricted}
                          onChange={(e) => {
                            onPickFiles(e.currentTarget.files);
                            e.currentTarget.value = "";
                          }}
                        />

                        <div className={classes.fileUploadRow}>
                          <DefaultButton
                            text={isRestricted ? "Adjuntar archivo" : "Adjuntar archivos"}
                            iconProps={{ iconName: "Upload" }}
                            onClick={() => fileInputRef.current?.click()}
                            styles={secondaryButtonStyles}
                          />
                          <Text variant="small" className={classes.fileHelpText}>
                            {isRestricted
                              ? "Se permite un solo archivo."
                              : "Puedes cargar varios archivos."}
                          </Text>
                        </div>

                        {isRestricted &&
                          !!restrictedName &&
                          !restrictedNameOk && (
                            <MessageBar
                              messageBarType={MessageBarType.severeWarning}
                              isMultiline
                              styles={{ root: { marginTop: 10 } }}
                            >
                              Debés adjuntar <strong>1</strong> archivo con
                              nombre exacto: <strong>{restrictedName}</strong>
                            </MessageBar>
                          )}

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
                      state.isSaving ? { iconName: "Sync" } : { iconName: "Save" }
                    }
                    styles={primaryButtonStyles}
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
                    styles={secondaryButtonStyles}
                  />
                </div>

                {state.isSaving && (
                  <div className={classes.progressWrap}>
                    <ProgressIndicator label="Guardando..." />
                  </div>
                )}

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
