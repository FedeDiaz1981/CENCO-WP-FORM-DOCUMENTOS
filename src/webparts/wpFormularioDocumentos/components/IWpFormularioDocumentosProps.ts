import { SPHttpClient } from "@microsoft/sp-http";

export interface IFieldMap {
  title?: string;
  fechaderegistro?: string;
  ruc?: string;
  proveedor?: string;
  usuarioregistrador?: string;
  codigodecontrato?: string;
  periododesde?: string;
  periodohasta?: string;
  anio?: string;
}

export interface IWpFormularioDocumentosProps {
  siteUrl: string;
  spHttpClient: SPHttpClient;

  /** Lista elegida en el Property Pane */
  listTitle?: string;

  /** Mapeo de campos (InternalName) elegido en el Property Pane */
  fieldMap: IFieldMap;
}
