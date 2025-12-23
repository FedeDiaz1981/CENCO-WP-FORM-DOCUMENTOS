import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

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
  codigodedocumentos?: string;
}

export interface IWpFormularioDocumentosProps {
  siteUrl: string;
  spHttpClient: SPHttpClient;
  listTitle?: string;
  fieldMap: IFieldMap;
  proveedor?: boolean;
  context?: WebPartContext;

  /** ✅ imágenes de ayuda (URL) */
  helpImgNombreContrato?: string;
  helpImgCodigoContrato?: string;
  helpImgFechaInicio?: string;
  helpImgFechaFin?: string;
}
