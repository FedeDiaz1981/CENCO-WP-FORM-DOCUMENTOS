import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel,
  PropertyPaneToggle,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import WpFormularioDocumentos from './components/WpFormularioDocumentos';
import { IWpFormularioDocumentosProps, IFieldMap } from './components/IWpFormularioDocumentosProps';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IWpFormularioDocumentosWebPartProps {
  listTitle?: string;
  fieldMap: IFieldMap;
  proveedor?: boolean;

  /** Ayudas (URLs) */
  helpImgNombreContrato?: string;
  helpImgCodigoContrato?: string;
  helpImgFechaInicio?: string;
  helpImgFechaFin?: string;
}

export default class WpFormularioDocumentosWebPart
  extends BaseClientSideWebPart<IWpFormularioDocumentosWebPartProps> {

  /** Opciones cacheadas */
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _fieldOptions: IPropertyPaneDropdownOption[] = [];
  private _loadingLists = false;
  private _loadingFields = false;

  public render(): void {
    const element: React.ReactElement<IWpFormularioDocumentosProps> = React.createElement(
      WpFormularioDocumentos,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient as SPHttpClient,
        listTitle: this.properties.listTitle,
        fieldMap: this.properties.fieldMap || {},
        proveedor: this.properties.proveedor,
        context: this.context,
        helpImgNombreContrato: this.properties.helpImgNombreContrato,
        helpImgCodigoContrato: this.properties.helpImgCodigoContrato,
        helpImgFechaInicio: this.properties.helpImgFechaInicio,
        helpImgFechaFin: this.properties.helpImgFechaFin
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Escapa comillas simples para OData */
  private _escODataString(value: string): string {
    return (value || '').replace(/'/g, "''");
  }

  /** Carga las listas del sitio */
  private async _loadLists(): Promise<void> {
    if (this._loadingLists) return;
    this._loadingLists = true;
    try {
      const url =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists?$select=Title,BaseTemplate,Hidden` +
        `&$filter=Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 101)`;

      const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._listOptions = [];
        return;
      }

      const json = await res.json();
      const rows = (json.value || []) as Array<{ Title: string; BaseTemplate: number }>;

      this._listOptions = rows
        .map(r => ({ key: r.Title, text: r.Title }))
        .sort((a, b) => a.text.localeCompare(b.text));
    } catch {
      this._listOptions = [];
    } finally {
      this._loadingLists = false;
    }
  }

  /** Carga campos de la lista seleccionada */
  private async _loadFields(listTitle?: string): Promise<void> {
    if (!listTitle || this._loadingFields) { this._fieldOptions = []; return; }
    this._loadingFields = true;

    try {
      const safeTitle = this._escODataString(listTitle);

      const url =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${safeTitle}')/fields` +
        `?$select=InternalName,Title,Hidden,ReadOnlyField,Sealed,TypeAsString` +
        `&$filter=Hidden eq false and Sealed eq false`;

      const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) {
        this._fieldOptions = [];
        return;
      }

      const json = await res.json();
      const rows = (json.value || []) as Array<{
        InternalName: string;
        Title: string;
        TypeAsString: string;
        ReadOnlyField?: boolean;
      }>;

      const allow = new Set<string>([
        'Text', 'Note', 'Number', 'DateTime', 'Currency', 'Boolean',
        'User', 'UserMulti',
        'Lookup', 'LookupMulti'
      ]);

      this._fieldOptions = rows
        .filter(r => allow.has(r.TypeAsString))
        .map(r => {
          const ro = r.ReadOnlyField ? ' RO' : '';
          return { key: r.InternalName, text: `${r.Title} (${r.InternalName})${ro}` };
        })
        .sort((a, b) => a.text.localeCompare(b.text));
    } catch {
      this._fieldOptions = [];
    } finally {
      this._loadingFields = false;
    }
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this._loadLists();
    await this._loadFields(this.properties.listTitle);
    this.context.propertyPane.refresh();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'listTitle' && newValue !== oldValue) {
      await this._loadFields(newValue as string);
      this.properties.fieldMap = {};
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fieldOptions = this._fieldOptions.length
      ? this._fieldOptions
      : [{ key: '', text: this._loadingFields ? 'Cargando campos…' : 'Seleccione una lista primero' }];

    return {
      pages: [
        {
          header: { description: 'Vinculación dinámica a listas/campos' },
          groups: [
            {
              groupName: 'Lista de SharePoint',
              groupFields: [
                PropertyPaneDropdown('listTitle', {
                  label: 'Lista de destino',
                  options: this._listOptions.length
                    ? this._listOptions
                    : [{ key: '', text: this._loadingLists ? 'Cargando listas…' : 'No hay listas' }],
                  selectedKey: this.properties.listTitle
                }),
                PropertyPaneLabel('', { text: 'Seleccione la lista y luego asigne los campos.' }),
                PropertyPaneToggle('proveedor', {
                  label: 'Proveedor',
                  onText: 'Sí',
                  offText: 'No'
                })
              ]
            },
            {
              groupName: 'Mapeo de campos (InternalName)',
              groupFields: [
                PropertyPaneDropdown('fieldMap.title', { label: '→ Nombre del contrato', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.fechaderegistro', { label: '→ Fecha de registro', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.ruc', { label: '→ RUC', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.proveedor', { label: '→ Razón social', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.usuarioregistrador', { label: '→ Usuario registrador', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.codigodecontrato', { label: '→ Código de contrato', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.periododesde', { label: '→ Fecha inicio (plazo / periodo)', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.periodohasta', { label: '→ Fecha término (plazo / periodo)', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.anio', { label: '→ Año', options: fieldOptions }),
                PropertyPaneDropdown('fieldMap.codigodedocumentos', { label: '→ Código de documentos (multilínea)', options: fieldOptions }),
              ]
            },
            {
              groupName: 'Ayudas (imágenes en modal)',
              groupFields: [
                PropertyPaneTextField('helpImgNombreContrato', {
                  label: 'Imagen ayuda → Nombre del contrato (URL)',
                  placeholder: 'https://.../nombre-contrato.png'
                }),
                PropertyPaneTextField('helpImgCodigoContrato', {
                  label: 'Imagen ayuda → Código de contrato (URL)',
                  placeholder: 'https://.../codigo-contrato.png'
                }),
                PropertyPaneTextField('helpImgFechaInicio', {
                  label: 'Imagen ayuda → Fecha de inicio (URL)',
                  placeholder: 'https://.../fecha-inicio.png'
                }),
                PropertyPaneTextField('helpImgFechaFin', {
                  label: 'Imagen ayuda → Fecha de término (URL)',
                  placeholder: 'https://.../fecha-fin.png'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
