import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel,
  //PropertyPaneToggle
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
        fieldMap: this.properties.fieldMap || {}
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

  /** Carga las listas del sitio */
  private async _loadLists(): Promise<void> {
    if (this._loadingLists) return;
    this._loadingLists = true;
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,BaseTemplate,Hidden&$filter=Hidden eq false and (BaseTemplate eq 100 or BaseTemplate eq 101)`;
      const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await res.json();
      const rows = (json.value || []) as Array<{ Title: string; BaseTemplate: number }>;
      this._listOptions = rows.map(r => ({ key: r.Title, text: r.Title })).sort((a,b)=> a.text.localeCompare(b.text));
    } catch (e) {
      // silencioso en PP
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
      const url =
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields` +
        `?$select=InternalName,Title,Hidden,ReadOnlyField,Sealed,TypeAsString` +
        `&$filter=Hidden eq false and ReadOnlyField eq false and Sealed eq false`;
      const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await res.json();
      const rows = (json.value || []) as Array<{ InternalName: string; Title: string; TypeAsString: string }>;

      // Dejamos campos comunes de entrada
      const allow = new Set<string>([
        'Text','Note','Number','DateTime','Currency','User','Lookup','Boolean'
      ]);

      this._fieldOptions = rows
        .filter(r => allow.has(r.TypeAsString))
        .map(r => ({ key: r.InternalName, text: `${r.Title} (${r.InternalName})` }))
        .sort((a,b)=> a.text.localeCompare(b.text));
    } catch (e) {
      this._fieldOptions = [];
    } finally {
      this._loadingFields = false;
    }
  }

  /** Carga inicial asincrónica del PP */
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this._loadLists();
    await this._loadFields(this.properties.listTitle);
    this.context.propertyPane.refresh();
  }

  /** Si cambia la lista, recarga campos y limpia mapeos */
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    if (propertyPath === 'listTitle' && newValue !== oldValue) {
      await this._loadFields(newValue as string);
      // Limpio mapeos al cambiar de lista
      this.properties.fieldMap = {};
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fieldOptions = this._fieldOptions.length ? this._fieldOptions : [{ key: '', text: this._loadingFields ? 'Cargando campos…' : 'Seleccione una lista primero' }];

    return {
      pages: [
        {
          header: { description: "Vinculación dinámica a listas/campos" },
          groups: [
            {
              groupName: "Lista de SharePoint",
              groupFields: [
                PropertyPaneDropdown('listTitle', {
                  label: 'Lista de destino',
                  options: this._listOptions.length ? this._listOptions : [{ key: '', text: this._loadingLists ? 'Cargando listas…' : 'No hay listas' }],
                  selectedKey: this.properties.listTitle
                }),
                PropertyPaneLabel('', { text: 'Seleccione la lista y luego asigne los campos.' }),
              ]
            },
            {
              groupName: "Mapeo de campos (InternalName)",
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
              ]
            }
          ]
        }
      ]
    };
  }
}
