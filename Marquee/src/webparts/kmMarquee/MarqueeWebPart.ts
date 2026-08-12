import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

import {
  BaseClientSideWebPart,
  WebPartContext,
} from '@microsoft/sp-webpart-base';

import {
IPropertyPaneConfiguration,
IPropertyPaneDropdownOption,
PropertyPaneCheckbox,
PropertyPaneDropdown,
PropertyPaneLabel,
PropertyPaneSlider,
PropertyPaneTextField ,
PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'KmMarqueeWebPartStrings';
import KmMarquee from './components/Marquee';
import { IKmMarqueeProps } from './components/IMarqueeProps';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IDropdownOption {
  key: string;
  text: string;
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

export interface IKmMarqueeWebPartProps {
  description: string;
  dMode: DisplayMode;
  spfxContext: WebPartContext;
  marqueeBackColor: string;
  marqueeTextColor: string;
  marqueeActive:boolean;
  fontFamily: string;
  fontSize: string;
  fontStyle: string;
  fontBold: boolean;
  marqueeHeight: string;
  scrollSpeed: number;
  scrollDirection: string;
  placement: string;
  listName: string;
  viewId: string;
  messageField: string;
  messageDuration: number;
  enableDiagnostics: boolean;
}

export default class KmMarqueeWebPart extends BaseClientSideWebPart<IKmMarqueeWebPartProps> {
  private _lists: IDropdownOption[] = [];
  private _views: Array<IDropdownOption & { isDefault?: boolean }> = [];
  private _listFields: IDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IKmMarqueeProps> = React.createElement(
      KmMarquee,
      {
        description: this.properties.description,
        dMode: this.displayMode,
        spfxContext:this.context,
        marqueeBackColor: this.properties.marqueeBackColor || '#333333',
        marqueeTextColor: this.properties.marqueeTextColor || '#fcfcfc',
        marqueeActive: this.properties.marqueeActive,
        fontFamily: this.properties.fontFamily || '"Segoe UI", sans-serif',
        fontSize: this.properties.fontSize || '13px',
        fontStyle: this.properties.fontStyle || 'normal',
        fontBold: this.properties.fontBold !== false,
        marqueeHeight: this.properties.marqueeHeight || '32px',
        scrollSpeed: typeof this.properties.scrollSpeed === 'number' ? this.properties.scrollSpeed : 50,
        scrollDirection: this.properties.scrollDirection || 'left',
        placement: this.properties.placement || 'top',
        listName: this.properties.listName || '',
        viewId: this.properties.viewId || '',
        messageField: this.properties.messageField || '',
        messageDuration: typeof this.properties.messageDuration === 'number' ? this.properties.messageDuration : 8,
        enableDiagnostics: this.properties.enableDiagnostics !== false
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private handleColorPropertyChange(propertyPath: string, oldValue: any, newValue: any): void {
    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected onInit(): Promise<void> {
    this.logDiagnostic('onInit started. listName=' + String(this.properties.listName || '(none)'));
    return this.loadLists().then(() => {
      if (this.properties.listName) {
        return this.loadViews(this.properties.listName);
      }
      return Promise.resolve();
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.logDiagnostic('Property pane opened. listName=' + String(this.properties.listName || '(none)'));
    if (this._lists.length === 0) {
      this.loadLists();
    }
    if (this.properties.listName && this._views.length === 0) {
      this.loadViews(this.properties.listName);
    } else if (this.properties.listName && this.properties.viewId && this._listFields.length === 0) {
      this.loadViewFields(this.properties.listName, this.properties.viewId);
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    if (propertyPath === 'listName' && oldValue !== newValue) {
      this.properties.viewId = '';
      this.properties.messageField = '';
      this._views = [];
      this._listFields = [];
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      if (newValue) {
        this.loadViews(String(newValue));
      }
      this.context.propertyPane.refresh();
      return;
    }

    if (propertyPath === 'viewId' && oldValue !== newValue) {
      this.properties.messageField = '';
      this._listFields = [];
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      if (this.properties.listName && newValue) {
        this.loadViewFields(this.properties.listName, String(newValue));
      }
      this.context.propertyPane.refresh();
      this.render();
      return;
    }

    if (propertyPath === 'placement' && oldValue !== newValue) {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      this.render();
      return;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private async getJsonWithAcceptFallback(url: string): Promise<any> {
    let response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=verbose' }
      });
    }

    if (!response.ok) {
      response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });
    }

    if (!response.ok) {
      throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
    }

    return response.json();
  }

  private async loadLists(): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      this.logDiagnostic('Loading lists from web URL: ' + webUrl);
      const data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists?$select=Title,Hidden,BaseTemplate&$filter=Hidden eq false"
      );
      const lists = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);

      this._lists = lists.map((list: any) => {
        return { key: list.Title, text: list.Title };
      });
      this.logDiagnostic('Loaded lists successfully. Count=' + String(this._lists.length));
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load lists: ' + (error && error.message ? error.message : String(error)));
      this._lists = [];
    }
  }

  private async loadViews(listName: string): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/views?$select=Id,Title,DefaultView"
      );
      const views = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
      this._views = views.map((view: any) => {
        return { key: String(view.Id), text: String(view.Title), isDefault: view.DefaultView === true };
      });

      if (!this.properties.viewId && this._views.length > 0) {
        const defaultView = this._views.filter((view) => { return view.isDefault === true; })[0];
        this.properties.viewId = String((defaultView || this._views[0]).key);
      }

      this.logDiagnostic('Loaded views for list "' + listName + '". Count=' + String(this._views.length));
      this.context.propertyPane.refresh();
      if (this.properties.viewId) {
        await this.loadViewFields(listName, this.properties.viewId);
      }
      this.render();
    } catch (error) {
      this._views = [];
      this._listFields = [];
      this.context.propertyPane.refresh();
      this.logDiagnostic('Failed to load views for list "' + listName + '": ' + (error && error.message ? error.message : String(error)));
    }
  }

  private async loadViewFields(listName: string, viewId: string): Promise<void> {
    if (!listName || !viewId) {
      this._listFields = [];
      this.context.propertyPane.refresh();
      return;
    }

    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const listPath = "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')";
      const normalizedViewId = String(viewId).replace(/[{}]/g, '');
      const viewData = await this.getJsonWithAcceptFallback(
        webUrl + listPath + "/views/getById('" + encodeURIComponent(normalizedViewId) + "')/ViewFields"
      );
      const rawViewFields = viewData && viewData.value ? viewData.value
        : (viewData && viewData.Items ? viewData.Items : (viewData && viewData.d && viewData.d.Items ? viewData.d.Items : []));
      const viewFields: string[] = Array.isArray(rawViewFields) ? rawViewFields
        : (rawViewFields && rawViewFields.results ? rawViewFields.results : []);
      const fieldData = await this.getJsonWithAcceptFallback(
        webUrl + listPath + "/fields?$select=InternalName,Title,Hidden"
      );
      const fields = fieldData && fieldData.value ? fieldData.value
        : (fieldData && fieldData.d && fieldData.d.results ? fieldData.d.results : []);
      const fieldByName: { [name: string]: any } = {};
      fields.forEach((field: any) => {
        fieldByName[String(field.InternalName || '')] = field;
      });

      this._listFields = viewFields.map((fieldName: string) => {
        const field = fieldByName[fieldName] || { InternalName: fieldName, Title: fieldName };
        const internalName = String(field.InternalName || '');
        const title = String(field.Title || internalName);
        return { key: internalName, text: title + (title !== internalName ? ' (' + internalName + ')' : '') };
      });

      if (this.properties.listName !== listName || this.properties.viewId !== viewId) {
        return;
      }
      this.logDiagnostic('Loaded fields for selected view. Count=' + String(this._listFields.length));
      this.context.propertyPane.refresh();
    } catch (error) {
      this._listFields = [];
      this.context.propertyPane.refresh();
      this.logDiagnostic('Failed to load fields for selected view: ' + (error && error.message ? error.message : String(error)));
    }
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[KmMarqueeWebPart] ' + message);
  }

  private getWebPartVersion(): string {
    const solutionVersion = packageSolutionConfig && packageSolutionConfig.solution
      ? String(packageSolutionConfig.solution.version || '')
      : '';
    if (solutionVersion) {
      return solutionVersion;
    }

    const manifestVersion = this.context && this.context.manifest ? String(this.context.manifest.version || '') : '';
    if (manifestVersion && manifestVersion !== '*') {
      return manifestVersion;
    }

    return 'Unknown';
  }

  private _validateText(value: string): string {
    // If validation is not successful, return a string with error message.   
    if (/<(br|basefont|form|div|script|action|src|img|hr|input|source|iframe|param|area|meta|!--|col|link|option|base|img|wbr|!DOCTYPE).*?>|<(a|abbr|acronym|address|applet|article|aside|audio|b|bdi|bdo|big|blockquote|body|button|canvas|caption|center|cite|code|colgroup|command|datalist|dd|del|details|dfn|dialog|dir|div|dl|dt|em|embed|fieldset|figcaption|figure|font|footer|form|frameset|head|header|hgroup|h1|h2|h3|h4|h5|h6|html|i|iframe|ins|kbd|keygen|label|legend|li|map|mark|menu|meter|nav|noframes|noscript|object|ol|optgroup|output|p|pre|progress|q|rp|rt|ruby|s|samp|script|section|select|small|span|strike|strong|style|sub|summary|sup|table|tbody|td|textarea|tfoot|th|thead|time|title|tr|track|tt|u|ul|var|video).*?<\/\2>/i.test(value)) { 
      return "Text Required, please try again!";
    }
    // A list data source can supply the scrolling text instead of this field.
    else if (this.properties.listName && this.properties.messageField) {
      return "";
    }
    else if (value === null || value.trim().length === 0) {
      return 'Provide Description';
    }

  else if (value.length > 250) {
      return 'Description should not be longer than 250 characters';
    }
   else {
    // If validation is successful, return an empty string.
    return "";}
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fontFamilyOptions: IPropertyPaneDropdownOption[] = [
      { key: '"Segoe UI", sans-serif', text: 'Segoe UI' },
      { key: 'Arial, sans-serif', text: 'Arial' },
      { key: 'Calibri, sans-serif', text: 'Calibri' },
      { key: 'Cambria, serif', text: 'Cambria' },
      { key: 'Georgia, serif', text: 'Georgia' },
      { key: 'Tahoma, sans-serif', text: 'Tahoma' },
      { key: '"Times New Roman", serif', text: 'Times New Roman' },
      { key: '"Trebuchet MS", sans-serif', text: 'Trebuchet MS' },
      { key: 'Verdana, sans-serif', text: 'Verdana' }
    ];
    const fontStyleOptions: IPropertyPaneDropdownOption[] = [
      { key: 'normal', text: 'Normal' },
      { key: 'italic', text: 'Italic' },
      { key: 'oblique', text: 'Oblique' }
    ];
    const directionOptions: IPropertyPaneDropdownOption[] = [
      { key: 'left', text: 'Right to left' },
      { key: 'right', text: 'Left to right' }
    ];
    const listOptions: IPropertyPaneDropdownOption[] = this._lists.map((list) => {
      return { key: list.key, text: list.text };
    });
    const listFieldOptions: IPropertyPaneDropdownOption[] = this._listFields.map((field) => {
      return { key: field.key, text: field.text };
    });
    const viewOptions: IPropertyPaneDropdownOption[] = this._views.map((view) => {
      return { key: view.key, text: view.text };
    });

    const fullVersionLabel = 'Version: ' + this.getWebPartVersion();

    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: fullVersionLabel,
              groupFields: [
                PropertyPaneLabel('propertyPaneVersionInfo', {
                  text: ' '
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  placeholder: "This is the scrolling text",
                  onGetErrorMessage: this._validateText.bind(this)
                }),
                PropertyPaneToggle('marqueeActive', {
                  label: 'Active',
                  checked: true
                }),
                PropertyPaneDropdown('placement', {
                  label: strings.PlacementFieldLabel,
                  options: [
                    { key: 'top', text: strings.PlacementTopOption },
                    { key: 'aboveChrome', text: strings.PlacementAboveChromeOption },
                    { key: 'belowChrome', text: strings.PlacementBelowChromeOption },
                    { key: 'aboveContent', text: strings.PlacementAboveContentOption }
                  ],
                  selectedKey: this.properties.placement || 'top'
                })
              ]
            },
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneDropdown('listName', {
                  label: strings.ListNameFieldLabel,
                  options: listOptions,
                  selectedKey: this.properties.listName
                }),
                PropertyPaneDropdown('viewId', {
                  label: strings.ViewFieldLabel,
                  options: viewOptions,
                  selectedKey: this.properties.viewId,
                  disabled: !this.properties.listName
                }),
                PropertyPaneDropdown('messageField', {
                  label: strings.MessageFieldLabel,
                  options: listFieldOptions,
                  selectedKey: this.properties.messageField,
                  disabled: !this.properties.listName || !this.properties.viewId
                }),
              ]
            },
            {
              groupName: strings.FontGroupName,
              groupFields: [
                PropertyPaneDropdown('fontFamily', {
                  label: strings.FontFamilyFieldLabel,
                  options: fontFamilyOptions,
                  selectedKey: this.properties.fontFamily || '"Segoe UI", sans-serif'
                }),
                PropertyPaneDropdown('fontSize', {
                  label: strings.FontSizeFieldLabel,
                  options: [
                    { key: '10px', text: '10px' },
                    { key: '12px', text: '12px' },
                    { key: '13px', text: '13px' },
                    { key: '14px', text: '14px' },
                    { key: '16px', text: '16px' },
                    { key: '18px', text: '18px' },
                    { key: '20px', text: '20px' },
                    { key: '24px', text: '24px' },
                    { key: '28px', text: '28px' },
                    { key: '32px', text: '32px' }
                  ],
                  selectedKey: this.properties.fontSize || '13px'
                }),
                PropertyPaneDropdown('fontStyle', {
                  label: strings.FontStyleFieldLabel,
                  options: fontStyleOptions,
                  selectedKey: this.properties.fontStyle || 'normal'
                }),
                PropertyPaneToggle('fontBold', {
                  label: strings.FontBoldFieldLabel,
                  checked: true
                }),
                PropertyFieldColorPicker('marqueeBackColor', {
                  label: "Announcements Background Color",
                  selectedColor: this.properties.marqueeBackColor || '#333333',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  disabled: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'marqueeBackColorField'
                }),
                PropertyFieldColorPicker('marqueeTextColor', {
                  label: "Announcements Text Color",
                  selectedColor: this.properties.marqueeTextColor || '#fcfcfc',
                  onPropertyChange: this.handleColorPropertyChange.bind(this),
                  properties: this.properties,
                  disabled: false,
                  style: PropertyFieldColorPickerStyle.Inline,
                  key: 'marqueeTextColorField'
                }),
                PropertyPaneTextField('marqueeHeight', {
                  label: strings.MarqueeHeightFieldLabel,
                  placeholder: '32px'
                })
              ]
            },
            {
              groupName: strings.ScrollGroupName,
              groupFields: [
                PropertyPaneSlider('scrollSpeed', {
                  label: strings.ScrollSpeedFieldLabel,
                  min: 5,
                  max: 120,
                  step: 1,
                  value: typeof this.properties.scrollSpeed === 'number' ? this.properties.scrollSpeed : 50
                }),
                PropertyPaneSlider('messageDuration', {
                  label: strings.MessageDurationFieldLabel,
                  min: 2,
                  max: 60,
                  step: 1,
                  value: typeof this.properties.messageDuration === 'number' ? this.properties.messageDuration : 8
                }),
                PropertyPaneDropdown('scrollDirection', {
                  label: strings.ScrollDirectionFieldLabel,
                  options: directionOptions,
                  selectedKey: this.properties.scrollDirection || 'left'
                })
              ]
            },
            {
              groupName: strings.DiagnosticsGroupName,
              groupFields: [
                PropertyPaneCheckbox('enableDiagnostics', {
                  text: strings.PropEnableDiagnosticsLabel,
                  checked: this.properties.enableDiagnostics !== false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
