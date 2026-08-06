import * as React from 'react';
import * as ReactDom from 'react-dom';
//import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';

import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar';
import { ICalendarProps } from './components/ICalendarProps';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IDropdownOption {
  key: string;
  text: string;
}

export interface ICalendarWebPartProps {
  title: string;
  listName: string;
  defaultView: string;
  showWeekends: boolean;
  calendarHeight: number;
  enableDiagnostics: boolean;
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {
  private _lists: IDropdownOption[] = [];

  public render(): void {
    this.logDiagnostic('render() called. listName=' + String(this.properties.listName || '(none)'));
    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        spfxContext: this.context,
        listName: this.properties.listName || '',
        defaultView: this.properties.defaultView || 'dayGridMonth',
        showWeekends: this.properties.showWeekends !== false,
        calendarHeight: typeof this.properties.calendarHeight === 'number' ? this.properties.calendarHeight : 600,
        enableDiagnostics: this.properties.enableDiagnostics !== false,
        fUpdateProperty: (value: string) => {
          this.properties.title = value;
        },
        fPropertyPaneOpen: this.context.propertyPane.open
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.logDiagnostic('onInit started. listName=' + String(this.properties.listName || '(none)'));
    return this.loadLists();
  }

  protected onDispose(): void {
    this.logDiagnostic('onDispose called.');
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  protected onPropertyPaneConfigurationStart(): void {
    this.logDiagnostic('Property pane opened. listName=' + String(this.properties.listName || '(none)'));
    if (this._lists.length === 0) {
      this.loadLists();
    }
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }

    console.log('[CalendarWebPart] ' + message);
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
      this.logDiagnostic('Loading Events lists from web URL: ' + webUrl);
      const data = await this.getJsonWithAcceptFallback(
        webUrl + "/_api/web/lists?$select=Title,Hidden,BaseTemplate&$filter=Hidden eq false and BaseTemplate eq 106"
      );
      const lists = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);

      this._lists = lists.map((list: any) => {
        return { key: list.Title, text: list.Title };
      });
      this.logDiagnostic('Loaded Events lists successfully. Count=' + String(this._lists.length));
      this.context.propertyPane.refresh();
    } catch (error) {
      this.logDiagnostic('Failed to load lists: ' + (error && error.message ? error.message : String(error)));
      this._lists = [];
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const viewOptions: IPropertyPaneDropdownOption[] = [
      { key: 'dayGridMonth', text: strings.ViewMonthLabel },
      { key: 'timeGridWeek', text: strings.ViewWeekLabel },
      { key: 'timeGridDay', text: strings.ViewDayLabel },
      { key: 'listWeek', text: strings.ViewListLabel }
    ];
    const listOptions: IPropertyPaneDropdownOption[] = this._lists.map((list) => {
      return { key: list.key, text: list.text };
    });

    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: 'Version: ' + this.getWebPartVersion(),
              groupFields: [
                PropertyPaneLabel('propertyPaneVersionInfo', {
                  text: ' '
                })
              ]
            },
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
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
                PropertyPaneDropdown('defaultView', {
                  label: strings.DefaultViewFieldLabel,
                  options: viewOptions,
                  selectedKey: this.properties.defaultView || 'dayGridMonth'
                }),
                PropertyPaneToggle('showWeekends', {
                  label: strings.ShowWeekendsFieldLabel,
                  checked: this.properties.showWeekends !== false
                }),
                PropertyPaneSlider('calendarHeight', {
                  label: strings.CalendarHeightFieldLabel,
                  min: 300,
                  max: 1200,
                  step: 20,
                  value: typeof this.properties.calendarHeight === 'number' ? this.properties.calendarHeight : 600
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
