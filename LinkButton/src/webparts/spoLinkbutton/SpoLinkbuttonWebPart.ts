import * as React from 'react';
import * as ReactDOM from 'react-dom';
// protected get dataVersion(): Version {} intentionally omitted -- BaseClientSideWebPart in
// SPFx 1.5.1 does not support overriding this accessor cleanly (see repo memory notes).
import {
  BaseClientSideWebPart,
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import styles from './SpoLinkbuttonWebPart.module.scss';
import * as strings from 'SpoLinkbuttonWebPartStrings';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface ILinkButtonWebPartProps {
  description: string;
  Link: string;
  Align: string;
  buttonShape: string;
  fontSize: string;
  buttonBackgroundColor: string;
  buttonFontColor: string;
  overrideCssUrl: string;
  fontFamily: string;
  fontStyle: string;
  fontBold: boolean;
  showIcon: boolean;
  iconImageUrl: string;
  iconSize: string;
  iconPosition: string;
  buttonHeight: string;
  buttonWidth: string;
  enableDiagnostics: boolean;
}

require ('./button.css');

export default class LinkButtonWebPart extends BaseClientSideWebPart<ILinkButtonWebPartProps> {

  private _overrideStylesheetId: string;

  protected onInit(): Promise<void> {
    this._overrideStylesheetId = 'linkbutton-override-css-' + this.instanceId;
    this.logDiagnostic('onInit started. instanceId=' + String(this.instanceId));
    return super.onInit().then(() => {
      this.logDiagnostic('onInit completed.');
    });
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }
    console.log('[LinkButtonWebPart] ' + message);
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
    if (!value || value === null || value.trim().length === 0) {
      this.logDiagnostic('Validation failed: description is empty.');
      return 'Provide Description';
    }
    if (/<[^>]*>/.test(value)) {
      this.logDiagnostic('Validation failed: description contains markup.');
      return "Text Required, please try again!";
    }
    if (value.length > 50) {
      this.logDiagnostic('Validation failed: description exceeds 50 characters.');
      return 'Description should not be longer than 50 characters';
    }
    // If validation is successful, return an empty string.
    return "";
  }

  private _validateUrl(value: string): string {
    // If validation is not successful, return a string with error message.
    if (!value || value === null || value.trim().length === 0) {
      this.logDiagnostic('Validation failed: link URL is empty.');
      return 'Provide URL';
    }
    if (/<[^>]*>/.test(value)) {
      this.logDiagnostic('Validation failed: link URL contains markup.');
      return "Valid URL required, please try again!";
    }
    // If validation is successful, return an empty string.
    return "";
  }

  private _validateIconUrl(value: string): string {
    if (!value) { return ""; }
    if (/<[^>]*>/.test(value) || /["']/.test(value)) {
      this.logDiagnostic('Validation failed: icon URL is invalid.');
      return "Invalid icon URL";
    }
    return "";
  }

  private _validateCssUrl(value: string): string {
    if (!value) { return ""; }
    const sanitized = value.trim();
    if (/["'<>]/.test(sanitized)) {
      this.logDiagnostic('Validation failed: override CSS URL is invalid.');
      return 'Invalid CSS URL';
    }
    if (!/^https?:\/\//i.test(sanitized) && sanitized.indexOf('/') !== 0) {
      this.logDiagnostic('Validation failed: override CSS URL is not absolute or site-relative.');
      return 'Use an absolute URL (http/https) or site-relative path (/...)';
    }
    return '';
  }

  private _applyOverrideStylesheet(): void {
    const cssUrl = (this.properties.overrideCssUrl || '').trim();
    const existing = document.getElementById(this._overrideStylesheetId) as HTMLLinkElement;

    if (!cssUrl) {
      if (existing) {
        this.logDiagnostic('Removing override stylesheet (overrideCssUrl cleared).');
        existing.parentNode.removeChild(existing);
      }
      return;
    }

    let resolvedCssUrl = cssUrl;
    try {
      resolvedCssUrl = new URL(cssUrl, window.location.href).href;
    } catch (urlError) {
      this.logDiagnostic('Could not normalize override CSS URL, using as-is: ' + String(urlError));
    }

    if (existing) {
      const currentRawUrl = existing.getAttribute('data-source-href') || '';
      if (currentRawUrl !== cssUrl) {
        this.logDiagnostic('Updating override stylesheet href to: ' + cssUrl);
        existing.setAttribute('data-source-href', cssUrl);
        existing.href = resolvedCssUrl;
      }
      return;
    }

    this.logDiagnostic('Injecting override stylesheet: ' + cssUrl);
    const link = document.createElement('link');
    link.id = this._overrideStylesheetId;
    link.rel = 'stylesheet';
    link.type = 'text/css';
    link.setAttribute('data-source-href', cssUrl);
    link.href = resolvedCssUrl;
    document.head.appendChild(link);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
    if (propertyPath === 'showIcon') {
      this.context.propertyPane.refresh();
    }
  }

  public render(): void {
    let element: React.ReactElement<{}>;
    this._applyOverrideStylesheet();

    // Show placeholder only when required content is missing.
    // This keeps live visual updates visible while editing properties.
    if (!this.properties.description || this.properties.description.trim() === '') {
      this.logDiagnostic('Rendering placeholder; description is not yet configured.');
      element = React.createElement(
        'div',
        { className: styles.linkButton },
        React.createElement(
          'div',
          { className: styles.container },
          React.createElement(
            'div',
            { style: { padding: '20px', textAlign: 'center' } },
            React.createElement('h3', {}, 'Link Button'),
            React.createElement('p', {}, 'Configure your link button properties. Add a title, URL, and customize appearance.'),
            React.createElement(
              'button',
              {
                onClick: () => this.context.propertyPane.open(),
                style: { padding: '10px 20px', marginTop: '10px' }
              },
              'Configure Button'
            )
          )
        )
      );
    } else {
      this.logDiagnostic('Rendering button. description=' + this.properties.description + ', link=' + String(this.properties.Link));
      const shape = (this.properties.buttonShape || 'rounded').toLowerCase();
      const borderRadius = shape === 'square' ? '0' : shape === 'round' ? '9999px' : '5px';

      // Typed as `any`: TS 2.4.2 cannot narrow the fontWeight ternary to the CSSProperties literal union.
      const buttonStyle: any = {
        backgroundColor: this.properties.buttonBackgroundColor || '#8A1717',
        color: this.properties.buttonFontColor || '#ffffff',
        height: this.properties.buttonHeight || '5vh',
        width: this.properties.buttonWidth || '200px',
        fontSize: this.properties.fontSize || '1.25rem',
        fontFamily: this.properties.fontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
        fontStyle: this.properties.fontStyle || 'normal',
        fontWeight: this.properties.fontBold ? 'bold' : 'normal',
        borderRadius: borderRadius
      };

      const buttonText = (this.properties.description || '').replace(/[<>]/g, "");
      const iconImageUrl = this.properties.showIcon ? (this.properties.iconImageUrl || '').replace(/["'<>]/g, '') : '';
      const iconSize = (this.properties.iconSize || '24px').trim();
      const iconOnRight = this.properties.iconPosition === 'right';
      const marginStyle = this.properties.Align === '3' ? { marginLeft: 'auto' } :
                         this.properties.Align === '2' ? { margin: 'auto' } : {};

      if (iconImageUrl) {
        this.logDiagnostic('Rendering with icon. position=' + String(this.properties.iconPosition) + ', size=' + iconSize);
      }

      const textElement = React.createElement(
        'div',
        { className: 'button-text-container' },
        buttonText
      );

      const iconElement = iconImageUrl ? React.createElement(
        'div',
        {
          className: 'button-icon-container',
          style: { width: iconSize, height: iconSize }
        },
        React.createElement('img', {
          className: 'button-icon',
          src: iconImageUrl,
          alt: '',
          'aria-hidden': true,
          style: { width: '100%', height: '100%' }
        })
      ) : undefined;

      const buttonChildren = iconElement
        ? (iconOnRight ? [textElement, iconElement] : [iconElement, textElement])
        : [textElement];

      const mergedStyle: any = {};
      for (const key in buttonStyle) { if (buttonStyle.hasOwnProperty(key)) { mergedStyle[key] = (buttonStyle as any)[key]; } }
      for (const key in marginStyle) { if (marginStyle.hasOwnProperty(key)) { mergedStyle[key] = (marginStyle as any)[key]; } }

      element = React.createElement(
        'div',
        { className: styles.linkButton },
        React.createElement(
          'div',
          { className: styles.container },
          React.createElement(
            'a',
            {
              className: 'button',
              style: mergedStyle,
              href: this.properties.Link || '#'
            },
            ...buttonChildren
          )
        )
      );
    }

    ReactDOM.render(element, this.domElement);
  }

  protected onDispose(): void {
    this.logDiagnostic('onDispose called; cleaning up React tree and override stylesheet.');
    const existing = document.getElementById(this._overrideStylesheetId);
    if (existing) {
      existing.parentNode.removeChild(existing);
    }
    ReactDOM.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.logDiagnostic('Property pane opened.');
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
              groupName: 'Button Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Button Title",
                  placeholder: "Go to...",
                  onGetErrorMessage: this._validateText.bind(this)
                }),
                PropertyPaneTextField('Link', {
                  label: "Link",
                  placeholder: "https://",
                  onGetErrorMessage: this._validateUrl.bind(this)
                }),
                PropertyPaneDropdown('Align', {
                  label: 'Align Button',
                  options: [
                    { key: '1', text: 'Left' },
                    { key: '2', text: 'Center' },
                    { key: '3', text: 'Right' }
                  ]
                }),
                PropertyPaneDropdown('buttonShape', {
                  label: 'Button Shape',
                  options: [
                    { key: 'rounded', text: 'Rounded corners' },
                    { key: 'square', text: 'Square' },
                    { key: 'round', text: 'Round' }
                  ]
                }),
                PropertyFieldColorPicker('buttonBackgroundColor', {
                  label: 'Button Background Color',
                  selectedColor: this.properties.buttonBackgroundColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'BackgroundColor',
                  properties: this.properties,
                  key: 'buttonBackgroundColorField'
                }),
                PropertyFieldColorPicker('buttonFontColor', {
                  label: 'Button Font Color',
                  selectedColor: this.properties.buttonFontColor,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  alphaSliderHidden: true,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'FontColor',
                  properties: this.properties,
                  key: 'buttonFontColorField'
                }),
                PropertyPaneTextField('overrideCssUrl', {
                  label: 'Override CSS URL',
                  placeholder: 'https://contoso.com/styles/linkbutton-overrides.css or /sites/siteassets/linkbutton-overrides.css',
                  onGetErrorMessage: this._validateCssUrl.bind(this),
                  description: 'Optional stylesheet loaded after component CSS to standardize styles across web parts.'
                }),
                PropertyPaneDropdown('fontFamily', {
                  label: 'Font Family',
                  options: [
                    { key: "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif", text: 'Segoe UI' },
                    { key: 'Arial, Helvetica, sans-serif', text: 'Arial' },
                    { key: "'Trebuchet MS', Helvetica, sans-serif", text: 'Trebuchet MS' },
                    { key: "'Georgia', serif", text: 'Georgia' },
                    { key: "'Times New Roman', Times, serif", text: 'Times New Roman' },
                    { key: "'Courier New', Courier, monospace", text: 'Courier New' }
                  ]
                }),
                PropertyPaneDropdown('fontStyle', {
                  label: 'Font Style',
                  options: [
                    { key: 'normal', text: 'Normal' },
                    { key: 'italic', text: 'Italic' },
                    { key: 'oblique', text: 'Oblique' }
                  ]
                }),
                PropertyPaneTextField('fontSize', {
                  label: 'Font Size',
                  placeholder: '16px, 1rem, 120%, etc.'
                }),
                PropertyPaneCheckbox('fontBold', {
                  text: 'Bold Font'
                }),
                PropertyPaneTextField('buttonHeight', {
                  label: 'Button Height',
                  placeholder: '5vh, 50px, 3rem, etc.'
                }),
                PropertyPaneTextField('buttonWidth', {
                  label: 'Button Width',
                  placeholder: '200px, 20vw, 10rem, etc.'
                })
              ]
            },
            {
              groupName: 'Icon',
              groupFields: [
                PropertyPaneCheckbox('showIcon', {
                  text: 'Set Icon'
                })
              ].concat(this.properties.showIcon ? [
                PropertyPaneTextField('iconImageUrl', {
                  label: 'Icon Image URL',
                  placeholder: 'https://example.com/icon.png',
                  onGetErrorMessage: this._validateIconUrl.bind(this)
                }),
                PropertyPaneTextField('iconSize', {
                  label: 'Icon Size',
                  placeholder: '24px, 2rem, etc.'
                }),
                PropertyPaneDropdown('iconPosition', {
                  label: 'Icon Position',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'right', text: 'Right' }
                  ]
                })
              ] : [])
            },
            {
              groupName: 'Diagnostics',
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
