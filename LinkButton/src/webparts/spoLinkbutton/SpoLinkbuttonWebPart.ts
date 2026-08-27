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
  PropertyPaneLabel,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import { PropertyFieldPicturePicker } from 'sp-client-custom-fields/lib/PropertyFieldPicturePicker';
import styles from './SpoLinkbuttonWebPart.module.scss';
import * as strings from 'SpoLinkbuttonWebPartStrings';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface ILinkButtonWebPartProps {
  description: string;
  Link: string;
  linkTarget: string;
  Align: string;
  buttonShape: string;
  buttonCornerRadius: number;
  fontSize: number | string;
  buttonBackgroundColor: string;
  buttonFontColor: string;
  overrideCssUrl: string;
  fontFamily: string;
  fontStyle: string;
  fontWeight: string;
  fontBold: boolean;
  textDecoration: string;
  textTransform: string;
  showIcon: boolean;
  iconImageUrl: string;
  imageAsBackground: boolean;
  imageDisplayMode: string;
  imageOpacity: number;
  iconSize: number | string;
  iconPosition: string;
  buttonHeight: number | string;
  buttonWidth: number | string;
  enableDiagnostics: boolean;
}

require ('./button.css');

export default class LinkButtonWebPart extends BaseClientSideWebPart<ILinkButtonWebPartProps> {

  private _overrideStylesheetId: string;

  public constructor() {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

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

  private toCssPixels(value: number | string, fallback: number): string {
    if (typeof value === 'number' && isFinite(value)) {
      return String(value) + 'px';
    }
    const savedValue = String(value || '').trim();
    return savedValue || String(fallback) + 'px';
  }

  private toSliderNumber(value: number | string, fallback: number): number {
    if (typeof value === 'number' && isFinite(value)) {
      return value;
    }
    const savedValue = String(value || '').trim().toLowerCase();
    const parsedValue = parseFloat(savedValue);
    if (!isFinite(parsedValue)) {
      return fallback;
    }
    if (savedValue.indexOf('rem') > -1 || savedValue.indexOf('em') > -1) {
      return Math.round(parsedValue * 16);
    }
    if (savedValue.indexOf('vh') > -1) {
      return Math.round(window.innerHeight * parsedValue / 100);
    }
    if (savedValue.indexOf('vw') > -1) {
      return Math.round(window.innerWidth * parsedValue / 100);
    }
    return Math.round(parsedValue);
  }

  private getImageDisplayMode(): string {
    if (this.properties.imageDisplayMode) {
      return this.properties.imageDisplayMode;
    }
    return this.properties.imageAsBackground === true ? 'background-text' : 'side';
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
    if (propertyPath === 'showIcon' || propertyPath === 'buttonShape' || propertyPath === 'imageDisplayMode') {
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
      const roundedRadius = typeof this.properties.buttonCornerRadius === 'number' ? this.properties.buttonCornerRadius : 5;
      const borderRadius = shape === 'square' ? '0' : shape === 'round' ? '9999px' : String(roundedRadius) + 'px';

      // Typed as `any`: TS 2.4.2 cannot narrow the fontWeight ternary to the CSSProperties literal union.
      const buttonStyle: any = {
        backgroundColor: this.properties.buttonBackgroundColor || '#8A1717',
        color: this.properties.buttonFontColor || '#ffffff',
        height: this.toCssPixels(this.properties.buttonHeight, 50),
        width: this.toCssPixels(this.properties.buttonWidth, 200),
        fontSize: this.toCssPixels(this.properties.fontSize, 20),
        fontFamily: this.properties.fontFamily || "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif",
        fontStyle: this.properties.fontStyle || 'normal',
        fontWeight: this.properties.fontWeight || (this.properties.fontBold ? '700' : '400'),
        textDecoration: this.properties.textDecoration || 'none',
        textTransform: this.properties.textTransform || 'none',
        borderRadius: borderRadius
      };

      const buttonText = (this.properties.description || '').replace(/[<>]/g, "");
      const iconImageUrl = this.properties.showIcon ? (this.properties.iconImageUrl || '').replace(/["'<>]/g, '') : '';
      const iconSize = this.toCssPixels(this.properties.iconSize, 24);
      const imageDisplayMode = this.getImageDisplayMode();
      const imageAsBackground = imageDisplayMode === 'background-text' || imageDisplayMode === 'background-only';
      const showButtonText = imageDisplayMode !== 'background-only';
      const imageOpacity = typeof this.properties.imageOpacity === 'number'
        ? Math.max(0, Math.min(100, this.properties.imageOpacity)) / 100
        : 1;
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

      const sideImageElement = iconImageUrl && !imageAsBackground ? React.createElement(
        'div',
        {
          className: 'button-icon-container',
          style: { width: iconSize, height: '100%', opacity: imageOpacity }
        },
        React.createElement('img', {
          className: 'button-icon',
          src: iconImageUrl,
          alt: '',
          'aria-hidden': true,
          style: { width: '100%', height: '100%' }
        })
      ) : undefined;

      const backgroundImageElement = iconImageUrl && imageAsBackground ? React.createElement('img', {
        className: 'button-background-image',
        src: iconImageUrl,
        alt: '',
        'aria-hidden': true,
        style: { opacity: imageOpacity }
      }) : undefined;

      const buttonChildren = backgroundImageElement
        ? (showButtonText ? [backgroundImageElement, textElement] : [backgroundImageElement])
        : sideImageElement
          ? (iconOnRight ? [textElement, sideImageElement] : [sideImageElement, textElement])
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
              className: 'button' + (sideImageElement ? ' button-with-side-image' : '') + (backgroundImageElement ? ' button-with-background-image' : ''),
              style: mergedStyle,
              href: this.properties.Link || '#',
              target: this.properties.linkTarget || '_self',
              rel: this.properties.linkTarget === '_blank' ? 'noopener noreferrer' : undefined,
              'aria-label': buttonText
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
              groupFields: ([
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
                PropertyPaneDropdown('linkTarget', {
                  label: 'Link Target',
                  selectedKey: this.properties.linkTarget || '_self',
                  options: [
                    { key: '_self', text: 'Current window' },
                    { key: '_blank', text: 'New tab' },
                    { key: '_parent', text: 'Parent frame' },
                    { key: '_top', text: 'Top frame' }
                  ]
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
                })
              ] as any[]).concat((this.properties.buttonShape || 'rounded') === 'rounded' ? [
                PropertyPaneSlider('buttonCornerRadius', {
                  label: 'Corner Radius (px)',
                  min: 0,
                  max: 60,
                  step: 1,
                  value: typeof this.properties.buttonCornerRadius === 'number' ? this.properties.buttonCornerRadius : 5,
                  showValue: true
                })
              ] : []).concat([
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
                    { key: "'Courier New', Courier, monospace", text: 'Courier New' },
                    { key: 'Verdana, Geneva, sans-serif', text: 'Verdana' },
                    { key: "'Arial Black', Gadget, sans-serif", text: 'Arial Black' },
                    { key: "'Palatino Linotype', 'Book Antiqua', Palatino, serif", text: 'Palatino' },
                    { key: 'Impact, Charcoal, sans-serif', text: 'Impact' }
                  ]
                }),
                PropertyPaneDropdown('fontStyle', {
                  label: 'Font Style',
                  options: [
                    { key: 'normal', text: 'Normal' },
                    { key: 'italic', text: 'Italic' },
                    { key: 'oblique', text: 'Oblique' }
                  ],
                  selectedKey: this.properties.fontStyle || 'normal'
                }),
                PropertyPaneDropdown('fontWeight', {
                  label: 'Font Weight',
                  options: [
                    { key: '300', text: 'Light' },
                    { key: '400', text: 'Regular' },
                    { key: '500', text: 'Medium' },
                    { key: '600', text: 'Semi-bold' },
                    { key: '700', text: 'Bold' },
                    { key: '800', text: 'Extra bold' }
                  ],
                  selectedKey: this.properties.fontWeight || (this.properties.fontBold ? '700' : '400')
                }),
                PropertyPaneDropdown('textDecoration', {
                  label: 'Text Decoration',
                  options: [
                    { key: 'none', text: 'None' },
                    { key: 'underline', text: 'Underline' },
                    { key: 'line-through', text: 'Strikethrough' }
                  ],
                  selectedKey: this.properties.textDecoration || 'none'
                }),
                PropertyPaneDropdown('textTransform', {
                  label: 'Text Case',
                  options: [
                    { key: 'none', text: 'As entered' },
                    { key: 'uppercase', text: 'Uppercase' },
                    { key: 'lowercase', text: 'Lowercase' },
                    { key: 'capitalize', text: 'Capitalize' }
                  ],
                  selectedKey: this.properties.textTransform || 'none'
                }),
                PropertyPaneSlider('fontSize', {
                  label: 'Font Size (px)',
                  min: 8,
                  max: 72,
                  step: 1,
                  value: this.toSliderNumber(this.properties.fontSize, 20),
                  showValue: true
                }),
                PropertyPaneSlider('buttonHeight', {
                  label: 'Button Height (px)',
                  min: 24,
                  max: 200,
                  step: 1,
                  value: this.toSliderNumber(this.properties.buttonHeight, 50),
                  showValue: true
                }),
                PropertyPaneSlider('buttonWidth', {
                  label: 'Button Width (px)',
                  min: 60,
                  max: 600,
                  step: 5,
                  value: this.toSliderNumber(this.properties.buttonWidth, 200),
                  showValue: true
                })
              ]) as any[]
            },
            {
              groupName: 'Icon',
              groupFields: [
                PropertyPaneCheckbox('showIcon', {
                  text: 'Set Icon'
                })
              ].concat(this.properties.showIcon ? [
                PropertyFieldPicturePicker('iconImageUrl', {
                  label: 'Icon Image',
                  initialValue: this.properties.iconImageUrl || '',
                  context: this.context,
                  previewImage: true,
                  allowedFileExtensions: '.gif,.jpg,.jpeg,.bmp,.png,.svg,.webp',
                  readOnly: false,
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  onGetErrorMessage: this._validateIconUrl.bind(this),
                  render: this.render.bind(this),
                  key: 'iconImagePicker'
                }),
                PropertyPaneDropdown('imageDisplayMode', {
                  label: 'Image Display',
                  selectedKey: this.getImageDisplayMode(),
                  options: [
                    { key: 'side', text: 'Side image' },
                    { key: 'background-text', text: 'Image as Background with Text' },
                    { key: 'background-only', text: 'Image as Background without Text' }
                  ]
                }),
                PropertyPaneSlider('imageOpacity', {
                  label: 'Image Opacity (%)',
                  min: 0,
                  max: 100,
                  step: 1,
                  value: typeof this.properties.imageOpacity === 'number' ? this.properties.imageOpacity : 100,
                  showValue: true
                })
              ].concat(this.getImageDisplayMode() === 'side' ? [
                PropertyPaneSlider('iconSize', {
                  label: 'Side Image Width (px)',
                  min: 8,
                  max: 128,
                  step: 1,
                  value: this.toSliderNumber(this.properties.iconSize, 24),
                  showValue: true
                }),
                PropertyPaneDropdown('iconPosition', {
                  label: 'Icon Position',
                  options: [
                    { key: 'left', text: 'Left' },
                    { key: 'right', text: 'Right' }
                  ]
                })
              ] : []) : []) as any[]
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
