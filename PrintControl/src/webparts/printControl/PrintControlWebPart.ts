import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';
import * as strings from 'PrintControlWebPartStrings';

const packageSolutionConfig: any = require('../../../config/package-solution.json');

export interface IPrintControlWebPartProps {
  buttonLabel?: string;
  alignment?: 'left' | 'center' | 'right';
  buttonTextColor?: string;
  buttonBackgroundColor?: string;
  buttonBorderColor?: string;
  buttonBorderWidth?: number;
  buttonFontFamily?: string;
  buttonFontSize?: number;
  buttonFontStyle?: string;
  buttonFontBold?: boolean;
  buttonCornerStyle?: 'square' | 'rounded';
  buttonCornerRadius?: number;
  contentSelector?: string;
  printCornerText?: string;
  enableDiagnostics?: boolean;
}

interface IDropdownOption {
  key: string;
  text: string;
}

export default class PrintControlWebPart extends BaseClientSideWebPart<IPrintControlWebPartProps> {
  private _printButton: HTMLButtonElement | undefined;

  public constructor() {
    super();
    this._printButton = undefined;
    this.handlePrint = this.handlePrint.bind(this);
  }

  private logDiagnostic(message: string): void {
    if (this.properties.enableDiagnostics === false) {
      return;
    }
    console.log('[PrintControlWebPart] ' + message);
  }

  private getWebPartVersion(): string {
    const solutionVersion = packageSolutionConfig && packageSolutionConfig.solution
      ? String(packageSolutionConfig.solution.version || '') : '';
    return solutionVersion || (this.context.manifest ? String(this.context.manifest.version || 'Unknown') : 'Unknown');
  }

  private getFontFamilyOptions(): IDropdownOption[] {
    return [
      { key: '', text: 'Default' },
      { key: 'Segoe UI', text: 'Segoe UI' },
      { key: 'Calibri', text: 'Calibri' },
      { key: 'Arial', text: 'Arial' },
      { key: 'Tahoma', text: 'Tahoma' },
      { key: 'Verdana', text: 'Verdana' },
      { key: 'Georgia', text: 'Georgia' },
      { key: '"Times New Roman"', text: 'Times New Roman' },
      { key: '"Courier New"', text: 'Courier New' }
    ];
  }

  private getFontStyleOptions(): IDropdownOption[] {
    return [
      { key: 'normal', text: 'Normal' },
      { key: 'italic', text: 'Italic' },
      { key: 'oblique', text: 'Oblique' },
      { key: 'initial', text: 'Initial' },
      { key: 'inherit', text: 'Inherit' },
      { key: 'unset', text: 'Unset' }
    ];
  }

  private getPrintableContent(): HTMLElement | undefined {
    const configuredSelector = String(this.properties.contentSelector || '').trim();
    if (configuredSelector) {
      try {
        const configuredTarget = document.querySelector(configuredSelector) as HTMLElement;
        if (configuredTarget) {
          this.logDiagnostic('Using configured content selector: ' + configuredSelector);
          return configuredTarget;
        }
      } catch (selectorError) {
        this.logDiagnostic('Configured content selector is invalid: ' + String(selectorError));
      }
    }

    const selectors = [
      '[data-automation-id="Canvas"]',
      '.SPCanvas',
      '.Canvas',
      'main[role="main"]',
      '#contentBox'
    ];
    for (let selectorIndex = 0; selectorIndex < selectors.length; selectorIndex += 1) {
      const target = document.querySelector(selectors[selectorIndex]) as HTMLElement;
      if (target && target.contains(this.domElement)) {
        this.logDiagnostic('Using automatic content selector: ' + selectors[selectorIndex]);
        return target;
      }
    }

    let ancestor = this.domElement.parentElement;
    while (ancestor && ancestor !== document.body) {
      if (ancestor.getAttribute('role') === 'main') {
        return ancestor;
      }
      ancestor = ancestor.parentElement;
    }
    return this.domElement.parentElement;
  }

  private copyCurrentFormValues(source: HTMLElement, clone: HTMLElement): void {
    const sourceControls = source.querySelectorAll('input, textarea, select');
    const cloneControls = clone.querySelectorAll('input, textarea, select');
    const controlCount = Math.min(sourceControls.length, cloneControls.length);
    for (let controlIndex = 0; controlIndex < controlCount; controlIndex += 1) {
      const sourceControl = sourceControls[controlIndex] as any;
      const cloneControl = cloneControls[controlIndex] as any;
      if (sourceControl.tagName === 'INPUT') {
        cloneControl.value = sourceControl.value;
        cloneControl.checked = sourceControl.checked;
      } else if (sourceControl.tagName === 'TEXTAREA') {
        cloneControl.value = sourceControl.value;
        cloneControl.textContent = sourceControl.value;
      } else if (sourceControl.tagName === 'SELECT') {
        cloneControl.value = sourceControl.value;
        for (let optionIndex = 0; optionIndex < cloneControl.options.length; optionIndex += 1) {
          cloneControl.options[optionIndex].selected = sourceControl.options[optionIndex].selected;
        }
      }
    }
  }

  private resetClonedFullWidthLayout(clone: HTMLElement): void {
    const fullWidthElements: HTMLElement[] = [];
    if (clone.hasAttribute('data-sps-full-width-owners')) {
      fullWidthElements.push(clone);
    }
    const ownedElements = clone.querySelectorAll('[data-sps-full-width-owners]');
    for (let elementIndex = 0; elementIndex < ownedElements.length; elementIndex += 1) {
      fullWidthElements.push(ownedElements[elementIndex] as HTMLElement);
    }
    for (let elementIndex = 0; elementIndex < fullWidthElements.length; elementIndex += 1) {
      const element = fullWidthElements[elementIndex];
      element.style.removeProperty('width');
      element.style.removeProperty('min-width');
      element.style.removeProperty('max-width');
      element.style.removeProperty('margin-left');
      element.style.removeProperty('margin-right');
      element.style.removeProperty('flex-basis');
      element.style.removeProperty('flex-grow');
      element.style.removeProperty('flex-shrink');
      element.style.removeProperty('grid-column');
      element.style.removeProperty('overflow-x');
      element.style.removeProperty('overflow-y');
    }
  }

  private hasFullWidthLayout(content: HTMLElement): boolean {
    return content.hasAttribute('data-sps-full-width-owners')
      || !!content.querySelector('[data-sps-full-width-owners]');
  }

  private getPrintPageName(): string {
    const pageHeading = document.querySelector('[data-automation-id="pageHeader"] h1,[data-automation-id="pageTitle"],h1[data-automation-id="pageTitle"]') as HTMLElement;
    const headingText = pageHeading ? String(pageHeading.textContent || '').trim() : '';
    const sourceTitle = String(document.title || '').trim();
    return headingText || (sourceTitle.toLowerCase() !== 'about:blank' ? sourceTitle : '') || 'Printed Record';
  }

  private escapePrintCssContent(value: string): string {
    return String(value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/[\r\n]+/g, ' ');
  }

  private handlePrint(): void {
    const sourceContent = this.getPrintableContent();
    if (!sourceContent) {
      this.logDiagnostic('No printable content area was found.');
      return;
    }

    const isFullWidth = this.hasFullWidthLayout(sourceContent);
    const printWindow = window.open('', '_blank', isFullWidth ? 'width=1400,height=900' : 'width=1100,height=800');
    if (!printWindow) {
      this.logDiagnostic('The browser blocked the print window.');
      return;
    }

    const printableContent = sourceContent.cloneNode(true) as HTMLElement;
    this.copyCurrentFormValues(sourceContent, printableContent);
    this.resetClonedFullWidthLayout(printableContent);
    const excludedSelector = '[data-sps-print-control="true"],button,[role="button"],[data-automation-id="CanvasControlMenu"],[data-automation-id="webPartHeader"]';
    const excludedNodes = printableContent.querySelectorAll(excludedSelector);
    for (let excludedIndex = excludedNodes.length - 1; excludedIndex >= 0; excludedIndex -= 1) {
      const excludedNode = excludedNodes[excludedIndex];
      const printControlZone = excludedNode.hasAttribute('data-sps-print-control')
        ? excludedNode.closest('[data-automation-id="CanvasControl"],.ControlZone') : undefined;
      const nodeToRemove = printControlZone || excludedNode;
      if (nodeToRemove.parentNode) {
        nodeToRemove.parentNode.removeChild(nodeToRemove);
      }
    }

    const printDocument = printWindow.document;
    const printPageName = this.getPrintPageName();
    const printCornerText = String(this.properties.printCornerText || '').trim() || printPageName + ' - Printed Record';
    const printDateTime = new Date().toLocaleString();
    printDocument.open();
    printDocument.write('<!doctype html><html><head></head><body></body></html>');
    printDocument.close();
    const titleElement = printDocument.createElement('title');
    titleElement.textContent = printPageName;
    printDocument.head.appendChild(titleElement);
    printDocument.title = printPageName;
    try {
      printWindow.history.replaceState({}, printPageName, document.location.href);
    } catch (historyError) {
      this.logDiagnostic('Unable to replace the temporary print URL: ' + String(historyError));
    }

    const baseElement = printDocument.createElement('base');
    baseElement.href = document.baseURI;
    printDocument.head.appendChild(baseElement);
    const styleNodes = document.querySelectorAll('link[rel="stylesheet"], style');
    for (let styleIndex = 0; styleIndex < styleNodes.length; styleIndex += 1) {
      printDocument.head.appendChild(styleNodes[styleIndex].cloneNode(true));
    }
    const printStyle = printDocument.createElement('style');
    printStyle.textContent = '@page{' + (isFullWidth ? 'size:landscape;' : '') + 'margin:18mm 12mm;@top-left{content:"' + this.escapePrintCssContent(printDateTime) + '";font:10px "Segoe UI",Arial,sans-serif;color:#444}@top-center{content:"' + this.escapePrintCssContent(printPageName) + '";font:10px "Segoe UI",Arial,sans-serif;color:#444}@bottom-left{content:"' + this.escapePrintCssContent(printCornerText) + '";font:10px "Segoe UI",Arial,sans-serif;color:#444}@bottom-right{content:"Page " counter(page) " of " counter(pages);font:10px "Segoe UI",Arial,sans-serif;color:#444}}html,body{width:100%!important;max-width:none!important;margin:0!important;padding:0!important;background:#fff!important}body{color:#000;print-color-adjust:exact;-webkit-print-color-adjust:exact}button,[role="button"],[data-sps-print-control="true"]{display:none!important}*{box-sizing:border-box}a{text-decoration:none;color:inherit}[data-automation-id="Canvas"],.SPCanvas,.Canvas,main[role="main"],#contentBox,.CanvasZone,.CanvasZoneContainer,.CanvasZoneSectionContainer,.CanvasSection,.ControlZone,[data-sps-full-width-owners]{width:100%!important;min-width:0!important;max-width:none!important;margin-left:0!important;margin-right:0!important}.sharePointDynamicForm,.container{width:100%!important;max-width:none!important}.CanvasZone,.ControlZone{break-inside:auto!important;page-break-inside:auto!important}';
    printDocument.head.appendChild(printStyle);
    printDocument.body.appendChild(printableContent);

    let printMediaWasActive = false;
    let printWindowClosed = false;
    const closePrintWindow = (): void => {
      if (printWindowClosed) {
        return;
      }
      printWindowClosed = true;
      window.setTimeout(() => {
        if (!printWindow.closed) {
          printWindow.close();
        }
        window.focus();
      }, 100);
    };
    printWindow.addEventListener('afterprint', closePrintWindow);
    const printMedia = printWindow.matchMedia ? printWindow.matchMedia('print') : undefined;
    if (printMedia && printMedia.addListener) {
      printMedia.addListener((event: any) => {
        if (event.matches) {
          printMediaWasActive = true;
        } else if (printMediaWasActive) {
          closePrintWindow();
        }
      });
    }

    this.logDiagnostic('Printable content cloned; opening browser print dialog in ' + (isFullWidth ? 'landscape' : 'browser default') + ' orientation.');
    window.setTimeout(() => {
      printWindow.focus();
      try {
        printWindow.print();
      } finally {
        closePrintWindow();
      }
    }, 500);
  }

  public render(): void {
    this.domElement.innerHTML = '';
    const root = document.createElement('div');
    root.setAttribute('data-sps-print-control', 'true');
    root.style.display = 'flex';
    root.style.justifyContent = this.properties.alignment === 'left' ? 'flex-start'
      : this.properties.alignment === 'center' ? 'center' : 'flex-end';
    root.style.width = '100%';

    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = String(this.properties.buttonLabel || 'Print');
    button.setAttribute('aria-label', String(this.properties.buttonLabel || 'Print'));
    button.style.backgroundColor = String(this.properties.buttonBackgroundColor || '#f0f0f0');
    button.style.color = String(this.properties.buttonTextColor || '#000000');
    button.style.borderColor = String(this.properties.buttonBorderColor || this.properties.buttonBackgroundColor || '#f0f0f0');
    button.style.borderWidth = String(this.properties.buttonBorderWidth === undefined ? 1 : Math.max(0, this.properties.buttonBorderWidth)) + 'px';
    button.style.borderStyle = 'solid';
    button.style.borderRadius = this.properties.buttonCornerStyle === 'rounded'
      ? String(Math.max(0, this.properties.buttonCornerRadius || 4)) + 'px' : '0';
    button.style.fontFamily = String(this.properties.buttonFontFamily || 'inherit');
    button.style.fontSize = String(Math.max(10, this.properties.buttonFontSize || 14)) + 'px';
    button.style.fontStyle = String(this.properties.buttonFontStyle || 'normal');
    button.style.fontWeight = this.properties.buttonFontBold ? 'bold' : 'normal';
    button.style.padding = '8px 18px';
    button.style.cursor = 'pointer';
    button.addEventListener('click', this.handlePrint);
    root.appendChild(button);
    this.domElement.appendChild(root);
    this._printButton = button;
  }

  protected onDispose(): void {
    if (this._printButton) {
      this._printButton.removeEventListener('click', this.handlePrint);
    }
    this._printButton = undefined;
    this.domElement.innerHTML = '';
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'buttonCornerStyle') {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const fontFamilyOptions = this.getFontFamilyOptions();
    const savedFontFamily = String(this.properties.buttonFontFamily || '');
    if (savedFontFamily && !fontFamilyOptions.some((option: IDropdownOption) => option.key === savedFontFamily)) {
      fontFamilyOptions.push({ key: savedFontFamily, text: savedFontFamily });
    }

    return {
      pages: [{
        header: { description: strings.PropertyPaneDescription },
        groups: [
          {
            groupName: 'Version: ' + this.getWebPartVersion(),
            groupFields: [PropertyPaneLabel('propertyPaneVersionInfo', { text: ' ' })]
          },
          {
            groupName: 'Print Button',
            groupFields: [
              PropertyPaneTextField('buttonLabel', { label: 'Button Label' }),
              PropertyPaneDropdown('alignment', { label: 'Alignment', options: [{ key: 'left', text: 'Left' }, { key: 'center', text: 'Center' }, { key: 'right', text: 'Right' }] }),
              PropertyFieldColorPicker('buttonTextColor', { label: 'Text Color', selectedColor: this.properties.buttonTextColor || '#000000', onPropertyChange: this.onPropertyPaneFieldChanged, properties: this.properties, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Inline, key: 'printButtonTextColor' }),
              PropertyFieldColorPicker('buttonBackgroundColor', { label: 'Background Color', selectedColor: this.properties.buttonBackgroundColor || '#f0f0f0', onPropertyChange: this.onPropertyPaneFieldChanged, properties: this.properties, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Inline, key: 'printButtonBackgroundColor' }),
              PropertyFieldColorPicker('buttonBorderColor', { label: 'Border Color', selectedColor: this.properties.buttonBorderColor || '#f0f0f0', onPropertyChange: this.onPropertyPaneFieldChanged, properties: this.properties, alphaSliderHidden: true, style: PropertyFieldColorPickerStyle.Inline, key: 'printButtonBorderColor' }),
              PropertyPaneSlider('buttonBorderWidth', { label: 'Border Width', min: 0, max: 10, step: 1, value: this.properties.buttonBorderWidth === undefined ? 1 : this.properties.buttonBorderWidth }),
              PropertyPaneDropdown('buttonFontFamily', { label: 'Font Family', options: fontFamilyOptions }),
              PropertyPaneSlider('buttonFontSize', { label: 'Font Size', min: 10, max: 40, step: 1, value: this.properties.buttonFontSize || 14 }),
              PropertyPaneDropdown('buttonFontStyle', { label: 'Font Style', options: this.getFontStyleOptions() }),
              PropertyPaneCheckbox('buttonFontBold', { text: 'Bold' }),
              PropertyPaneDropdown('buttonCornerStyle', { label: 'Corner Style', options: [{ key: 'square', text: 'Square' }, { key: 'rounded', text: 'Rounded' }] }),
              ...(this.properties.buttonCornerStyle === 'rounded' ? [PropertyPaneSlider('buttonCornerRadius', { label: 'Corner Radius', min: 0, max: 40, step: 1, value: this.properties.buttonCornerRadius || 4 })] : [])
            ]
          },
          {
            groupName: 'Printable Area',
            groupFields: [
              PropertyPaneTextField('contentSelector', { label: 'Content CSS Selector (Optional)', description: 'Leave blank to automatically print the SharePoint page canvas containing this control.' }),
              PropertyPaneTextField('printCornerText', { label: 'Printed Corner Text', description: 'Defaults to the SharePoint page name followed by Printed Record.' }),
              PropertyPaneLabel('browserPrintHeadersInfo', { text: 'Turn off Headers and footers in the browser print dialog to hide the browser date, URL, title, and page count and use only these print labels.' }),
              PropertyPaneCheckbox('enableDiagnostics', { text: 'Enable Diagnostics', checked: this.properties.enableDiagnostics !== false })
            ]
          }
        ]
      }]
    };
  }
}