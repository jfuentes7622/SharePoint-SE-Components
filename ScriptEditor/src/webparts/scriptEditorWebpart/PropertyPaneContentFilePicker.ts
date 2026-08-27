import {
  IPropertyPaneCustomFieldProps,
  IPropertyPaneField,
  PropertyPaneFieldType
} from '@microsoft/sp-webpart-base';

export interface IPropertyPaneContentFilePickerOptions {
  label: string;
  buttonText: string;
  accept: string;
  onFileLoaded: (fileName: string, content: string) => void;
}

export default class PropertyPaneContentFilePicker implements IPropertyPaneField<IPropertyPaneCustomFieldProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string = 'scriptFile';
  public properties: IPropertyPaneCustomFieldProps;

  private options: IPropertyPaneContentFilePickerOptions;
  private input: HTMLInputElement | undefined;
  private status: HTMLElement | undefined;

  constructor(options: IPropertyPaneContentFilePickerOptions) {
    this.options = options;
    this.properties = {
      key: 'scriptContentFilePicker',
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    };
  }

  private onRender(element: HTMLElement): void {
    this.onDispose();
    element.innerHTML = '';

    const label = document.createElement('div');
    label.textContent = this.options.label;
    label.style.fontSize = '14px';
    label.style.marginBottom = '8px';
    element.appendChild(label);

    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'ms-Button ms-Button--default';
    button.textContent = this.options.buttonText;
    button.style.minWidth = '100px';
    element.appendChild(button);

    this.input = document.createElement('input');
    this.input.type = 'file';
    this.input.accept = this.options.accept;
    this.input.style.display = 'none';
    this.input.addEventListener('change', this.onFileChanged);
    element.appendChild(this.input);

    this.status = document.createElement('div');
    this.status.setAttribute('aria-live', 'polite');
    this.status.style.fontSize = '12px';
    this.status.style.marginTop = '6px';
    this.status.style.color = '#605e5c';
    element.appendChild(this.status);

    button.addEventListener('click', this.openFilePicker);
  }

  private openFilePicker = (): void => {
    if (this.input) {
      this.input.click();
    }
  }

  private onFileChanged = (): void => {
    if (!this.input || !this.input.files || this.input.files.length === 0) {
      return;
    }

    const file = this.input.files[0];
    const reader = new FileReader();
    reader.onload = (): void => {
      const content = typeof reader.result === 'string' ? reader.result : '';
      this.options.onFileLoaded(file.name, content);
      if (this.status) {
        this.status.textContent = 'Loaded ' + file.name;
      }
    };
    reader.onerror = (): void => {
      if (this.status) {
        this.status.textContent = 'Unable to read ' + file.name;
      }
    };
    reader.readAsText(file);
  }

  private onDispose(): void {
    if (this.input) {
      this.input.removeEventListener('change', this.onFileChanged);
    }
    this.input = undefined;
    this.status = undefined;
  }
}
