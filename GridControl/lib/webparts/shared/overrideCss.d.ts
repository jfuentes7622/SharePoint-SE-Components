import { IPropertyPaneCustomFieldProps, IPropertyPaneField, WebPartContext } from '@microsoft/sp-webpart-base';
export declare function applyOverrideCss(value: string, instanceId: string): void;
export declare function PropertyPaneOverrideCss(targetProperty: string, currentValue: string, context: WebPartContext, onChange: (newValue: string) => void): IPropertyPaneField<IPropertyPaneCustomFieldProps>;
