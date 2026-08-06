import { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';
/**
 * @interface
 * Public properties of the PropertyFieldSearchPropertiesPicker custom field
 *
 */
export interface IPropertyFieldSearchPropertiesPickerProps {
    /**
     * @var
     * Property field label displayed on top
     */
    label: string;
    /**
     * @var
     * Default Selected properties
     */
    selectedProperties?: string[];
    /**
     * @var
     * Suggestions Header Text
     */
    suggestionsHeaderText: string;
    /**
     * @var
     * Text to display when no results found
     */
    noResultsFoundText: string;
    /**
     * @var
     * Text to display during loading
     */
    loadingText: string;
    /**
     * @function
     * Defines a onPropertyChange function to raise when the selected value changed.
     * Normally this function must be always defined with the 'this.onPropertyChange'
     * method of the web part object.
     */
    onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    /**
     * @function
     * This API is called to render the web part.
     * Normally this function must be always defined with the 'this.render.bind(this)'
     * method of the web part object.
     */
    render(): void;
    /**
     * This property is used to indicate the web part's PropertyPane interaction mode: Reactive or NonReactive.
     * The default behaviour is Reactive.
     */
    disableReactivePropertyChanges?: boolean;
    /**
     * @var
     * Parent Web Part properties
     */
    properties: any;
    /**
     * @var
     * An UNIQUE key indicates the identity of this control
     */
    key?: string;
    /**
     * Whether the property pane field is enabled or not.
     */
    disabled?: boolean;
    /**
     * The method is used to get the validation error message and determine whether the input value is valid or not.
     *
     *   When it returns string:
     *   - If valid, it returns empty string.
     *   - If invalid, it returns the error message string and the text field will
     *     show a red border and show an error message below the text field.
     *
     *   When it returns Promise<string>:
     *   - The resolved value is display as error message.
     *   - The rejected, the value is thrown away.
     *
     */
    onGetErrorMessage?: (value: string[]) => string | Promise<string>;
    /**
     * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
     * Default value is 200.
     */
    deferredValidationTime?: number;
}
/**
 * @interface
 * Private properties of the PropertyFieldSearchPropertiesPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldSearchPropertiesPicker.
 *
 */
export interface IPropertyFieldSearchPropertiesPickerPropsInternal extends IPropertyPaneCustomFieldProps {
    label: string;
    selectedProperties?: string[];
    suggestionsHeaderText: string;
    noResultsFoundText: string;
    loadingText: string;
    targetProperty: string;
    onRender(elem: HTMLElement): void;
    onDispose(elem: HTMLElement): void;
    onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    render(): void;
    disableReactivePropertyChanges?: boolean;
    properties: any;
    disabled?: boolean;
    onGetErrorMessage?: (value: string[]) => string | Promise<string>;
    deferredValidationTime?: number;
}
/**
 * @function
 * Helper method to create the customer field on the PropertyPane.
 * @param targetProperty - Target property the custom field is associated to.
 * @param properties - Strongly typed custom field properties.
 */
export declare function PropertyFieldSearchPropertiesPicker(targetProperty: string, properties: IPropertyFieldSearchPropertiesPickerProps): IPropertyPaneField<IPropertyFieldSearchPropertiesPickerPropsInternal>;
