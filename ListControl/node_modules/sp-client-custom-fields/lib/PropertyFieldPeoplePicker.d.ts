import { IPropertyPaneField, IPropertyPaneCustomFieldProps } from '@microsoft/sp-webpart-base';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
/**
 * @interface
 * Defines a People object for the PropertyFieldPeoplePicker
 *
 */
export interface IPropertyFieldPeople {
    /**
     * @var
     * User's full name
     */
    fullName: string;
    /**
     * @var
     * User's login
     */
    login: string;
    /**
     * @var
     * User's email (optional)
     */
    email?: string;
    /**
     * @var
     * User's job title (optional)
     */
    jobTitle?: string;
    /**
     * @var
     * User's initials (optional)
     */
    initials?: string;
    /**
     * @var
     * User's image url (optional)
     */
    imageUrl?: string;
}
/**
 * @interface
 * Public properties of the PropertyFieldPeoplePicker custom field
 *
 */
export interface IPropertyFieldPeoplePickerProps {
    /**
     * @var
     * Property field label
     */
    label: string;
    /**
     * @var
     * Web Part context
     */
    context: IWebPartContext;
    /**
     * @var
     * Intial data to load in the people picker (optional)
     */
    initialData?: IPropertyFieldPeople[];
    /**
     * @var
     * Defines if the People Picker allows to select duplicated users (optional)
     */
    allowDuplicate?: boolean;
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
    onGetErrorMessage?: (value: IPropertyFieldPeople[]) => string | Promise<string>;
    /**
     * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
     * Default value is 200.
     */
    deferredValidationTime?: number;
}
/**
 * @interface
 * Private properties of the PropertyFieldPeoplePicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldPeoplePicker.
 *
 */
export interface IPropertyFieldPeoplePickerPropsInternal extends IPropertyPaneCustomFieldProps {
    label: string;
    targetProperty: string;
    context: IWebPartContext;
    initialData?: IPropertyFieldPeople[];
    allowDuplicate?: boolean;
    onRender(elem: HTMLElement): void;
    onDispose(elem: HTMLElement): void;
    onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    render(): void;
    disableReactivePropertyChanges?: boolean;
    properties: any;
    onGetErrorMessage?: (value: IPropertyFieldPeople[]) => string | Promise<string>;
    deferredValidationTime?: number;
}
/**
 * @function
 * Helper method to create a People Picker on the PropertyPane.
 * @param targetProperty - Target property the people picker is associated to.
 * @param properties - Strongly typed people Picker properties.
 */
export declare function PropertyFieldPeoplePicker(targetProperty: string, properties: IPropertyFieldPeoplePickerProps): IPropertyPaneField<IPropertyFieldPeoplePickerPropsInternal>;
