import { IPropertyPaneField, IWebPartContext } from '@microsoft/sp-webpart-base';
/**
 * @interface
 * Generic Term Object (abstract interface)
 */
export interface ISPTermObject {
    Name: string;
    Guid: string;
    Identity: string;
    leaf: boolean;
    children?: ISPTermObject[];
    collapsed?: boolean;
    type: string;
}
/**
 * @interface
 * Defines a SharePoint Term Store
 */
export interface ISPTermStore extends ISPTermObject {
    IsOnline: boolean;
    WorkingLanguage: string;
    DefaultLanguage: string;
    Languages: string[];
}
/**
 * @interface
 * Defines an array of Term Stores
 */
export interface ISPTermStores extends Array<ISPTermStore> {
}
/**
 * @interface
 * Defines a Term Store Group of term sets
 */
export interface ISPTermGroup extends ISPTermObject {
    IsSiteCollectionGroup: boolean;
    IsSystemGroup: boolean;
    CreatedDate: string;
    LastModifiedDate: string;
}
/**
 * @interface
 * Array of Term Groups
 */
export interface ISPTermGroups extends Array<ISPTermGroup> {
}
/**
 * @interface
 * Defines a Term Set
 */
export interface ISPTermSet extends ISPTermObject {
    CustomSortOrder: string;
    IsAvailableForTagging: boolean;
    Owner: string;
    Contact: string;
    Description: string;
    IsOpenForTermCreation: boolean;
    TermStoreGuid: string;
}
/**
 * @interface
 * Array of Term Sets
 */
export interface ISPTermSets extends Array<ISPTermSet> {
}
/**
 * @interface
 * Public properties of the PropertyFieldTermSetPicker custom field
 *
 */
export interface IPropertyFieldTermSetPickerProps {
    /**
     * @var
     * Property field label displayed on top
     */
    label: string;
    /**
     * @var
     * TermSet Picker Panel title
     */
    panelTitle: string;
    /**
     * @var
     * Defines if the user can select only one or many term sets. Default value is false.
     *
     */
    allowMultipleSelections?: boolean;
    /**
     * @var
     * Defines the selected by default term sets.
     */
    initialValues?: ISPTermSets;
    /**
     * @var
     * Indicator to define if the system Groups are exclude. Default is false.
     */
    excludeSystemGroup?: boolean;
    /**
     * @var
     * Indicates if the offline term stores must be exclude. Default is false.
     */
    excludeOfflineTermStores?: boolean;
    /**
     * @var
     * Restrict term sets that are available for tagging. Default is false.
     */
    displayOnlyTermSetsAvailableForTagging?: boolean;
    /**
     * @var
     * WebPart's context
     */
    context: IWebPartContext;
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
    onGetErrorMessage?: (value: ISPTermSets) => string | Promise<string>;
    /**
     * Custom Field will start to validate after users stop typing for `deferredValidationTime` milliseconds.
     * Default value is 200.
     */
    deferredValidationTime?: number;
}
/**
 * @interface
 * Private properties of the PropertyFieldTermSetPicker custom field.
 * We separate public & private properties to include onRender & onDispose method waited
 * by the PropertyFieldCustom, witout asking to the developer to add it when he's using
 * the PropertyFieldTermSetPicker.
 *
 */
export interface IPropertyFieldTermSetPickerPropsInternal extends IPropertyFieldTermSetPickerProps {
    label: string;
    targetProperty: string;
    panelTitle: string;
    allowMultipleSelections?: boolean;
    initialValues?: ISPTermSets;
    excludeSystemGroup?: boolean;
    excludeOfflineTermStores?: boolean;
    displayOnlyTermSetsAvailableForTagging?: boolean;
    context: IWebPartContext;
    onRender(elem: HTMLElement): void;
    onDispose(elem: HTMLElement): void;
    onPropertyChange(propertyPath: string, oldValue: any, newValue: any): void;
    render(): void;
    disableReactivePropertyChanges?: boolean;
    properties: any;
    key: string;
    disabled?: boolean;
    onGetErrorMessage?: (value: ISPTermSets) => string | Promise<string>;
    deferredValidationTime?: number;
}
/**
 * @function
 * Helper method to create a SPList Picker on the PropertyPane.
 * @param targetProperty - Target property the SharePoint list picker is associated to.
 * @param properties - Strongly typed SPList Picker properties.
 */
export declare function PropertyFieldTermSetPicker(targetProperty: string, properties: IPropertyFieldTermSetPickerProps): IPropertyPaneField<IPropertyFieldTermSetPickerPropsInternal>;
