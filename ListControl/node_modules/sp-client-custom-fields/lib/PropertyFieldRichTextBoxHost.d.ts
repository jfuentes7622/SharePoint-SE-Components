/**
 * @file PropertyFieldRichTextBoxHost.tsx
 * Renders the controls for PropertyFieldRichTextBox component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldRichTextBoxPropsInternal } from './PropertyFieldRichTextBox';
/**
 * @interface
 * PropertyFieldRichTextBoxHost properties interface
 *
 */
export interface IPropertyFieldRichTextBoxHostProps extends IPropertyFieldRichTextBoxPropsInternal {
    keyCopy: string;
}
export interface IPropertyFieldRichTextBoxHostState {
}
/**
 * @class
 * Renders the controls for PropertyFieldRichTextBox component
 */
export default class PropertyFieldRichTextBoxHost extends React.Component<IPropertyFieldRichTextBoxHostProps, IPropertyFieldRichTextBoxHostState> {
    /**
     * @function
     * Constructor
     */
    constructor(props: IPropertyFieldRichTextBoxHostProps);
    /**
     * @function
     * Renders the controls
     */
    render(): JSX.Element;
}
