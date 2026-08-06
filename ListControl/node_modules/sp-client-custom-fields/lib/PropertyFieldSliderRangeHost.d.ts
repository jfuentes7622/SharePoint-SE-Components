/**
 * @file PropertyFieldSliderRangeHost.tsx
 * Renders the controls for PropertyFieldSliderRange component
 *
 * @copyright 2016 Olivier Carpentier
 * Released under MIT licence
 */
import * as React from 'react';
import { IPropertyFieldSliderRangePropsInternal } from './PropertyFieldSliderRange';
/**
 * @interface
 * PropertyFieldSliderRangeHost properties interface
 *
 */
export interface IPropertyFieldSliderRangeHostProps extends IPropertyFieldSliderRangePropsInternal {
}
export interface IPropertyFieldSliderRangeHostState {
}
/**
 * @class
 * Renders the controls for PropertyFieldSliderRange component
 */
export default class PropertyFieldSliderRangeHost extends React.Component<IPropertyFieldSliderRangeHostProps, IPropertyFieldSliderRangeHostState> {
    /**
     * @function
     * Constructor
     */
    constructor(props: IPropertyFieldSliderRangeHostProps);
    /**
     * @function
     * Renders the controls
     */
    render(): JSX.Element;
}
