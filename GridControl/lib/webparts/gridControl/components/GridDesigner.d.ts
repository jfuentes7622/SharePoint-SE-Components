/// <reference types="react" />
import * as React from 'react';
import { IGridAdvancedValidationRule } from './GridValidation';
import './GridDesigner.css';
export interface IGridDesignerProps {
    context: any;
    listName: string;
    schemaJson: string;
    onSave: (schemaJson: string) => void;
    onCancel: () => void;
}
export interface ISharePointField {
    internalName: string;
    title: string;
    description: string;
    type: string;
    required: boolean;
    readOnly: boolean;
    choices: string[];
    maxLength?: number;
    displayFormat?: string;
}
export interface IGridDesignerField {
    id: string;
    fieldName: string;
    sharePointType?: string;
    sharePointRequired?: boolean;
    sharePointDescription?: string;
    type: string;
    label: string;
    description?: string;
    visible?: boolean;
    required?: boolean;
    requiredMessage?: string;
    readOnly?: boolean;
    disabled?: boolean;
    defaultValue?: any;
    config?: any;
    validation?: any[];
    gridWidth?: string;
}
export interface IGridDesignerState {
    fields: IGridDesignerField[];
    sharePointFields: ISharePointField[];
    selectedFieldId: string;
    loading: boolean;
    error: string;
    advancedValidationEnabled: boolean;
    advancedValidationRules: IGridAdvancedValidationRule[];
    validationExpression: string;
    validationMessage: string;
    validationTargetField: string;
    selectedValidationRuleIndex: string;
    validationDesignerMessage: string;
}
export declare class GridDesigner extends React.Component<IGridDesignerProps, IGridDesignerState> {
    constructor(props: IGridDesignerProps);
    componentDidMount(): void;
    private getWebUrl();
    private getFieldsResponse(url);
    private loadFields();
    private getAvailableFields();
    private addField(source);
    private getSelectedField();
    private updateSelectedField(mutator);
    private moveField(fieldId, direction);
    private removeField(fieldId);
    private updateConfig(name, value);
    private changeSelectedControlType(controlType);
    private updateValidation(type, value, message);
    private updateValidationMessages(types, message);
    private getValidation(field, type);
    private getPatternError(pattern);
    private getValidationFields();
    private validateAdvancedExpression(expression);
    private appendValidationExpression(fragment);
    private loadValidationRule(indexValue);
    private saveValidationRule(updateExisting);
    private removeValidationRule();
    private save();
    private renderPalette();
    private renderCanvas();
    private renderFieldEditor();
    render(): JSX.Element;
}
