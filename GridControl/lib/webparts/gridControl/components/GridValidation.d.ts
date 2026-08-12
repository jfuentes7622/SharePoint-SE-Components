export interface IGridValidationField {
    id: string;
    fieldName: string;
    label?: string;
}
export interface IGridAdvancedValidationRule {
    id?: string;
    expression: string;
    message: string;
    targetField?: string;
}
export declare function evaluateGridValidationExpression(expression: string, fields: IGridValidationField[], values: {
    [fieldId: string]: any;
}): any;
