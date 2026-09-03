/// <reference types="react" />
import * as React from 'react';
import './ListDesigner.css';
export interface IListDesignerColumn {
    fieldName: string;
    displayName: string;
    width?: string;
}
export interface IListGroupingConfig {
    field1: string;
    field2: string;
    collapsedByDefault: boolean;
    showCount: boolean;
}
export interface IListDesignerProps {
    context: any;
    listName: string;
    viewId: string;
    viewColumns: IListDesignerColumn[];
    groupingJson: string;
    onSave: (columns: IListDesignerColumn[], groupingJson: string) => void;
    onCancel: () => void;
}
export interface ISharePointField {
    internalName: string;
    title: string;
}
export interface IListDesignerField {
    id: string;
    fieldName: string;
    displayName: string;
    width: string;
}
export interface IListDesignerState {
    fields: IListDesignerField[];
    sharePointFields: ISharePointField[];
    selectedFieldId: string;
    loading: boolean;
    error: string;
    grouping: IListGroupingConfig;
    viewLoading: boolean;
    viewLoadMessage: string;
}
export declare class ListDesigner extends React.Component<IListDesignerProps, IListDesignerState> {
    constructor(props: IListDesignerProps);
    componentDidMount(): void;
    private getWebUrl();
    private getFieldsResponse(url);
    private loadFields();
    private loadFromView();
    private getAvailableFields();
    private addField(source);
    private getSelectedField();
    private updateSelectedField(mutator);
    private moveField(fieldId, direction);
    private removeField(fieldId);
    private updateGrouping(mutator);
    private getFieldLabel(fieldName);
    private save();
    private renderPalette();
    private renderGroupingSettings();
    private renderColumnsList();
    private renderCanvas();
    private renderFieldEditor();
    render(): JSX.Element;
}
