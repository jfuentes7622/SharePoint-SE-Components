/// <reference types="react" />
import * as React from 'react';
import './GridControl.css';
export interface IGridControlViewOption {
    key: string;
    text: string;
    isDefault?: boolean;
}
export interface IGridControlColumnConfiguration {
    fieldName: string;
    displayName: string;
    width?: string;
}
export interface IGridControlProps {
    context: any;
    listName: string;
    defaultViewId: string;
    views: IGridControlViewOption[];
    viewColumns: IGridControlColumnConfiguration[];
    gridSchemaJson: string;
    pageSize: number;
    isEditMode: boolean;
    showViewSelector: boolean;
    showRefresh: boolean;
    showAdd: boolean;
    showDelete: boolean;
    showLinkToItem: boolean;
    linkTargetPageUrl: string;
    linkTargetIdParam: string;
    includeReturnUrlParam: boolean;
    enableDiagnostics: boolean;
    bodyTextColor: string;
    bodyFontFamily: string;
    bodyFontSize: string;
    bodyFontStyle: string;
    bodyFontBold: boolean;
    bodyTextAlign: string;
    selectedTextColor: string;
    selectedBackgroundColor: string;
    selectedFontStyle: string;
    selectedFontBold: boolean;
    headerTextColor: string;
    headerBackgroundColor: string;
    headerFontFamily: string;
    headerFontSize: string;
    headerFontStyle: string;
    headerFontBold: boolean;
    headerTextAlign: string;
    tableBackgroundColor: string;
    tableBorderColor: string;
    tableBorderWidth: string;
    tableCornerStyle: string;
    tableCornerRadius: number;
    tableRowLineWidth: number;
    cornerStyle: string;
    cornerRadius: number;
    alternateRowShading: boolean;
    alternateRowShadingColor: string;
    buttonTextColor: string;
    buttonBackgroundColor: string;
    buttonFontFamily: string;
    buttonFontSize: string;
    buttonFontStyle: string;
    buttonFontBold: boolean;
    buttonCornerStyle: string;
    buttonCornerRadius: number;
    webpartBackgroundColor: string;
    webpartBorderColor: string;
    webpartBorderWidth: number;
    filterJson?: string;
    conditionalStyleJson?: string;
    onSelectionChange: (itemId: number, mode: string) => void;
}
export interface IListFieldDefinition {
    Name: string;
    RealFieldName?: string;
    DisplayName?: string;
    Hidden?: string | boolean;
    ConfiguredWidth?: string;
}
export interface IGridFieldMetadata {
    internalName: string;
    title: string;
    typeAsString: string;
    required: boolean;
    readOnly: boolean;
    hidden: boolean;
    description: string;
    choices: string[];
    displayFormat: number;
}
export interface IGridControlState {
    selectedViewId: string;
    fields: IListFieldDefinition[];
    fieldMetadataByName: {
        [fieldName: string]: IGridFieldMetadata;
    };
    rows: any[];
    loading: boolean;
    error: string | null;
    selectedItemId: number;
    selectedItemIds: number[];
    selectedMode: string;
    deleting: boolean;
    deleteMessage: string;
    sortFieldName: string;
    sortDirection: 'asc' | 'desc' | '';
    activeFilterFieldName: string;
    draftFilterOperator: FilterOperator;
    draftFilterValue: string;
    columnFilters: {
        [fieldName: string]: IColumnFilter;
    };
    currentPage: number;
    editingItemId: number;
    editingValues: {
        [fieldName: string]: any;
    };
    editingErrors: {
        [fieldName: string]: string;
    };
    saving: boolean;
}
export declare type FilterOperator = 'eq' | 'ne' | 'contains' | 'notcontains' | 'startswith' | 'endswith' | 'gt' | 'ge' | 'lt' | 'le';
export interface IColumnFilter {
    operator: FilterOperator;
    value: string;
    compareDateOnly?: boolean;
}
export declare class GridControl extends React.Component<IGridControlProps, IGridControlState> {
    private _refreshEventHandler;
    constructor(props: IGridControlProps);
    componentDidMount(): void;
    componentWillUnmount(): void;
    componentDidUpdate(prevProps: IGridControlProps, prevState: IGridControlState): void;
    private getInitialViewId(views);
    private handleExternalRefresh(event);
    private getWebUrl();
    private getJsonWithFallback(url);
    private postJsonWithFallback(url, body);
    private buildViewIdCandidates(selectedViewId);
    private loadSelectedViewFieldNames(selectedViewId);
    private getDisplayFields();
    private getGridSchemaFields();
    private getGridSchema();
    private getAdvancedValidationRules();
    private getGridSchemaField(fieldName);
    private getConfiguredColumnStyle(field);
    private getFieldsForConsumption(rawFields, viewFieldNames);
    private loadListFieldTypeMap(viewFieldNames);
    private loadListFieldTitleMap();
    private loadGridFieldMetadata();
    private applyFieldDisplayNames(fields, titleMap);
    private getRowFieldValue(row, field);
    private stringifyCellValue(value);
    private isMeaningfulCellValue(value);
    private filterRenderableRows(rows, visibleFields);
    private loadRowsFromItemsEndpoint(viewFieldNames);
    private loadRows();
    private getRowItemId(row);
    private selectRow(row);
    private isItemChecked(itemId);
    private toggleItemChecked(itemId);
    private toggleVisibleItemsChecked(rows);
    private getGridFieldMetadata(field);
    private isEditableGridField(field);
    private normalizeEditingValue(value, metadata);
    private beginRowEdit(row);
    private beginNewRow();
    private cancelRowEdit();
    private updateEditingValue(fieldName, value);
    private validateEditingValues();
    private buildEditingPayload();
    private saveEditingRow();
    private renderEditingControl(field);
    private renderEditingRow(displayFields);
    private deleteSelected();
    private logDiagnostic(message);
    private getCellMarkup(row, field);
    private isTitleField(field);
    private getCellPlainText(row, field);
    private getItemLinkUrl(row);
    private getFieldKey(field);
    private getFilterOperatorOptions();
    private toggleSort(field);
    private openFilter(field);
    private closeFilter();
    private applyActiveFilter();
    private clearActiveFilter();
    private compareComparableValues(leftValue, rightValue);
    private rowMatchesFilter(row, field, filter);
    private parsePresetFilterConditions();
    private resolveFieldByReference(fieldsByKey, fieldRef);
    private rowMatchesPresetConditions(row, fieldsByKey, conditions);
    private resolvePresetFilterValue(condition);
    private getProcessedRows();
    private parseConditionalStyleRules();
    private toReactCssStyle(styleDefinition);
    private evaluateConditionalStyleRule(row, rule, fieldsByKey);
    private getConditionalStyleForRow(row, fieldsByKey, rules);
    private getMatchedConditionalRuleNames(row, fieldsByKey, rules);
    render(): JSX.Element;
}
