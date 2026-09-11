import * as React from 'react';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react';
import { SPHttpClient } from '@microsoft/sp-http';
import { SharePointDynamicFormContainer, SharePointDynamicFormContainerProps } from './SharePointDynamicForm';
import * as strings from 'SharePointDynamicFormWebPartStrings';

export interface IRepeatedReportFormsProps {
  reportProps: SharePointDynamicFormContainerProps;
}

interface IRepeatedReportFormsState {
  itemIds: number[];
  loading: boolean;
  error: string;
}

interface IRepeatFilterField {
  internalName: string;
  typeAsString: string;
}

interface IRepeatFilterCondition {
  field: string;
  type?: string;
  value: any;
  valueType?: string;
  operator?: string;
  logical?: string;
}

export class RepeatedReportForms extends React.Component<IRepeatedReportFormsProps, IRepeatedReportFormsState> {
  private _loadRequestId: number = 0;

  constructor(props: IRepeatedReportFormsProps) {
    super(props);
    this.state = {
      itemIds: [],
      loading: false,
      error: '',
    };
  }

  public componentDidMount(): void {
    this.loadItemIds();
  }

  public componentDidUpdate(prevProps: IRepeatedReportFormsProps): void {
    if (prevProps.reportProps.listName !== this.props.reportProps.listName
      || prevProps.reportProps.itemIdQueryParam !== this.props.reportProps.itemIdQueryParam
      || prevProps.reportProps.linkedFieldTarget !== this.props.reportProps.linkedFieldTarget
      || prevProps.reportProps.useDynamicValueAsFilter !== this.props.reportProps.useDynamicValueAsFilter
      || prevProps.reportProps.filterJson !== this.props.reportProps.filterJson) {
      this.loadItemIds();
    }
  }

  public componentWillUnmount(): void {
    this._loadRequestId += 1;
  }

  private async getWithAcceptFallback(url: string): Promise<any> {
    var response = await this.props.reportProps.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    var acceptValues = [
      'application/json;odata=verbose',
      'application/json;odata=minimalmetadata',
      'application/json;odata=nometadata',
    ];
    for (var i = 0; !response.ok && i < acceptValues.length; i += 1) {
      response = await this.props.reportProps.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: acceptValues[i] }
      });
    }
    return response;
  }

  private getUrlQueryValue(parameterName: string): string {
    if (typeof window === 'undefined') { return ''; }
    var query = String(window.location.search || '').replace(/^\?/, '').split('&');
    var requestedName = String(parameterName || '').toLowerCase();
    for (var i = 0; i < query.length; i += 1) {
      var parts = query[i].split('=');
      var name = decodeURIComponent(parts[0] || '').toLowerCase();
      if (name === requestedName) {
        return decodeURIComponent((parts.slice(1).join('=') || '').replace(/\+/g, ' '));
      }
    }
    return '';
  }

  private async loadRepeatFilterField(listName: string, fieldName: string): Promise<IRepeatFilterField | null> {
    var webUrl = String(this.props.reportProps.context.pageContext.web.absoluteUrl || '').replace(/\/$/, '');
    var endpoint = webUrl + "/_api/web/lists/getByTitle('" + listName.replace(/'/g, "''")
      + "')/fields/getByInternalNameOrTitle('" + fieldName.replace(/'/g, "''") + "')?$select=InternalName,TypeAsString";
    var response = await this.getWithAcceptFallback(endpoint);
    if (!response.ok) { return null; }
    var data = await response.json();
    var field = data && data.d ? data.d : data;
    return field && field.InternalName ? {
      internalName: String(field.InternalName),
      typeAsString: String(field.TypeAsString || '')
    } : null;
  }

  private buildRepeatFilter(field: IRepeatFilterField, rawValue: string): string {
    var fieldType = field.typeAsString.toLowerCase();
    var fieldName = field.internalName;
    if (fieldType === 'lookup' || fieldType === 'lookupmulti' || fieldType === 'user' || fieldType === 'usermulti') {
      var lookupId = parseInt(rawValue, 10);
      return !isNaN(lookupId) && lookupId > 0 ? fieldName + 'Id eq ' + String(lookupId) : '';
    }
    if (fieldType === 'number' || fieldType === 'currency' || fieldType === 'integer' || fieldType === 'counter') {
      var numericValue = Number(rawValue);
      return !isNaN(numericValue) ? fieldName + ' eq ' + String(numericValue) : '';
    }
    if (fieldType === 'boolean') {
      var booleanValue = String(rawValue).trim().toLowerCase();
      if (booleanValue === 'true' || booleanValue === '1' || booleanValue === 'yes') { return fieldName + ' eq true'; }
      if (booleanValue === 'false' || booleanValue === '0' || booleanValue === 'no') { return fieldName + ' eq false'; }
      return '';
    }
    return fieldName + " eq '" + String(rawValue).replace(/'/g, "''") + "'";
  }

  private resolveFilterExpressionValue(value: any): any {
    if (!value || typeof value !== 'string') { return value; }
    var normalized = value.trim().toLowerCase();
    var pageContext = this.props.reportProps.context && this.props.reportProps.context.pageContext;
    var user = pageContext && pageContext.user;
    var legacyContext = pageContext && pageContext.legacyPageContext;
    if (normalized === 'today') {
      var today = new Date();
      return today.getFullYear() + '-' + ('0' + String(today.getMonth() + 1)).slice(-2) + '-' + ('0' + String(today.getDate())).slice(-2);
    }
    if (normalized === 'now') { return new Date().toISOString(); }
    var dateMatch = /^date\(\s*([+-]?\d+)\s*\)$/.exec(normalized);
    if (dateMatch) {
      var dateValue = new Date();
      dateValue.setDate(dateValue.getDate() + parseInt(dateMatch[1], 10));
      return dateValue.getFullYear() + '-' + ('0' + String(dateValue.getMonth() + 1)).slice(-2) + '-' + ('0' + String(dateValue.getDate())).slice(-2);
    }
    if (normalized === 'me.email') { return String(user && user.email || ''); }
    if (normalized === 'me.login') { return String(user && user.loginName || ''); }
    if (normalized === 'me.id') { return legacyContext && legacyContext.userId || ''; }
    if (normalized === 'me') { return String(user && (user.displayName || user.loginName) || ''); }
    return value;
  }

  private buildJsonFilterClause(condition: IRepeatFilterCondition, value: any): string {
    var fieldName = String(condition.field || '');
    var fieldType = String(condition.type || 'text').toLowerCase();
    var operator = String(condition.operator || 'eq').toLowerCase();
    if (!fieldName || value === undefined || value === null || value === '') { return ''; }
    if (Array.isArray(value)) {
      var arrayOperator = operator === 'ne' ? 'eq' : (operator === 'notcontains' ? 'contains' : operator);
      var arrayClauses = value.map((entry: any) => this.buildJsonFilterClause(Object.assign({}, condition, { operator: arrayOperator }), entry))
        .filter(function(clause: string) { return !!clause; });
      if (operator === 'ne' || operator === 'notcontains') {
        return arrayClauses.map(function(clause: string) { return 'not (' + clause + ')'; }).join(' and ');
      }
      return arrayClauses.map(function(clause: string) { return '(' + clause + ')'; }).join(' or ');
    }
    if (fieldType === 'lookup' || fieldType === 'person') {
      var lookupId = parseInt(String(value), 10);
      if (isNaN(lookupId)) { return ''; }
      return fieldName + 'Id ' + (operator === 'ne' ? 'ne' : 'eq') + ' ' + String(lookupId);
    }
    if (fieldType === 'number') {
      var numericValue = Number(value);
      if (isNaN(numericValue)) { return ''; }
      return fieldName + ' ' + (/^(eq|ne|gt|ge|lt|le)$/.test(operator) ? operator : 'eq') + ' ' + String(numericValue);
    }
    if (fieldType === 'boolean') {
      var booleanValue = value === true || String(value).toLowerCase() === 'true' || String(value) === '1';
      return fieldName + ' ' + (operator === 'ne' ? 'ne' : 'eq') + ' ' + (booleanValue ? 'true' : 'false');
    }
    if (fieldType === 'datetime') {
      var parsedDate = new Date(String(value));
      if (isNaN(parsedDate.getTime())) { return ''; }
      return fieldName + ' ' + (/^(eq|ne|gt|ge|lt|le)$/.test(operator) ? operator : 'eq') + " datetime'" + parsedDate.toISOString() + "'";
    }
    var escapedValue = String(value).replace(/'/g, "''");
    if (operator === 'contains' || operator === 'notcontains' || operator === 'startswith' || operator === 'endswith') {
      var textOperator = operator === 'notcontains' ? 'contains' : operator;
      var textClause = textOperator + '(' + fieldName + ", '" + escapedValue + "')";
      return operator === 'notcontains' ? 'not (' + textClause + ')' : textClause;
    }
    return fieldName + ' ' + (/^(eq|ne|gt|ge|lt|le)$/.test(operator) ? operator : 'eq') + " '" + escapedValue + "'";
  }

  private buildJsonFilterExpression(filterJson: string, dynamicValue: any): string {
    var conditions = JSON.parse(filterJson) as IRepeatFilterCondition[];
    if (!Array.isArray(conditions)) { throw new Error('Filter JSON must be an array.'); }
    var expression = '';
    for (var i = 0; i < conditions.length; i += 1) {
      var condition = conditions[i];
      if (!condition || !condition.field) { continue; }
      var value = condition.value === 'dynamic'
        ? dynamicValue
        : (String(condition.valueType || '').toLowerCase() === 'expression'
          ? this.resolveFilterExpressionValue(condition.value)
          : condition.value);
      var clause = this.buildJsonFilterClause(condition, value);
      if (!clause) { continue; }
      expression = expression
        ? '(' + expression + ') ' + (String(condition.logical || '').toLowerCase() === 'or' ? 'or' : 'and') + ' (' + clause + ')'
        : clause;
    }
    return expression;
  }

  private getResponseItems(data: any): any[] {
    if (data && Array.isArray(data.value)) { return data.value; }
    if (data && data.d && Array.isArray(data.d.results)) { return data.d.results; }
    if (data && Array.isArray(data.results)) { return data.results; }
    var singleton = data && data.d ? data.d : data;
    return singleton && (singleton.Id !== undefined || singleton.ID !== undefined) ? [singleton] : [];
  }

  private async loadItemIds(): Promise<void> {
    var requestId = this._loadRequestId + 1;
    this._loadRequestId = requestId;
    var listName = String(this.props.reportProps.listName || '').trim();
    if (!listName) {
      this.setState({ itemIds: [], loading: false, error: '' });
      return;
    }

    this.setState({ itemIds: [], loading: true, error: '' });
    try {
      var webUrl = String(this.props.reportProps.context.pageContext.web.absoluteUrl || '').replace(/\/$/, '');
      var escapedListName = listName.replace(/'/g, "''");
      var nextUrl = webUrl + "/_api/web/lists/getByTitle('" + escapedListName + "')/items?$select=Id&$orderby=Id asc&$top=5000";
      var filterExpressions: string[] = [];
      var dynamicFilterValue: any = this.props.reportProps.linkedFieldValue;
      if (this.props.reportProps.useDynamicValueAsFilter && this.props.reportProps.linkedFieldTarget) {
        var queryParameter = String(this.props.reportProps.itemIdQueryParam || 'itemid').trim() || 'itemid';
        var queryValue = this.getUrlQueryValue(queryParameter);
        if (queryValue) {
          dynamicFilterValue = queryValue;
          var filterField = await this.loadRepeatFilterField(listName, String(this.props.reportProps.linkedFieldTarget));
          var filterExpression = filterField ? this.buildRepeatFilter(filterField, queryValue) : '';
          if (!filterExpression) {
            throw new Error(strings.RuntimeRepeatLoadFailed + ': unable to apply URL filter.');
          }
          filterExpressions.push(filterExpression);
        }
      }
      if (String(this.props.reportProps.filterJson || '').trim()) {
        var jsonFilterExpression = this.buildJsonFilterExpression(String(this.props.reportProps.filterJson), dynamicFilterValue);
        if (!jsonFilterExpression) {
          throw new Error(strings.RuntimeRepeatLoadFailed + ': the configured JSON filter produced no valid conditions.');
        }
        filterExpressions.push(jsonFilterExpression);
      }
      if (filterExpressions.length > 0) {
        nextUrl += '&$filter=' + encodeURIComponent(filterExpressions.map(function(expression: string) {
          return '(' + expression + ')';
        }).join(' and '));
      }
      var itemIds: number[] = [];
      var seenItemIds: { [itemId: string]: boolean } = {};
      var pageCount = 0;

      while (nextUrl && pageCount < 1000) {
        var response = await this.getWithAcceptFallback(nextUrl);
        if (!response.ok) {
          throw new Error(strings.RuntimeRepeatLoadFailed + ' Status: ' + String(response.status));
        }
        var data = await response.json();
        var items = this.getResponseItems(data);
        for (var i = 0; i < items.length; i += 1) {
          var itemId = parseInt(String(items[i].Id !== undefined ? items[i].Id : items[i].ID), 10);
          if (!isNaN(itemId) && itemId > 0 && !seenItemIds[String(itemId)]) {
            seenItemIds[String(itemId)] = true;
            itemIds.push(itemId);
          }
        }
        nextUrl = String((data && (data['@odata.nextLink'] || data['odata.nextLink'])) || (data && data.d && data.d.__next) || '');
        pageCount += 1;
      }

      if (requestId === this._loadRequestId) {
        this.setState({ itemIds: itemIds, loading: false, error: '' });
      }
    } catch (error) {
      if (requestId === this._loadRequestId) {
        var message = error && (error as any).message ? String((error as any).message) : strings.RuntimeRepeatLoadFailed;
        this.setState({ itemIds: [], loading: false, error: message });
      }
    }
  }

  public render(): JSX.Element {
    if (this.state.loading) {
      return <div>{strings.CommonLoading}</div>;
    }
    if (this.state.error) {
      return <MessageBar messageBarType={MessageBarType.error}>{this.state.error}</MessageBar>;
    }
    if (this.state.itemIds.length === 0) {
      return <MessageBar messageBarType={MessageBarType.info}>{strings.RuntimeRepeatNoItems}</MessageBar>;
    }

    return (
      <div data-reportforms-repeat="true">
        {this.state.itemIds.map((itemId) => {
          var childProps = Object.assign({}, this.props.reportProps, {
            isInDesignerMode: false,
            isDesignerAvailable: false,
            onToggleDesignerMode: function(): void { return; },
            dynamicItemId: 0,
            linkedFieldTarget: '',
            linkedFieldValue: undefined,
            useDynamicValueAsFilter: false,
            filterJson: '',
            useItemId: true,
            itemId: itemId,
            itemIdQueryParam: this.props.reportProps.itemIdQueryParam,
            dynamicItemReference: '',
            dynamicModeReference: '',
            dynamicPreferredSourceInstanceId: '',
            isPageEditMode: false,
            onRuntimeStateChange: undefined,
          }) as SharePointDynamicFormContainerProps;
          return (
            <div key={itemId} data-reportforms-repeat-item-id={String(itemId)} style={{ marginBottom: '24px' }}>
              {React.createElement(SharePointDynamicFormContainer, childProps)}
            </div>
          );
        })}
      </div>
    );
  }
}