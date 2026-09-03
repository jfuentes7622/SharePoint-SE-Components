import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
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

var SYSTEM_FIELDS: { [name: string]: boolean } = {
  ID: true,
  GUID: true,
  ContentType: true,
  AppAuthor: true,
  AppEditor: true,
  Edit: true,
  ItemChildCount: true,
  FolderChildCount: true,
  ComplianceAssetId: true
};

function createId(): string {
  return 'listfield_' + new Date().getTime().toString(36) + '_' + Math.floor(Math.random() * 100000).toString(36);
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function toArray(value: any): any[] {
  if (Array.isArray(value)) {
    return value;
  }
  if (value && Array.isArray(value.results)) {
    return value.results;
  }
  return [];
}

function trimGuidBraces(value: string): string {
  return String(value || '').replace(/^[{]/, '').replace(/[}]$/, '');
}

function buildViewIdCandidates(selectedViewId: string): string[] {
  var normalized = trimGuidBraces(selectedViewId);
  var candidates = [String(selectedViewId || ''), normalized, '{' + normalized + '}'];
  var unique: string[] = [];
  for (var i = 0; i < candidates.length; i += 1) {
    var value = String(candidates[i] || '');
    if (value && unique.indexOf(value) < 0) {
      unique.push(value);
    }
  }
  return unique;
}

function parseGroupByFieldNames(viewQuery: string): { fieldNames: string[]; collapsed: boolean } {
  var queryText = String(viewQuery || '').trim();
  if (!queryText) { return { fieldNames: [], collapsed: false }; }
  try {
    var xmlDocument = new DOMParser().parseFromString(
      /^<Query(?:\s|>)/i.test(queryText) ? queryText : '<Query>' + queryText + '</Query>',
      'text/xml'
    );
    if (xmlDocument.getElementsByTagName('parsererror').length > 0) { return { fieldNames: [], collapsed: false }; }
    var groupByElements = xmlDocument.getElementsByTagName('GroupBy');
    if (groupByElements.length === 0) { return { fieldNames: [], collapsed: false }; }
    var groupByElement = groupByElements[0];
    var fieldRefs = groupByElement.getElementsByTagName('FieldRef');
    var fieldNames: string[] = [];
    for (var i = 0; i < fieldRefs.length && fieldNames.length < 2; i += 1) {
      var name = fieldRefs[i].getAttribute('Name');
      if (name) { fieldNames.push(name); }
    }
    var collapsed = String(groupByElement.getAttribute('Collapse') || '').toUpperCase() === 'TRUE';
    return { fieldNames: fieldNames, collapsed: collapsed };
  } catch (_error) {
    return { fieldNames: [], collapsed: false };
  }
}

function copyFields(fields: IListDesignerField[]): IListDesignerField[] {
  return JSON.parse(JSON.stringify(fields || []));
}

function parseGrouping(groupingJson: string): IListGroupingConfig {
  var defaultGrouping: IListGroupingConfig = { field1: '', field2: '', collapsedByDefault: false, showCount: true };
  if (!String(groupingJson || '').trim()) { return defaultGrouping; }
  try {
    var grouping = JSON.parse(groupingJson);
    if (!grouping) { return defaultGrouping; }
    return {
      field1: String(grouping.field1 || ''),
      field2: String(grouping.field2 || ''),
      collapsedByDefault: grouping.collapsedByDefault === true,
      showCount: grouping.showCount !== false
    };
  } catch (_error) {
    return defaultGrouping;
  }
}

export class ListDesigner extends React.Component<IListDesignerProps, IListDesignerState> {
  public constructor(props: IListDesignerProps) {
    super(props);
    var initialFields: IListDesignerField[] = (props.viewColumns || []).filter(function(column) {
      return !!(column && column.fieldName);
    }).map(function(column) {
      return {
        id: createId(),
        fieldName: column.fieldName,
        displayName: column.displayName || column.fieldName,
        width: column.width || ''
      };
    });
    this.state = {
      fields: initialFields,
      sharePointFields: [],
      selectedFieldId: initialFields.length > 0 ? initialFields[0].id : '',
      loading: false,
      error: '',
      grouping: parseGrouping(props.groupingJson),
      viewLoading: false,
      viewLoadMessage: ''
    };
  }

  public componentDidMount(): void {
    this.loadFields();
  }

  private getWebUrl(): string {
    return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
  }

  private async getFieldsResponse(url: string): Promise<any> {
    var acceptHeaders = [
      '',
      'application/json;odata=verbose',
      'application/json;odata=minimalmetadata',
      'application/json;odata=nometadata'
    ];
    var response: any;
    for (var i = 0; i < acceptHeaders.length; i += 1) {
      response = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        acceptHeaders[i] ? { headers: { Accept: acceptHeaders[i] } } : undefined
      );
      if (response.ok) {
        return response;
      }
    }
    return response;
  }

  private async loadFields(): Promise<void> {
    if (!this.props.listName) {
      this.setState({ error: 'Select a SharePoint list before opening List Designer.' });
      return;
    }
    this.setState({ loading: true, error: '' });
    try {
      var url = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,Title,Hidden,FromBaseType";
      var response = await this.getFieldsResponse(url);
      if (!response.ok) {
        throw new Error('SharePoint fields could not be loaded. HTTP ' + String(response.status) + ' ' + String(response.statusText || ''));
      }
      var data = await response.json();
      var sourceFields = toArray(data && data.value).length > 0 ? toArray(data.value) : toArray(data && data.d && data.d.results);
      var fields = sourceFields.filter(function(field: any) {
        var isAttachment = field.InternalName === 'Attachments';
        var isCommonSystemField = field.InternalName === 'Author' || field.InternalName === 'Editor'
          || field.InternalName === 'Created' || field.InternalName === 'Modified';
        return !field.Hidden && !SYSTEM_FIELDS[field.InternalName]
          && (!field.FromBaseType || field.InternalName === 'Title' || isAttachment || isCommonSystemField);
      }).map(function(field: any) {
        return {
          internalName: String(field.InternalName || ''),
          title: String(field.Title || field.InternalName || '')
        } as ISharePointField;
      });
      this.setState({ sharePointFields: fields, loading: false });
    } catch (error) {
      this.setState({ loading: false, error: error && error.message ? error.message : 'SharePoint fields could not be loaded.' });
    }
  }

  private async loadFromView(): Promise<void> {
    if (!this.props.listName || !this.props.viewId) {
      this.setState({ viewLoadMessage: 'Select a SharePoint view in the web part properties first.' });
      return;
    }
    if (this.state.fields.length > 0 && typeof window !== 'undefined'
      && !window.confirm('Replace the current columns and grouping with the selected view\'s fields and grouping?')) {
      return;
    }

    this.setState({ viewLoading: true, viewLoadMessage: '' });
    try {
      var webUrl = this.getWebUrl();
      var listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
      var viewIds = buildViewIdCandidates(this.props.viewId);
      var viewQuery = '';
      var viewFieldNames: string[] = [];

      for (var i = 0; i < viewIds.length && !viewQuery; i += 1) {
        var encoded = encodeURIComponent(viewIds[i]);
        var normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
        var queryUrls = [
          webUrl + listPath + "/views/getById('" + encoded + "')?$select=ViewQuery",
          webUrl + listPath + "/views(guid'" + normalized + "')?$select=ViewQuery"
        ];
        for (var q = 0; q < queryUrls.length && !viewQuery; q += 1) {
          var queryResponse = await this.getFieldsResponse(queryUrls[q]);
          if (!queryResponse.ok) { continue; }
          var queryData = await queryResponse.json();
          var queryContainer = queryData && queryData.d ? queryData.d : queryData;
          if (queryContainer && queryContainer.ViewQuery !== undefined && queryContainer.ViewQuery !== null) {
            viewQuery = String(queryContainer.ViewQuery);
          }
        }

        var fieldsUrls = [
          webUrl + listPath + "/views/getById('" + encoded + "')/ViewFields",
          webUrl + listPath + "/views(guid'" + normalized + "')/ViewFields"
        ];
        for (var f = 0; f < fieldsUrls.length && viewFieldNames.length === 0; f += 1) {
          var fieldsResponse = await this.getFieldsResponse(fieldsUrls[f]);
          if (!fieldsResponse.ok) { continue; }
          var fieldsData = await fieldsResponse.json();
          var names = toArray(fieldsData.value);
          if (names.length === 0) { names = toArray(fieldsData.Items); }
          if (names.length === 0) { names = toArray(fieldsData && fieldsData.d && fieldsData.d.Items); }
          if (names.length > 0) {
            viewFieldNames = names.map(function(name: any) { return String(name || ''); }).filter(function(name: string) { return !!name; });
          }
        }
      }

      var groupByResult = parseGroupByFieldNames(viewQuery);

      if (viewFieldNames.length > 0) {
        var nextFields: IListDesignerField[] = [];
        var addedFieldNames: { [name: string]: boolean } = {};
        for (var v = 0; v < viewFieldNames.length; v += 1) {
          var viewFieldName = viewFieldNames[v];
          if (SYSTEM_FIELDS[viewFieldName]) { continue; }
          var matchedSharePointField: ISharePointField | undefined;
          for (var s = 0; s < this.state.sharePointFields.length; s += 1) {
            if (this.state.sharePointFields[s].internalName.toLowerCase() === viewFieldName.toLowerCase()) {
              matchedSharePointField = this.state.sharePointFields[s];
              break;
            }
          }
          if (!matchedSharePointField) { continue; }
          // A SharePoint view can list the same field more than once; only add it once.
          if (addedFieldNames[matchedSharePointField.internalName.toLowerCase()]) { continue; }
          addedFieldNames[matchedSharePointField.internalName.toLowerCase()] = true;
          nextFields.push({
            id: createId(),
            fieldName: matchedSharePointField.internalName,
            displayName: matchedSharePointField.title,
            width: ''
          });
        }
        if (nextFields.length > 0) {
          this.setState({ fields: nextFields, selectedFieldId: nextFields[0].id });
        }
      }

      var groupField1 = groupByResult.fieldNames.length > 0 ? groupByResult.fieldNames[0] : '';
      var groupField2 = groupByResult.fieldNames.length > 1 ? groupByResult.fieldNames[1] : '';
      this.setState({
        viewLoading: false,
        grouping: {
          field1: groupField1,
          field2: groupField2,
          collapsedByDefault: groupByResult.collapsed,
          showCount: this.state.grouping.showCount
        },
        viewLoadMessage: viewFieldNames.length > 0
          ? (groupField1 ? 'Loaded columns and grouping from the selected view.' : 'Loaded columns from the selected view. The view has no grouping configured.')
          : (groupField1 ? 'Loaded grouping from the selected view.' : 'The selected view has no fields or grouping to load.')
      });
    } catch (error) {
      this.setState({ viewLoading: false, viewLoadMessage: error && error.message ? error.message : 'Failed to load the selected view.' });
    }
  }

  private getAvailableFields(): ISharePointField[] {
    var used: { [name: string]: boolean } = {};
    this.state.fields.forEach(function(field) { used[String(field.fieldName || '').toLowerCase()] = true; });
    return this.state.sharePointFields.filter(function(field) { return !used[field.internalName.toLowerCase()]; });
  }

  private addField(source: ISharePointField): void {
    var field: IListDesignerField = {
      id: createId(),
      fieldName: source.internalName,
      displayName: source.title,
      width: ''
    };
    var fields = copyFields(this.state.fields);
    fields.push(field);
    this.setState({ fields: fields, selectedFieldId: field.id });
  }

  private getSelectedField(): IListDesignerField | undefined {
    for (var i = 0; i < this.state.fields.length; i += 1) {
      if (this.state.fields[i].id === this.state.selectedFieldId) {
        return this.state.fields[i];
      }
    }
    return undefined;
  }

  private updateSelectedField(mutator: (field: IListDesignerField) => void): void {
    var fields = copyFields(this.state.fields);
    for (var i = 0; i < fields.length; i += 1) {
      if (fields[i].id === this.state.selectedFieldId) {
        mutator(fields[i]);
        break;
      }
    }
    this.setState({ fields: fields });
  }

  private moveField(fieldId: string, direction: number): void {
    var fields = copyFields(this.state.fields);
    var index = -1;
    for (var i = 0; i < fields.length; i += 1) {
      if (fields[i].id === fieldId) { index = i; break; }
    }
    var target = index + direction;
    if (index < 0 || target < 0 || target >= fields.length) { return; }
    var moved = fields[index];
    fields[index] = fields[target];
    fields[target] = moved;
    this.setState({ fields: fields });
  }

  private removeField(fieldId: string): void {
    var fields = this.state.fields.filter(function(field) { return field.id !== fieldId; });
    this.setState({
      fields: fields,
      selectedFieldId: this.state.selectedFieldId === fieldId ? (fields.length > 0 ? fields[0].id : '') : this.state.selectedFieldId
    });
  }

  private updateGrouping(mutator: (grouping: IListGroupingConfig) => void): void {
    var grouping: IListGroupingConfig = {
      field1: this.state.grouping.field1,
      field2: this.state.grouping.field2,
      collapsedByDefault: this.state.grouping.collapsedByDefault,
      showCount: this.state.grouping.showCount
    };
    mutator(grouping);
    if (!grouping.field1) { grouping.field2 = ''; }
    this.setState({ grouping: grouping });
  }

  private getFieldLabel(fieldName: string): string {
    if (!fieldName) { return ''; }
    for (var i = 0; i < this.state.fields.length; i += 1) {
      if (this.state.fields[i].fieldName.toLowerCase() === fieldName.toLowerCase()) {
        return this.state.fields[i].displayName || this.state.fields[i].fieldName;
      }
    }
    for (var j = 0; j < this.state.sharePointFields.length; j += 1) {
      if (this.state.sharePointFields[j].internalName.toLowerCase() === fieldName.toLowerCase()) {
        return this.state.sharePointFields[j].title || fieldName;
      }
    }
    return fieldName;
  }

  private save(): void {
    if (this.state.loading) {
      return;
    }
    var columns: IListDesignerColumn[] = this.state.fields.map(function(field) {
      return { fieldName: field.fieldName, displayName: field.displayName, width: field.width || '' };
    });
    var grouping = this.state.grouping;
    var groupingJson = JSON.stringify({
      enabled: !!grouping.field1,
      field1: grouping.field1,
      field2: grouping.field1 ? grouping.field2 : '',
      collapsedByDefault: grouping.collapsedByDefault,
      showCount: grouping.showCount
    });
    this.props.onSave(columns, groupingJson);
  }

  private renderPalette(): JSX.Element {
    var available = this.getAvailableFields();
    return (
      <aside className="gd-panel gd-palette">
        <h3>SharePoint fields</h3>
        <p>Add fields as list columns.</p>
        {this.state.loading && <div className="gd-empty">Loading fields...</div>}
        {this.state.error && <div className="gd-error">{this.state.error}</div>}
        {!this.state.loading && !this.state.error && available.length === 0 && <div className="gd-empty">All available fields are in the list.</div>}
        {available.map((field) => (
          <button key={field.internalName} type="button" className="gd-palette-item" onClick={() => this.addField(field)}>
            <span>{field.title}</span>
            <small>{field.internalName}</small>
          </button>
        ))}
      </aside>
    );
  }

  private renderGroupingSettings(): JSX.Element {
    var grouping = this.state.grouping;
    return (
      <div className="gd-group-settings">
        <div className="gd-canvas-header">
          <div><h3>Grouping</h3><p>Group list rows by up to two columns, similar to a SharePoint view's Group By.</p></div>
        </div>
        {this.props.viewId && (
          <div className="gd-group-view-load">
            <button type="button" disabled={this.state.viewLoading} onClick={() => this.loadFromView()}>
              {this.state.viewLoading ? 'Loading from view…' : 'Load columns & grouping from selected view'}
            </button>
            {this.state.viewLoadMessage && <div className="gd-hint">{this.state.viewLoadMessage}</div>}
          </div>
        )}
        <div className="gd-inline-fields gd-inline-fields-two">
          <label>Group by<select value={grouping.field1} onChange={(ev) => { var value = ev.currentTarget.value; this.updateGrouping(function(next) { next.field1 = value; }); }}>
            <option value="">None</option>
            {this.state.fields.map(function(field) { return <option key={field.id} value={field.fieldName}>{field.displayName || field.fieldName}</option>; })}
          </select></label>
          <label>Then by<select value={grouping.field2} disabled={!grouping.field1} onChange={(ev) => { var value = ev.currentTarget.value; this.updateGrouping(function(next) { next.field2 = value; }); }}>
            <option value="">None</option>
            {this.state.fields.filter(function(field) { return field.fieldName !== grouping.field1; }).map(function(field) { return <option key={field.id} value={field.fieldName}>{field.displayName || field.fieldName}</option>; })}
          </select></label>
        </div>
        <label className="gd-check"><input type="checkbox" checked={grouping.collapsedByDefault} disabled={!grouping.field1} onChange={(ev) => { var checked = ev.currentTarget.checked; this.updateGrouping(function(next) { next.collapsedByDefault = checked; }); }} /> Collapse groups by default</label>
        <label className="gd-check"><input type="checkbox" checked={grouping.showCount} disabled={!grouping.field1} onChange={(ev) => { var checked = ev.currentTarget.checked; this.updateGrouping(function(next) { next.showCount = checked; }); }} /> Show item count per group</label>
      </div>
    );
  }

  private renderColumnsList(): JSX.Element {
    return (
      <div>
        {this.state.fields.length === 0 && <div className="gd-empty gd-empty-canvas">Add SharePoint fields from the left panel.</div>}
        {this.state.fields.map((field, index) => (
          <div key={field.id} className={'gd-column ' + (field.id === this.state.selectedFieldId ? 'gd-column-selected' : '')} onClick={() => this.setState({ selectedFieldId: field.id })}>
            <div className="gd-column-order">{index + 1}</div>
            <div className="gd-column-main">
              <strong>{field.displayName || field.fieldName}</strong>
              <span>{field.fieldName}</span>
            </div>
            <div className="gd-column-actions">
              <button type="button" title="Move up" disabled={index === 0} onClick={(ev) => { ev.stopPropagation(); this.moveField(field.id, -1); }}>↑</button>
              <button type="button" title="Move down" disabled={index === this.state.fields.length - 1} onClick={(ev) => { ev.stopPropagation(); this.moveField(field.id, 1); }}>↓</button>
              <button type="button" title="Remove column" onClick={(ev) => { ev.stopPropagation(); this.removeField(field.id); }}>×</button>
            </div>
          </div>
        ))}
      </div>
    );
  }

  private renderCanvas(): JSX.Element {
    var grouping = this.state.grouping;
    var columnsPreview = this.renderColumnsList();
    if (grouping.field1) {
      var innerPreview = grouping.field2 ? (
        <div className="gd-group-container gd-group-container-nested">
          <div className="gd-group-container-header">
            <span className="gd-group-container-icon">▾</span>
            <span>Then by: {this.getFieldLabel(grouping.field2)} = <em>(example value)</em></span>
            {grouping.showCount && <span className="gd-group-container-count">3 items</span>}
          </div>
          <div className="gd-group-container-body">{columnsPreview}</div>
        </div>
      ) : columnsPreview;
      columnsPreview = (
        <div className="gd-group-container">
          <div className="gd-group-container-header">
            <span className="gd-group-container-icon">{grouping.collapsedByDefault ? '▸' : '▾'}</span>
            <span>Grouped by: {this.getFieldLabel(grouping.field1)} = <em>(example value)</em></span>
            {grouping.showCount && <span className="gd-group-container-count">{grouping.field2 ? '' : '5 items'}</span>}
          </div>
          <div className="gd-group-container-body">{innerPreview}</div>
        </div>
      );
    }
    return (
      <main className="gd-canvas">
        {this.renderGroupingSettings()}
        <div className="gd-canvas-header">
          <div><h3>List columns</h3><p>Column order follows this list from top to bottom.</p></div>
          <span>{this.state.fields.length} columns</span>
        </div>
        {columnsPreview}
      </main>
    );
  }

  private renderFieldEditor(): JSX.Element {
    var field = this.getSelectedField();
    if (!field) {
      return <aside className="gd-panel gd-properties"><div className="gd-empty">Select a list column to configure it.</div></aside>;
    }
    return (
      <aside className="gd-panel gd-properties">
        <h3>Column settings</h3>
        <label>Column label<input type="text" value={field.displayName || ''} onChange={(ev) => this.updateSelectedField(function(next) { next.displayName = ev.currentTarget.value; })} /></label>
        <label>SharePoint field<div className="gd-readonly-value">{field.fieldName}</div></label>
        <label>Column width<input type="text" value={field.width || ''} placeholder="180px or 20%" onChange={(ev) => this.updateSelectedField(function(next) { next.width = ev.currentTarget.value; })} /></label>
      </aside>
    );
  }

  public render(): JSX.Element {
    return (
      <div className="gd-overlay">
        <header className="gd-toolbar">
          <div><h2>List Designer</h2><span>{this.props.listName || 'No list selected'}</span></div>
          <div className="gd-toolbar-actions"><button type="button" onClick={this.props.onCancel}>Cancel</button><button type="button" className="gd-primary" onClick={() => this.save()} disabled={!this.props.listName || this.state.loading}>Save list</button></div>
        </header>
        <div className="gd-workspace">{this.renderPalette()}{this.renderCanvas()}{this.renderFieldEditor()}</div>
      </div>
    );
  }
}
