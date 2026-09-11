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
    if (prevProps.reportProps.listName !== this.props.reportProps.listName) {
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
      var itemIds: number[] = [];
      var seenItemIds: { [itemId: string]: boolean } = {};
      var pageCount = 0;

      while (nextUrl && pageCount < 1000) {
        var response = await this.getWithAcceptFallback(nextUrl);
        if (!response.ok) {
          throw new Error(strings.RuntimeRepeatLoadFailed + ' Status: ' + String(response.status));
        }
        var data = await response.json();
        var items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
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
            itemIdQueryParam: '__spse_repeat_item_id_disabled__',
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