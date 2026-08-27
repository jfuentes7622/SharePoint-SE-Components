import * as React from 'react';
import styles from './Marquee.module.scss';
import { IKmMarqueeProps } from './IMarqueeProps';
import {DisplayMode} from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'KmMarqueeWebPartStrings';
import './Marquee.css';

export interface IKmMarqueeState {
  loading: boolean;
  messages: string[];
  currentIndex: number;
}

interface IRenderViewData {
  rows: any[];
  fields: any[];
}

function escapeODataText(value: string): string {
  return value.replace(/'/g, "''");
}

function trimGuidBraces(value: string): string {
  return String(value || '').replace(/[{}]/g, '');
}

function buildViewIdCandidates(viewId: string): string[] {
  const normalized = trimGuidBraces(viewId);
  const candidates = [String(viewId || ''), normalized, '{' + normalized + '}'];
  return candidates.filter((candidate: string, index: number) => {
    return !!candidate && candidates.indexOf(candidate) === index;
  });
}

function toArray(value: any): any[] {
  if (!value) {
    return [];
  }
  if (typeof value === 'string') {
    try {
      value = JSON.parse(value);
    } catch (_error) {
      return [];
    }
  }
  if (Array.isArray(value)) {
    return value;
  }
  return Array.isArray(value.results) ? value.results : [];
}

function tryParseObject(value: any): any {
  if (!value || typeof value !== 'string') {
    return value;
  }
  try {
    return JSON.parse(value);
  } catch (_error) {
    return value;
  }
}

function parseRenderViewData(data: any): IRenderViewData {
  const directData = data && data.d && data.d.RenderListDataAsStream
    ? data.d.RenderListDataAsStream : data;
  const responseData = tryParseObject(directData) || {};
  const listData = tryParseObject(responseData.ListData) || responseData;
  const schema = tryParseObject(responseData.ListSchema) || tryParseObject(responseData.Schema) || {};

  let rows = toArray(responseData && responseData.Row);
  if (rows.length === 0) {
    rows = toArray(responseData && responseData.Rows);
  }
  if (rows.length === 0) {
    rows = toArray(listData && listData.Row);
  }
  if (rows.length === 0) {
    rows = toArray(listData && listData.Rows);
  }
  if (rows.length === 0) {
    rows = toArray(data && data.d && data.d.Row);
  }
  let fields = toArray(schema && schema.Field);
  if (fields.length === 0) {
    fields = toArray(responseData && responseData.Field);
  }
  return { rows: rows, fields: fields };
}

function getFieldCandidates(fieldName: string, fields: any[]): string[] {
  const candidates = [fieldName];
  const normalized = String(fieldName || '').toLowerCase();
  fields.forEach((field: any) => {
    const responseName = String(field.Name || '');
    const realName = String(field.RealFieldName || '');
    if (responseName.toLowerCase() === normalized || realName.toLowerCase() === normalized) {
      if (responseName && candidates.indexOf(responseName) < 0) {
        candidates.push(responseName);
      }
      if (realName && candidates.indexOf(realName) < 0) {
        candidates.push(realName);
      }
    }
  });
  if ((normalized === 'linktitle' || normalized === 'linktitlenomenu') && candidates.indexOf('Title') < 0) {
    candidates.push('Title');
  }
  if ((normalized === 'linkfilename' || normalized === 'linkfilename2') && candidates.indexOf('FileLeafRef') < 0) {
    candidates.push('FileLeafRef');
  }
  return candidates;
}

function getMessageValue(item: any, candidates: string[]): string {
  for (let candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
    const candidate = candidates[candidateIndex];
    let value = item[candidate];
    if (value === undefined || value === null || value === '') {
      const itemKeys = Object.keys(item);
      const matchingKey = itemKeys.filter((key: string) => { return key.toLowerCase() === candidate.toLowerCase(); })[0];
      value = matchingKey ? item[matchingKey] : value;
    }
    if (value !== undefined && value !== null && String(value).trim().length > 0) {
      return String(value);
    }
  }
  return '';
}

function getItemId(item: any): string {
  const candidateKeys = ['ID', 'Id', 'id', 'ID.'];
  for (let keyIndex = 0; keyIndex < candidateKeys.length; keyIndex += 1) {
    const value = item[candidateKeys[keyIndex]];
    if (value !== undefined && value !== null && String(value)) {
      return String(value);
    }
  }
  return '';
}

export default class KmMarquee extends React.Component<IKmMarqueeProps, IKmMarqueeState> {
  private _cycleTimer: number;

  constructor(props: IKmMarqueeProps) {
    super(props);
    this.state = { loading: false, messages: [], currentIndex: 0 };
    this.handleLoad = this.handleLoad.bind(this);
   }

 componentWillMount() {
    this.logDiagnostic('componentWillMount');
    this.handleLoad();
  }

 componentDidMount() {
  this.logDiagnostic('componentDidMount');
  setTimeout(() => {window.addEventListener('load', this.handleLoad);},0);
  this.loadMessages();
}

  componentDidUpdate(prevProps: IKmMarqueeProps):void {
    this.logDiagnostic('componentDidUpdate');
    if (prevProps.listName !== this.props.listName || prevProps.viewId !== this.props.viewId
      || prevProps.messageField !== this.props.messageField) {
      this.loadMessages();
      return;
    }
    if (prevProps.messageDuration !== this.props.messageDuration) {
      this.scheduleNextPass();
    }
    if (
      prevProps.description !== this.props.description
      || prevProps.marqueeTextColor !== this.props.marqueeTextColor
      || prevProps.marqueeBackColor !== this.props.marqueeBackColor
      || prevProps.marqueeActive !== this.props.marqueeActive
      || prevProps.fontFamily !== this.props.fontFamily
      || prevProps.fontSize !== this.props.fontSize
      || prevProps.fontStyle !== this.props.fontStyle
      || prevProps.fontBold !== this.props.fontBold
      || prevProps.marqueeHeight !== this.props.marqueeHeight
      || prevProps.scrollSpeed !== this.props.scrollSpeed
      || prevProps.scrollDirection !== this.props.scrollDirection
      || prevProps.placement !== this.props.placement
    ) {
      this.handleLoad();
      this.restartAnimation();
      this.scheduleNextPass();
    }
  }


  componentWillUnmount() { 
    this.logDiagnostic('componentWillUnmount');
    window.removeEventListener('load', this.handleLoad);
    this.clearCycle();
    //this.removeDiv();
  }

  private logDiagnostic(message: string): void {
    if (this.props.enableDiagnostics === false) {
      return;
    }
    console.log('[KmMarquee] ' + message);
  }

  private clearCycle(): void {
    if (this._cycleTimer) {
      window.clearTimeout(this._cycleTimer);
      this._cycleTimer = undefined;
    }
  }

  private scheduleNextPass(): void {
    this.clearCycle();
    const messageSeconds = typeof this.props.messageDuration === 'number' && this.props.messageDuration >= 2
      ? this.props.messageDuration : 8;
    const durationMs = messageSeconds * 1000;
    this._cycleTimer = window.setTimeout(() => {
      if (this.state.messages.length > 1) {
        this.setState((prevState: IKmMarqueeState) => {
          const nextIndex = (prevState.currentIndex + 1) % prevState.messages.length;
          return { currentIndex: nextIndex };
        }, () => {
          this.logDiagnostic('Showing message ' + String(this.state.currentIndex + 1)
            + ' of ' + String(this.state.messages.length) + '.');
          this.handleLoad();
          this.restartAnimation();
          this.scheduleNextPass();
        });
      } else {
        this.restartAnimation();
        this.scheduleNextPass();
      }
    }, durationMs);
  }

  private restartAnimation(): void {
    const marqueeSpan = document.getElementById('marqueeSpan') as HTMLElement;
    if (!marqueeSpan) {
      return;
    }
    marqueeSpan.style.animationName = 'none';
    window.setTimeout(() => {
      marqueeSpan.style.animationName = this.props.scrollDirection === 'right' ? 'moveRight' : 'moveLeft';
    }, 0);
  }

  private async loadMessages(): Promise<void> {
    if (!this.props.listName || !this.props.viewId || !this.props.messageField) {
      this.clearCycle();
      this.setState({ messages: [], currentIndex: 0 }, () => {
        this.handleLoad();
        this.restartAnimation();
        this.scheduleNextPass();
      });
      return;
    }

    try {
      this.logDiagnostic('Loading messages from list "' + this.props.listName + '", selected view, field "' + this.props.messageField + '"');
      const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const endpoint = webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/RenderListDataAsStream";
      const viewXml = await this.loadSelectedViewXml();
      const response = await this.postViewRequest(endpoint, viewXml);
      if (!response.ok) {
        throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
      }

      const data = await response.json();
      const renderData = parseRenderViewData(data);
      const selectedRows = renderData.rows;
      const selectedFieldCandidates = getFieldCandidates(this.props.messageField, renderData.fields);
      let messages: string[] = [];
      messages = renderData.rows
        .map((item: any) => { return getMessageValue(item, selectedFieldCandidates); })
        .filter((message: string) => { return message.trim().length > 0; });
      this.logDiagnostic('Authoritative selected view returned rows=' + String(renderData.rows.length)
        + ', messages=' + String(messages.length) + ', field candidates=' + selectedFieldCandidates.join(',') + '.');
      if (messages.length === 0) {
        messages = await this.loadMessagesDirectly(selectedRows, selectedFieldCandidates);
      }
      if (messages.length === 0) {
        messages = await this.loadMessagesAsText(selectedRows, selectedFieldCandidates);
      }
      this.logDiagnostic('Loaded messages successfully. Count=' + String(messages.length));
      this.setState({ messages: messages, currentIndex: 0 }, () => {
        this.handleLoad();
        this.restartAnimation();
        this.scheduleNextPass();
      });
    } catch (error) {
      console.error('[KmMarquee] Failed to load messages from list "' + this.props.listName + '": ' + (error && error.message ? error.message : String(error)));
      this.setState({ messages: [], currentIndex: 0 }, () => { this.handleLoad(); });
    }
  }

  private async loadMessagesDirectly(viewRows: any[], fieldCandidates: string[]): Promise<string[]> {
    const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const itemIds = viewRows.map((row: any) => { return getItemId(row); })
      .filter((itemId: string, index: number, values: string[]) => {
        return !!itemId && values.indexOf(itemId) === index;
      });
    if (itemIds.length === 0) {
      return [];
    }

    for (let candidateIndex = 0; candidateIndex < fieldCandidates.length; candidateIndex += 1) {
      const fieldName = fieldCandidates[candidateIndex];
      let endpoint = webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items"
        + '?$select=' + encodeURIComponent('ID,' + fieldName) + '&$orderby=ID&$top=500';
      if (itemIds.length > 0) {
        endpoint += '&$filter=' + encodeURIComponent(itemIds.map((itemId: string) => {
          return 'ID eq ' + String(parseInt(itemId, 10));
        }).join(' or '));
      }
      const response = await this.getJsonResponse(endpoint);
      if (!response.ok) {
        continue;
      }
      const data = await response.json();
      const items = toArray(data && data.value).length > 0
        ? toArray(data.value) : toArray(data && data.d && data.d.results);
      const messageById: { [itemId: string]: string } = {};
      const unorderedMessages: string[] = [];
      items.forEach((item: any) => {
        const message = getMessageValue(item, [fieldName]);
        const itemId = getItemId(item);
        if (itemId && message) {
          messageById[itemId] = message;
        }
        if (message) {
          unorderedMessages.push(message);
        }
      });
      const orderedMessages = itemIds.map((itemId: string) => { return messageById[itemId] || ''; })
        .filter((message: string) => { return message.trim().length > 0; });
      const messages = itemIds.length > 0 ? orderedMessages : unorderedMessages;
      this.logDiagnostic('Direct items fallback field=' + fieldName + ', messages=' + String(messages.length)
        + ', selected view item IDs=' + String(itemIds.length) + '.');
      if (messages.length > 0) {
        return messages;
      }
    }
    return [];
  }

  private async loadMessagesAsText(viewRows: any[], fieldCandidates: string[]): Promise<string[]> {
    const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const itemIds = viewRows.map((row: any) => { return getItemId(row); })
      .filter((itemId: string, index: number, values: string[]) => {
        return !!itemId && values.indexOf(itemId) === index;
      });
    if (itemIds.length === 0) {
      return [];
    }
    let endpoint = webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items"
      + '?$select=ID,FieldValuesAsText&$expand=FieldValuesAsText&$top=500';
    if (itemIds.length > 0) {
      endpoint += '&$filter=' + encodeURIComponent(itemIds.map((itemId: string) => {
        return 'ID eq ' + String(parseInt(itemId, 10));
      }).join(' or '));
    }

    const response = await this.getJsonResponse(endpoint);
    if (!response.ok) {
      this.logDiagnostic('FieldValuesAsText fallback failed. HTTP ' + String(response.status) + '.');
      return [];
    }
    const data = await response.json();
    const items = toArray(data && data.value).length > 0
      ? toArray(data.value) : toArray(data && data.d && data.d.results);
    const textById: { [itemId: string]: string } = {};
    const unorderedMessages: string[] = [];
    items.forEach((item: any) => {
      const values = item.FieldValuesAsText || {};
      const message = getMessageValue(values, fieldCandidates);
      const itemId = getItemId(item);
      if (itemId && message) {
        textById[itemId] = message;
      }
      if (message) {
        unorderedMessages.push(message);
      }
    });

    const orderedMessages = itemIds.map((itemId: string) => { return textById[itemId] || ''; })
      .filter((message: string) => { return message.trim().length > 0; });
    const messages = itemIds.length > 0 ? orderedMessages : unorderedMessages;
    this.logDiagnostic('FieldValuesAsText fallback loaded messages=' + String(messages.length)
      + ', selected view item IDs=' + String(itemIds.length) + '.');
    return messages;
  }

  private async getJsonResponse(url: string): Promise<any> {
    let response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=verbose' }
      });
    }
    if (!response.ok) {
      response = await this.props.spfxContext.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });
    }
    return response;
  }

  private buildMinimalViewXml(viewQuery: string, rowLimit: number, scope: any): string {
    const queryText = String(viewQuery || '').trim();
    const queryDocument = new DOMParser().parseFromString(
      /^<Query(?:\s|>)/i.test(queryText) ? queryText : '<Query>' + queryText + '</Query>',
      'text/xml'
    );
    if (queryDocument.getElementsByTagName('parsererror').length > 0 || queryDocument.getElementsByTagName('Query').length === 0) {
      throw new Error('The selected SharePoint view query is invalid.');
    }

    const xmlDocument = new DOMParser().parseFromString('<View><Query/><ViewFields/></View>', 'text/xml');
    const viewElement = xmlDocument.getElementsByTagName('View')[0];
    const existingQuery = xmlDocument.getElementsByTagName('Query')[0];
    const queryElement = xmlDocument.importNode(queryDocument.getElementsByTagName('Query')[0], true);
    viewElement.replaceChild(queryElement, existingQuery);

    const viewFieldsElement = xmlDocument.getElementsByTagName('ViewFields')[0];
    ['ID', this.props.messageField].forEach((fieldName: string) => {
      if (fieldName) {
        const fieldRef = xmlDocument.createElement('FieldRef');
        fieldRef.setAttribute('Name', fieldName);
        viewFieldsElement.appendChild(fieldRef);
      }
    });

    const normalizedScope = String(scope === undefined || scope === null ? '' : scope).toLowerCase();
    const scopeNames: { [key: string]: string } = {
      '1': 'Recursive',
      '2': 'RecursiveAll',
      '3': 'FilesOnly',
      'recursive': 'Recursive',
      'recursiveall': 'RecursiveAll',
      'filesonly': 'FilesOnly'
    };
    if (scopeNames[normalizedScope]) {
      viewElement.setAttribute('Scope', scopeNames[normalizedScope]);
    }

    if (rowLimit > 0) {
      const rowLimitElement = xmlDocument.createElement('RowLimit');
      rowLimitElement.setAttribute('Paged', 'TRUE');
      rowLimitElement.appendChild(xmlDocument.createTextNode(String(rowLimit)));
      viewElement.appendChild(rowLimitElement);
    }
    return new XMLSerializer().serializeToString(xmlDocument);
  }

  private async loadSelectedViewXml(): Promise<string> {
    const webUrl = this.props.spfxContext.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
    const viewIds = buildViewIdCandidates(this.props.viewId);
    const urls: string[] = [];
    viewIds.forEach((viewId: string) => {
      const encoded = encodeURIComponent(viewId);
      const normalized = encodeURIComponent(trimGuidBraces(viewId));
      urls.push(webUrl + listPath + "/views/getById('" + encoded + "')?$select=ViewQuery,RowLimit,Scope");
      urls.push(webUrl + listPath + "/views(guid'" + normalized + "')?$select=ViewQuery,RowLimit,Scope");
    });

    for (let urlIndex = 0; urlIndex < urls.length; urlIndex += 1) {
      const response = await this.getJsonResponse(urls[urlIndex]);
      if (!response.ok) {
        continue;
      }
      const data = await response.json();
      const viewData = data && data.d ? data.d : data;
      if (viewData && viewData.ViewQuery !== undefined && viewData.ViewQuery !== null) {
        const parsedRowLimit = parseInt(String(viewData.RowLimit || ''), 10);
        const authoritativeViewXml = this.buildMinimalViewXml(String(viewData.ViewQuery), isNaN(parsedRowLimit) ? 0 : parsedRowLimit, viewData.Scope);
        this.logDiagnostic('Loaded minimal selected view CAML. HasFilter=' + String(/<Where(?:\s|>)/i.test(authoritativeViewXml))
          + ', hasSort=' + String(/<OrderBy(?:\s|>)/i.test(authoritativeViewXml)) + '.');
        return authoritativeViewXml;
      }
    }
    throw new Error('Failed to load the selected SharePoint view definition.');
  }

  private async postViewRequest(url: string, viewXml: string): Promise<any> {
    const body = JSON.stringify({ parameters: { RenderOptions: 7, ViewXml: viewXml } });
    const formats = [
      { 'Content-Type': 'application/json; charset=utf-8' },
      { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' },
      { Accept: 'application/json;odata=minimalmetadata', 'Content-Type': 'application/json;odata=minimalmetadata' },
      { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' }
    ];
    let response: any;
    for (let formatIndex = 0; formatIndex < formats.length; formatIndex += 1) {
      response = await this.props.spfxContext.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: formats[formatIndex],
        body: body
      });
      if (response.ok) {
        break;
      }
    }
    return response;
  }

  private getCurrentMessage(): string {
    if (this.state.messages.length > 0) {
      return this.state.messages[this.state.currentIndex] || '';
    }
    return this.props.description || '';
  }

  private handleLoad() {
    this.logDiagnostic('marqueeActive:' + this.props.marqueeActive);
    if (this.props.marqueeActive) {
      this.insertDiv();
      this.setState({ loading: true }); }
    else {
      this.removeDiv();
      this.setState({ loading: false });
    }
   }
  
   
    private getParentElement(parentFallbackList: string): HTMLElement {
      // check for null or empty fallback list (only whitespace characters present)
      if (!parentFallbackList ||
          parentFallbackList.match(/^\s*$/)) {
          this.logDiagnostic('WARNING: empty fallback list supplied; defaulting to document body');
          return document.body;
      }
      // parse string for fallback list
      const fallbackList: string[] = parentFallbackList.split('|');
      // declare null parent element
      let parent: HTMLElement | undefined = undefined;
      // get the first valid element
      fallbackList.forEach(element=> {
          try {
              const queryElement = document.querySelector(element) as HTMLElement;
              parent = parent || queryElement;
          } catch (err) {
              console.error('[KmMarquee] ' + err);
          }
      });
      // return the proper parent, or the body element if none of the fallbacks succeed
      return parent || document.body;
  }

  public removeDiv() {    
    this.logDiagnostic('In removeDiv');
    const marqueeDiv: HTMLElement = document.getElementById('marqueeDiv') as HTMLElement;
    const marqueeSpan: HTMLElement=document.getElementById('marqueeSpan') as HTMLElement;
    this.logDiagnostic('marqueeDiv found:' + marqueeDiv);
    // check for existing marqueeDiv and span and remove
    if (marqueeDiv) {marqueeDiv.remove();}
    if(marqueeSpan) {marqueeSpan.remove();}
    } 

 private findElementByIdOrClassToken(token: string): HTMLElement {
    const normalizedToken = token.toLowerCase();
    const candidates = document.querySelectorAll('[id], [class]');
    for (let index = 0; index < candidates.length; index += 1) {
      const candidate = candidates[index] as HTMLElement;
      const id = String(candidate.id || '').toLowerCase();
      const className = String(candidate.className || '').toLowerCase();
      if (id === normalizedToken || id.indexOf(normalizedToken) >= 0
        || className.split(/\s+/).some((classToken: string) => classToken === normalizedToken)) {
        return candidate;
      }
    }
    return undefined;
  }
   
 private placeMarquee(marqueeDiv: HTMLElement): void {
    const placement = this.props.placement || 'top';
    let anchor: HTMLElement;
    if (placement === 'aboveChrome') {
      anchor = this.findElementByIdOrClassToken('SPPageChrome');
    } else if (placement === 'belowChrome') {
      anchor = this.findElementByIdOrClassToken('spAppAndPropertyPanelContainer');
    } else if (placement === 'aboveContent') {
      anchor = document.querySelector('div[class*="pageContainer_"][class*="container_"]') as HTMLElement;
      if (anchor) {
        anchor.insertBefore(marqueeDiv, anchor.firstChild);
        this.logDiagnostic('Placed marquee above content using dynamic class prefixes.');
        return;
      }
    }
    if (!anchor && placement !== 'top') {
      this.logDiagnostic('Placement target for "' + placement + '" was not found; using Top.');
    }
    anchor = anchor || this.getParentElement(strings.ParentElement);
    if (anchor.parentNode) {
      anchor.parentNode.insertBefore(marqueeDiv, anchor);
    }
    this.logDiagnostic('Placed marquee at "' + placement + '".');
  }

 public insertDiv() {
    const displayText: string = this.getCurrentMessage();
    this.logDiagnostic('In insertDiv');
    let marqueeDiv: HTMLElement = document.getElementById('marqueeDiv') as HTMLElement;
    let marqueeSpan: HTMLElement= document.getElementById('marqueeSpan') as HTMLElement;
    // check for existing marqueeDiv
    if (!marqueeDiv) {
      marqueeDiv = document.createElement('div');
      marqueeDiv.id = 'marqueeDiv';
      marqueeDiv.className = 'bar';
      const marqueeSpan2 = document.createElement('span');
      marqueeSpan2.className='bar_content';
      marqueeSpan2.id='marqueeSpan';
      marqueeDiv.appendChild(marqueeSpan2);
      marqueeSpan = marqueeSpan2;
    }
    this.placeMarquee(marqueeDiv);
    this.applyStyles(marqueeDiv, marqueeSpan);
    marqueeSpan.textContent = displayText;
  }

  private applyStyles(marqueeDiv: HTMLElement, marqueeSpan: HTMLElement): void {
    marqueeDiv.style.color = this.props.marqueeTextColor;
    marqueeDiv.style.backgroundColor = this.props.marqueeBackColor;
    marqueeDiv.style.minHeight = this.props.marqueeHeight || '32px';
    marqueeDiv.style.display = 'flex';
    marqueeDiv.style.alignItems = 'center';
    marqueeDiv.style.boxSizing = 'border-box';

    marqueeSpan.style.fontFamily = this.props.fontFamily || '"Segoe UI", sans-serif';
    marqueeSpan.style.fontSize = this.props.fontSize || '13px';
    marqueeSpan.style.fontStyle = this.props.fontStyle || 'normal';
    marqueeSpan.style.fontWeight = this.props.fontBold !== false ? 'bold' : 'normal';

    const speed = typeof this.props.scrollSpeed === 'number' && this.props.scrollSpeed > 0 ? this.props.scrollSpeed : 50;
    const isRightward = this.props.scrollDirection === 'right';
    marqueeSpan.style.animationDuration = speed + 's';
    marqueeSpan.style.animationTimingFunction = 'linear';
    marqueeSpan.style.animationIterationCount = '1';
    marqueeSpan.style.animationName = isRightward ? 'moveRight' : 'moveLeft';
    marqueeSpan.style.transform = isRightward ? 'translateX(-100%)' : 'translateX(100%)';
  }

  public openPane(): void {
    this.props.spfxContext.propertyPane.open();
  }

  public render(): React.ReactElement<IKmMarqueeProps> {
    if (DisplayMode.Edit===this.props.dMode || (this.props.description===undefined && !this.props.listName)) {
      return (
        <div className={styles.kmMarquee}>
          <div className="col">
            <p className="$ms-color-themePrimary ms-font-xl">Marquee Web Part</p>
            <button className={styles.button} onClick={this.openPane.bind(this)}>Click here to configure</button>
          </div>
        </div>
      );
    }

    if (this.getCurrentMessage().length>0 && DisplayMode.Read===this.props.dMode) {
    return( 
      <div className={styles.kmMarquee}>
      <div className="col">
       <button className={styles.buttonH}>Hidden button</button>
      </div>    
      </div >
    );
     }
     return (
      <div></div>
     );
  }
}
