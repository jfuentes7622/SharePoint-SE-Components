import * as React from 'react';
import * as strings from 'SharePointDynamicFormWebPartStrings';

export interface IListControlHostProps {
  fieldId: string;
  sourceId: string;
  sourceName: string;
  sourceListName: string;
  runtimeFilterJson: string;
  isPageEditMode: boolean;
}

interface IListControlHostState {
  status: 'waiting' | 'embedded' | 'missing';
}

export class ListControlHost extends React.Component<IListControlHostProps, IListControlHostState> {
  private _mountElement: HTMLDivElement | null;
  private _targetHost: HTMLElement | null;
  private _originalParent: Node | null;
  private _originalNextSibling: Node | null;
  private _configurationMarker: HTMLDivElement | null;
  private _retryHandle: number | undefined;
  private _attemptCount: number;
  private _runtimeConfigRequestHandler: any;

  public constructor(props: IListControlHostProps) {
    super(props);
    this.state = { status: 'waiting' };
    this._mountElement = null;
    this._targetHost = null;
    this._originalParent = null;
    this._originalNextSibling = null;
    this._configurationMarker = null;
    this._retryHandle = undefined;
    this._attemptCount = 0;
    this._runtimeConfigRequestHandler = this.handleRuntimeConfigRequest.bind(this);
  }

  public componentDidMount(): void {
    if (typeof window !== 'undefined' && window.addEventListener) {
      window.addEventListener('spse:listcontrol-runtime-config-request', this._runtimeConfigRequestHandler);
    }
    this.startDiscovery();
  }

  public componentDidUpdate(prevProps: IListControlHostProps): void {
    if (prevProps.sourceId !== this.props.sourceId) {
      this.stopDiscovery();
      this.dispatchRuntimeConfig(false, prevProps.sourceId, prevProps.runtimeFilterJson);
      this.restoreTarget();
      this.setState({ status: 'waiting' });
      this.startDiscovery();
    } else if (this._targetHost) {
      if (prevProps.runtimeFilterJson !== this.props.runtimeFilterJson) {
        this.dispatchRuntimeConfig(true);
      }
      if (prevProps.isPageEditMode !== this.props.isPageEditMode || prevProps.sourceListName !== this.props.sourceListName || prevProps.sourceName !== this.props.sourceName) {
        this.syncConfigurationMarker();
      }
    }
  }

  public componentWillUnmount(): void {
    this.stopDiscovery();
    if (typeof window !== 'undefined' && window.removeEventListener) {
      window.removeEventListener('spse:listcontrol-runtime-config-request', this._runtimeConfigRequestHandler);
    }
    this.dispatchRuntimeConfig(false);
    this.restoreTarget();
  }

  private handleRuntimeConfigRequest(event: any): void {
    if (!this._targetHost) { return; }
    var detail = event && event.detail ? event.detail : {};
    var requestedInstanceId = String(detail.instanceId || '').toLowerCase();
    if (requestedInstanceId && requestedInstanceId === this.getInstanceId()) {
      this.dispatchRuntimeConfig(true);
    }
  }

  private startDiscovery(): void {
    this._attemptCount = 0;
    if (this.tryEmbedTarget()) {
      return;
    }

    this._retryHandle = window.setInterval(() => {
      this._attemptCount += 1;
      if (this.tryEmbedTarget()) {
        this.stopDiscovery();
      } else if (this._attemptCount >= 40) {
        this.stopDiscovery();
        this.setState({ status: 'missing' });
      }
    }, 250);
  }

  private stopDiscovery(): void {
    if (this._retryHandle !== undefined) {
      window.clearInterval(this._retryHandle);
      this._retryHandle = undefined;
    }
  }

  private getInstanceId(): string {
    var matches = String(this.props.sourceId || '').match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/ig);
    return matches && matches.length > 0 ? matches[matches.length - 1].toLowerCase() : String(this.props.sourceId || '').toLowerCase();
  }

  private dispatchRuntimeConfig(active: boolean, sourceId?: string, filterJson?: string): void {
    var configuredSourceId = sourceId === undefined ? this.props.sourceId : sourceId;
    var matches = String(configuredSourceId || '').match(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/ig);
    var instanceId = matches && matches.length > 0 ? matches[matches.length - 1].toLowerCase() : String(configuredSourceId || '').toLowerCase();
    if (!instanceId || typeof window === 'undefined') {
      return;
    }

    var detail = {
      instanceId: instanceId,
      owner: this.props.fieldId,
      active: active,
      filterJson: filterJson === undefined ? this.props.runtimeFilterJson : filterJson
    };
    var runtimeEvent: any;
    if (typeof (window as any).CustomEvent === 'function') {
      runtimeEvent = new (window as any).CustomEvent('spse:listcontrol-runtime-config', { detail: detail });
    } else {
      runtimeEvent = document.createEvent('CustomEvent');
      runtimeEvent.initCustomEvent('spse:listcontrol-runtime-config', false, false, detail);
    }
    window.dispatchEvent(runtimeEvent);
  }

  private findTargetHost(): HTMLElement | null {
    var sourceId = String(this.props.sourceId || '').toLowerCase();
    var instanceId = this.getInstanceId();
    var roots = document.querySelectorAll('.lc-root');
    var availableHosts: HTMLElement[] = [];

    for (var i = 0; i < roots.length; i += 1) {
      var root = roots[i] as HTMLElement;
      var host = root.parentElement;
      if (!host) {
        continue;
      }
      var owner = host.getAttribute('data-reportforms-listcontrol-owner');
      if (owner && owner !== this.props.fieldId) {
        continue;
      }
      availableHosts.push(host);

      var ancestor: HTMLElement | null = host;
      while (ancestor && ancestor !== document.body) {
        var candidateId = String(
          ancestor.getAttribute('data-sp-webpart-instance-id')
          || ancestor.getAttribute('data-sp-webpart-id')
          || ancestor.id
          || ''
        ).toLowerCase();
        if (candidateId === instanceId || candidateId === sourceId || candidateId.indexOf(instanceId) >= 0) {
          return host;
        }
        ancestor = ancestor.parentElement;
      }
    }

    return availableHosts.length === 1 ? availableHosts[0] : null;
  }

  private tryEmbedTarget(): boolean {
    if (!this.props.sourceId || !this._mountElement) {
      return false;
    }

    var targetHost = this.findTargetHost();
    if (!targetHost) {
      return false;
    }

    if (!this._targetHost) {
      this._targetHost = targetHost;
      this._originalParent = targetHost.parentNode;
      this._originalNextSibling = targetHost.nextSibling;
      targetHost.setAttribute('data-reportforms-listcontrol-owner', this.props.fieldId);
      this.syncConfigurationMarker();
    }

    if (targetHost.parentNode !== this._mountElement) {
      this._mountElement.appendChild(targetHost);
    }
    if (this.state.status !== 'embedded') {
      this.setState({ status: 'embedded' });
    }
    this.dispatchRuntimeConfig(true);
    return true;
  }

  private restoreTarget(): void {
    if (!this._targetHost) {
      return;
    }

    this.removeConfigurationMarker();
    this._targetHost.removeAttribute('data-reportforms-listcontrol-owner');
    if (this._originalParent) {
      if (this._originalNextSibling && this._originalNextSibling.parentNode === this._originalParent) {
        this._originalParent.insertBefore(this._targetHost, this._originalNextSibling);
      } else {
        this._originalParent.appendChild(this._targetHost);
      }
    }

    this._targetHost = null;
    this._originalParent = null;
    this._originalNextSibling = null;
  }

  private syncConfigurationMarker(): void {
    this.removeConfigurationMarker();
    if (!this.props.isPageEditMode || !this._originalParent) {
      return;
    }
    var marker = document.createElement('div');
    marker.setAttribute('data-reportforms-listcontrol-configure', this.props.fieldId);
    marker.style.boxSizing = 'border-box';
    marker.style.minHeight = '44px';
    marker.style.padding = '10px 12px';
    marker.style.border = '1px dashed #a19f9d';
    marker.style.backgroundColor = '#faf9f8';
    marker.style.color = '#323130';
    marker.style.fontSize = '14px';
    marker.style.display = 'flex';
    marker.style.alignItems = 'center';
    marker.textContent = 'Configure List: ' + (this.props.sourceListName || this.props.sourceName || this.props.sourceId);
    if (this._originalNextSibling && this._originalNextSibling.parentNode === this._originalParent) {
      this._originalParent.insertBefore(marker, this._originalNextSibling);
    } else {
      this._originalParent.appendChild(marker);
    }
    this._configurationMarker = marker;
  }

  private removeConfigurationMarker(): void {
    if (this._configurationMarker && this._configurationMarker.parentNode) {
      this._configurationMarker.parentNode.removeChild(this._configurationMarker);
    }
    this._configurationMarker = null;
  }

  public render(): JSX.Element {
    var sourceName = this.props.sourceName || this.props.sourceId;
    return (
      <div data-reportforms-listcontrol-field={this.props.fieldId}>
        {this.state.status === 'waiting' && <div style={{ padding: '12px', color: '#605e5c' }}>{strings.ListControlFieldWaiting.replace('{0}', sourceName)}</div>}
        {this.state.status === 'missing' && <div style={{ padding: '12px', color: '#a80000' }}>{strings.ListControlFieldMissing.replace('{0}', sourceName)}</div>}
        <div ref={(element) => { this._mountElement = element; }} />
      </div>
    );
  }
}