import { DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';
import {
  releaseSectionFullWidth,
  updateResponsiveSectionFullWidth
} from '../shared/sectionFullWidth';

interface IInlineStyleValue {
  value: string;
  priority: string;
}

interface IHostStyleSnapshot {
  element: HTMLElement;
  values: { [propertyName: string]: IInlineStyleValue };
}

export interface IFullWidthControlWebPartProps {
}

export default class FullWidthControlWebPart extends BaseClientSideWebPart<IFullWidthControlWebPartProps> {
  private hostSnapshot: IHostStyleSnapshot | undefined;

  public render(): void {
    updateResponsiveSectionFullWidth(this.domElement, this.context.instanceId);
    this.updateOwnHostVisibility(this.displayMode === DisplayMode.Edit);

    if (this.displayMode === DisplayMode.Edit) {
      this.renderEditMarker();
    } else {
      this.domElement.innerHTML = '';
    }
  }

  protected onDispose(): void {
    this.domElement.innerHTML = '';
    releaseSectionFullWidth(this.domElement.ownerDocument, this.context.instanceId);
    this.restoreOwnHost();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }

  private renderEditMarker(): void {
    this.domElement.innerHTML = '';
    const marker = document.createElement('div');
    marker.setAttribute('role', 'status');
    marker.textContent = 'Full Width Control: this section spans the available page width';
    marker.style.boxSizing = 'border-box';
    marker.style.width = '100%';
    marker.style.minHeight = '36px';
    marker.style.padding = '8px 12px';
    marker.style.border = '1px solid #0078d4';
    marker.style.borderRadius = '2px';
    marker.style.backgroundColor = '#f3f9fd';
    marker.style.color = '#323130';
    marker.style.fontFamily = 'Segoe UI, sans-serif';
    marker.style.fontSize = '12px';
    marker.style.fontWeight = '600';
    marker.style.lineHeight = '18px';
    this.domElement.appendChild(marker);
  }

  private updateOwnHostVisibility(isEditMode: boolean): void {
    const host = this.findOwnHost();
    if (!host) {
      return;
    }

    if (!this.hostSnapshot || this.hostSnapshot.element !== host) {
      this.restoreOwnHost();
      const propertyNames = [
        'height', 'min-height', 'margin-top', 'margin-bottom',
        'padding-top', 'padding-bottom', 'border-top-width',
        'border-bottom-width', 'overflow', 'position'
      ];
      const values: { [propertyName: string]: IInlineStyleValue } = {};
      propertyNames.forEach((propertyName: string) => {
        values[propertyName] = {
          value: host.style.getPropertyValue(propertyName),
          priority: host.style.getPropertyPriority(propertyName)
        };
      });
      this.hostSnapshot = { element: host, values: values };
    }

    host.style.setProperty('height', isEditMode ? 'auto' : '0', 'important');
    host.style.setProperty('min-height', isEditMode ? '40px' : '0', 'important');
    host.style.setProperty('margin-top', '0', 'important');
    host.style.setProperty('margin-bottom', '0', 'important');
    host.style.setProperty('padding-top', isEditMode ? '2px' : '0', 'important');
    host.style.setProperty('padding-bottom', isEditMode ? '2px' : '0', 'important');
    host.style.setProperty('border-top-width', '0', 'important');
    host.style.setProperty('border-bottom-width', '0', 'important');
    host.style.setProperty('overflow', isEditMode ? 'hidden' : 'visible', 'important');
    if (!host.style.position || host.style.position === 'static') {
      host.style.setProperty('position', 'relative');
    }

    this.domElement.style.setProperty('position', 'relative');
    this.domElement.style.setProperty('height', isEditMode ? 'auto' : '0', 'important');
    this.domElement.style.setProperty('min-height', isEditMode ? '36px' : '0', 'important');
    this.domElement.style.setProperty('margin', '0', 'important');
    this.domElement.style.setProperty('padding', '0', 'important');
    this.domElement.style.setProperty('overflow', isEditMode ? 'hidden' : 'visible', 'important');
  }

  private findOwnHost(): HTMLElement | undefined {
    let ancestor = this.domElement.parentElement;
    let controlZone: HTMLElement | undefined;
    let depth = 0;
    while (ancestor && ancestor !== this.domElement.ownerDocument.body && depth < 16) {
      if (String(ancestor.getAttribute('data-automation-id') || '') === 'CanvasControl') {
        return ancestor;
      }
      if (!controlZone && this.hasClassPrefix(ancestor, 'ControlZone')) {
        controlZone = ancestor;
      }
      ancestor = ancestor.parentElement;
      depth += 1;
    }
    return controlZone;
  }

  private hasClassPrefix(element: HTMLElement, prefix: string): boolean {
    const classNames = String(element.className || '').split(/\s+/);
    return classNames.some((className: string) => {
      return className === prefix || className.indexOf(prefix + '-') === 0 || className.indexOf(prefix + '_') === 0;
    });
  }

  private restoreOwnHost(): void {
    if (!this.hostSnapshot) {
      return;
    }
    const snapshot = this.hostSnapshot;
    Object.keys(snapshot.values).forEach((propertyName: string) => {
      const saved = snapshot.values[propertyName];
      if (saved.value) {
        snapshot.element.style.setProperty(propertyName, saved.value, saved.priority);
      } else {
        snapshot.element.style.removeProperty(propertyName);
      }
    });
    this.hostSnapshot = undefined;
  }
}
