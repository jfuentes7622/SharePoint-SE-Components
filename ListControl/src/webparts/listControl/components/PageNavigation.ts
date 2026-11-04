import * as strings from 'ListControlWebPartStrings';

export interface IPageNavigationTarget {
  instanceId: string;
  title: string;
  element: HTMLElement;
}

function normalizeInstanceId(value: string): string {
  var id = String(value || '').replace(/[{}]/g, '').toLowerCase();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(id) ? id : '';
}

function readIdentity(element: HTMLElement): { instanceId: string; title: string } {
  var instanceId = normalizeInstanceId(element.getAttribute('data-spse-instance-id') || element.getAttribute('data-spse-form-instance-id')
    || element.getAttribute('data-sp-webpart-instance-id') || element.getAttribute('data-sp-feature-instance-id') || '');
  var title = element.getAttribute('data-spse-instance-name') || '';
  var serialized = element.getAttribute('data-sp-controldata');
  if (serialized) {
    try {
      var data = JSON.parse(serialized);
      var webPart = data.webPartData || {};
      if (data.controlType === undefined || data.controlType === 3 || data.webPartData) {
        instanceId = instanceId || normalizeInstanceId(webPart.instanceId || data.id || '');
        title = title || String((webPart.properties && webPart.properties.instanceName) || webPart.title || data.title || '');
      }
    } catch (error) {
      console.error('[PageNavigation] Unable to read page-control metadata.', error);
    }
  }
  if (!instanceId && /^(div_|WebPart[._])/i.test(element.id)) {
    instanceId = normalizeInstanceId(element.id.replace(/^(div_|WebPart[._])/i, ''));
  }
  return { instanceId: instanceId, title: title };
}

export function getPageNavigationTargets(ownerDocument: Document, sourceInstanceId: string): IPageNavigationTarget[] {
  var targets: IPageNavigationTarget[] = [];
  var sourceId = normalizeInstanceId(sourceInstanceId);
  var candidates = ownerDocument.querySelectorAll(
    '[data-spse-instance-id],[data-spse-form-instance-id],[data-spse-tab-type],[data-sp-webpart-instance-id],'
    + '[data-sp-feature-instance-id],[data-sp-controldata],[id^="div_"],[id^="WebPart."]'
  );
  // Prefer the control itself over its SharePoint canvas wrapper.
  var ordered: HTMLElement[] = [];
  for (var index = 0; index < candidates.length; index += 1) {
    var candidate = candidates[index] as HTMLElement;
    if (candidate.hasAttribute('data-spse-instance-id') || candidate.hasAttribute('data-spse-form-instance-id') || candidate.hasAttribute('data-spse-tab-type')) {
      ordered.unshift(candidate);
    } else {
      ordered.push(candidate);
    }
  }
  ordered.forEach(function(element) {
    var instanceId = '';
    var title = '';
    var ancestor: HTMLElement = element;
    while (ancestor && ancestor !== ownerDocument.body) {
      var identity = readIdentity(ancestor);
      instanceId = instanceId || identity.instanceId;
      title = title || identity.title;
      if (instanceId && title) { break; }
      ancestor = ancestor.parentElement;
    }
    if (!instanceId || instanceId === sourceId || targets.some(function(target) { return target.instanceId === instanceId; })) {
      return;
    }
    title = title || (element.hasAttribute('data-spse-tab-type') ? strings.PropNavigationTabs
      : element.hasAttribute('data-spse-form-instance-id') ? strings.PropNavigationForm : strings.PropNavigationPageControl);
    targets.push({ instanceId: instanceId, title: title, element: element });
  });
  return targets.sort(function(left, right) { return left.title.localeCompare(right.title); });
}

export function navigateToPageTarget(ownerDocument: Document, targetInstanceId: string, sourceInstanceId: string): void {
  var targets = getPageNavigationTargets(ownerDocument, sourceInstanceId);
  var targetId = normalizeInstanceId(targetInstanceId);
  var target = targets.filter(function(candidate) { return candidate.instanceId === targetId; })[0];
  if (!target) { throw new Error(strings.RuntimeNavigationMissing); }
  var ancestor: HTMLElement = target.element;
  while (ancestor) {
    var style = ownerDocument.defaultView.getComputedStyle(ancestor);
    if (ancestor.hidden || style.display === 'none' || style.visibility === 'hidden' || style.visibility === 'collapse') {
      throw new Error(strings.RuntimeNavigationHidden);
    }
    ancestor = ancestor.parentElement;
  }
  var landing = (target.element.querySelector('[role="tablist"]') as HTMLElement) || target.element;
  if (!landing.scrollIntoView) { throw new Error(strings.RuntimeNavigationFailed); }
  var addedTabIndex = !landing.hasAttribute('tabindex');
  if (addedTabIndex) { landing.setAttribute('tabindex', '-1'); }
  landing.focus();
  if (ownerDocument.activeElement !== landing) {
    if (addedTabIndex) { landing.removeAttribute('tabindex'); }
    throw new Error(strings.RuntimeNavigationFailed);
  }
  if (addedTabIndex) {
    var restoreTabIndex = function(): void {
      landing.removeAttribute('tabindex');
      landing.removeEventListener('blur', restoreTabIndex);
    };
    landing.addEventListener('blur', restoreTabIndex);
  }
  landing.scrollIntoView({ block: 'center', behavior: 'auto' });
}
