import * as strings from 'ListControlWebPartStrings';

export interface IPageNavigationTarget {
  instanceId: string;
  title: string;
  element: HTMLElement;
  aliases: string[];
}

function normalizeInstanceId(value: string): string {
  var id = String(value || '').replace(/[{}]/g, '').toLowerCase();
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(id) ? id : '';
}

interface IPageControlIdentity {
  instanceId: string;
  title: string;
  isText: boolean;
}

interface IPageControlData {
  id?: string;
  instanceId?: string;
  title?: string;
  controlType?: number;
  properties?: { instanceName?: string };
  webPartData?: IPageControlData;
}

function label(value: string): string {
  return typeof value === 'string' ? value.replace(/\s+/g, ' ').trim() : '';
}

function readIdentity(element: HTMLElement): IPageControlIdentity {
  var instanceId = '';
  ['data-spse-instance-id', 'data-spse-form-instance-id', 'data-sp-webpart-instance-id',
    'data-sp-feature-instance-id', 'webpartid'].forEach(function(attribute) {
    instanceId = instanceId || normalizeInstanceId(element.getAttribute(attribute) || '');
  });
  var title = label(element.getAttribute('data-spse-instance-name') || '');
  var isText = false;
  ['data-sp-controldata', 'data-sp-webpartdata'].forEach(function(attribute) {
    var serialized = element.getAttribute(attribute);
    if (!serialized) { return; }
    try {
      var data: IPageControlData = JSON.parse(serialized);
      if (!data || typeof data !== 'object' || Array.isArray(data)) {
        throw new Error('Page-control metadata must be an object.');
      }
      var webPart = data.webPartData || data;
      // In web-part metadata, id identifies the component, not its page instance.
      instanceId = instanceId || normalizeInstanceId(webPart.instanceId || '')
        || normalizeInstanceId(data.instanceId || '')
        || (attribute === 'data-sp-controldata' ? normalizeInstanceId(data.id || '') : '');
      title = title || label(webPart.properties && webPart.properties.instanceName)
        || label(webPart.title) || label(data.title);
      isText = isText || data.controlType === 4;
    } catch (error) {
      console.error('[PageNavigation] Unable to read page-control metadata.', error);
    }
  });
  if (!instanceId && /^(div_|WebPart[._])/i.test(element.id)) {
    instanceId = normalizeInstanceId(element.id.replace(/^(div_|WebPart[._])/i, ''));
  }
  return { instanceId: instanceId, title: title, isText: isText };
}

function controlRoot(element: HTMLElement, ownerDocument: Document): HTMLElement {
  var webPartRoot: HTMLElement;
  var identityRoot: HTMLElement;
  var ancestor = element;
  while (ancestor && ancestor !== ownerDocument.body) {
    if (ancestor.hasAttribute('data-sp-canvascontrol') || ancestor.hasAttribute('data-sp-controldata')) { return ancestor; }
    if (!webPartRoot && ancestor.hasAttribute('data-sp-webpart')) { webPartRoot = ancestor; }
    if (!identityRoot && readIdentity(ancestor).instanceId) { identityRoot = ancestor; }
    ancestor = ancestor.parentElement;
  }
  return webPartRoot || identityRoot || element;
}

function displayTitle(element: HTMLElement, isText: boolean): string {
  var title = label(element.getAttribute('aria-label') || '');
  var labelledBy = (element.getAttribute('aria-labelledby') || '').split(/\s+/).map(function(id) {
    var heading = element.ownerDocument.getElementById(id);
    return heading ? label(heading.textContent) : '';
  }).join(' ').trim();
  var heading = element.querySelector('h1,h2,h3,h4,h5,h6,[role="heading"]');
  title = title || labelledBy || (heading ? label(heading.textContent) : '');
  var feature = label(element.getAttribute('data-sp-feature-tag') || '');
  isText = isText || /^(rich )?text(?: web part)?$/i.test(feature);
  return title || (isText ? label(element.textContent).substring(0, 80) || strings.PropNavigationText : feature);
}

export function getPageNavigationTargets(ownerDocument: Document, sourceInstanceId: string): IPageNavigationTarget[] {
  var targets: IPageNavigationTarget[] = [];
  var sourceId = normalizeInstanceId(sourceInstanceId);
  var candidates = ownerDocument.querySelectorAll(
    '[data-spse-instance-id],[data-spse-form-instance-id],[data-spse-tab-type],[data-sp-webpart-instance-id],'
    + '[data-sp-feature-instance-id],[data-sp-controldata],[data-sp-webpartdata],[data-sp-canvascontrol],'
    + '[data-sp-webpart],[webpartid],[id^="div_"],[id^="WebPart."],[id^="WebPart_"],[id^="webpart."],[id^="webpart_"]'
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
  var sourceRoots = ordered.filter(function(element) {
    return !!sourceId && readIdentity(element).instanceId === sourceId;
  }).map(function(element) { return controlRoot(element, ownerDocument); });
  var roots: HTMLElement[] = [];
  ordered.forEach(function(element) {
    var root = controlRoot(element, ownerDocument);
    if (sourceRoots.indexOf(root) >= 0) { return; }
    var rootIndex = roots.indexOf(root);
    if (rootIndex >= 0) {
      var alias = readIdentity(element).instanceId;
      var existing = targets[rootIndex];
      if (alias && alias !== existing.instanceId && existing.aliases.indexOf(alias) < 0) { existing.aliases.push(alias); }
      return;
    }
    var instanceId = '';
    var title = '';
    var isText = false;
    var ancestor: HTMLElement = element;
    while (ancestor && ancestor !== ownerDocument.body) {
      var identity = readIdentity(ancestor);
      instanceId = instanceId || identity.instanceId;
      title = title || identity.title;
      isText = isText || identity.isText;
      if ((instanceId && title) || ancestor === root) { break; }
      ancestor = ancestor.parentElement;
    }
    if (!instanceId || instanceId === sourceId || targets.some(function(target) { return target.instanceId === instanceId; })) {
      return;
    }
    if (!title) {
      var metadataHosts = element.querySelectorAll('[data-sp-webpartdata]');
      for (var hostIndex = 0; hostIndex < metadataHosts.length; hostIndex += 1) {
        var host = metadataHosts[hostIndex] as HTMLElement;
        if (controlRoot(host, ownerDocument) === root) { title = readIdentity(host).title; }
        if (title) { break; }
      }
    }
    title = title || element.getAttribute('data-spse-control-title') || displayTitle(element, isText)
      || (element !== root ? displayTitle(root, isText) : '') || (element.hasAttribute('data-spse-tab-type') ? strings.PropNavigationTabs
      : element.hasAttribute('data-spse-form-instance-id') ? strings.PropNavigationForm : strings.PropNavigationPageControl);
    targets.push({ instanceId: instanceId, title: title, element: element, aliases: [] });
    roots.push(root);
  });
  return targets.sort(function(left, right) { return left.title.localeCompare(right.title); });
}

export function navigateToPageTarget(ownerDocument: Document, targetInstanceId: string, sourceInstanceId: string): void {
  var targets = getPageNavigationTargets(ownerDocument, sourceInstanceId);
  var targetId = normalizeInstanceId(targetInstanceId);
  var target = targets.filter(function(candidate) {
    return candidate.instanceId === targetId || candidate.aliases.indexOf(targetId) >= 0;
  })[0];
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
