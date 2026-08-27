"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var OWNER_ATTRIBUTE = 'data-sps-full-width-owners';
var ORIGINAL_WIDTH_ATTRIBUTE = 'data-sps-full-width-width';
var ORIGINAL_MIN_WIDTH_ATTRIBUTE = 'data-sps-full-width-min-width';
var ORIGINAL_MAX_WIDTH_ATTRIBUTE = 'data-sps-full-width-max-width';
var ORIGINAL_FLEX_BASIS_ATTRIBUTE = 'data-sps-full-width-flex-basis';
var ORIGINAL_FLEX_GROW_ATTRIBUTE = 'data-sps-full-width-flex-grow';
var ORIGINAL_FLEX_SHRINK_ATTRIBUTE = 'data-sps-full-width-flex-shrink';
var ORIGINAL_GRID_COLUMN_ATTRIBUTE = 'data-sps-full-width-grid-column';
var ORIGINAL_BOX_SIZING_ATTRIBUTE = 'data-sps-full-width-box-sizing';
var ORIGINAL_MARGIN_LEFT_ATTRIBUTE = 'data-sps-full-width-margin-left';
var ORIGINAL_MARGIN_RIGHT_ATTRIBUTE = 'data-sps-full-width-margin-right';
var ORIGINAL_OVERFLOW_X_ATTRIBUTE = 'data-sps-full-width-overflow-x';
var ORIGINAL_OVERFLOW_Y_ATTRIBUTE = 'data-sps-full-width-overflow-y';
var ADDED_NATIVE_CLASS_ATTRIBUTE = 'data-sps-full-width-native-class';
function hasLayoutName(element, name) { var automationId = String(element.getAttribute('data-automation-id') || ''); if (automationId === name) {
    return true;
} var classes = String(element.className || '').split(/\s+/); return classes.some(function (className) { return className === name || className.indexOf(name + '-') === 0 || className.indexOf(name + '_') === 0; }); }
function isLayoutContainer(element) { return hasLayoutName(element, 'CanvasZone') || hasLayoutName(element, 'CanvasZoneContainer') || hasLayoutName(element, 'CanvasZoneSectionContainer') || hasLayoutName(element, 'CanvasSection') || hasLayoutName(element, 'ControlZone'); }
function getOwners(element) { return String(element.getAttribute(OWNER_ATTRIBUTE) || '').split(',').filter(function (owner) { return !!owner; }); }
function getFullWidthLeftOffset(documentRef) { var hostname = String(documentRef.location && documentRef.location.hostname || '').toLowerCase(); var isLocal = hostname === 'localhost' || hostname === '127.0.0.1'; var isSharePointOnline = /\.sharepoint(?:-mil)?\.[a-z.]+$/.test(hostname); return !isLocal && !isSharePointOnline ? -25 : -15; }
function getFullWidthLayout(element) { var elementRect = element.getBoundingClientRect(); var documentRef = element.ownerDocument; var viewportWidth = documentRef.documentElement.clientWidth; var pageLeft = 0; var boundary = element; var ancestor = element.parentElement; while (ancestor && ancestor !== documentRef.body) {
    var rect = ancestor.getBoundingClientRect();
    var left = Math.max(0, rect.left + ancestor.clientLeft);
    var right = Math.min(viewportWidth, rect.left + ancestor.clientLeft + ancestor.clientWidth);
    var width = Math.max(0, right - left);
    if (width >= viewportWidth * 0.8 && right >= viewportWidth * 0.9 && left >= pageLeft) {
        pageLeft = left;
        boundary = ancestor;
    }
    ancestor = ancestor.parentElement;
} var measuredLeftOffset = Math.floor(pageLeft - elementRect.left); var rightOffset = Math.floor(elementRect.right - viewportWidth); return { bounds: { width: Math.max(1, Math.floor(elementRect.width - measuredLeftOffset - rightOffset + 10)), leftOffset: getFullWidthLeftOffset(documentRef), rightOffset: -5 }, boundary: boundary }; }
function addOwner(element, ownerId, bounds, overflowOnly) { var owners = getOwners(element); if (owners.indexOf(ownerId) >= 0) {
    return;
} if (owners.length === 0) {
    element.setAttribute(ORIGINAL_WIDTH_ATTRIBUTE, element.style.width || '');
    element.setAttribute(ORIGINAL_MIN_WIDTH_ATTRIBUTE, element.style.minWidth || '');
    element.setAttribute(ORIGINAL_MAX_WIDTH_ATTRIBUTE, element.style.maxWidth || '');
    element.setAttribute(ORIGINAL_FLEX_BASIS_ATTRIBUTE, element.style.flexBasis || '');
    element.setAttribute(ORIGINAL_FLEX_GROW_ATTRIBUTE, element.style.flexGrow || '');
    element.setAttribute(ORIGINAL_FLEX_SHRINK_ATTRIBUTE, element.style.flexShrink || '');
    element.setAttribute(ORIGINAL_GRID_COLUMN_ATTRIBUTE, element.style.getPropertyValue('grid-column') || '');
    element.setAttribute(ORIGINAL_BOX_SIZING_ATTRIBUTE, element.style.boxSizing || '');
    element.setAttribute(ORIGINAL_MARGIN_LEFT_ATTRIBUTE, element.style.marginLeft || '');
    element.setAttribute(ORIGINAL_MARGIN_RIGHT_ATTRIBUTE, element.style.marginRight || '');
    element.setAttribute(ORIGINAL_OVERFLOW_X_ATTRIBUTE, element.style.overflowX || '');
    element.setAttribute(ORIGINAL_OVERFLOW_Y_ATTRIBUTE, element.style.overflowY || '');
    if (!overflowOnly && hasLayoutName(element, 'CanvasZone') && !element.classList.contains('CanvasZone--fullWidth')) {
        element.classList.add('CanvasZone--fullWidth');
        element.setAttribute(ADDED_NATIVE_CLASS_ATTRIBUTE, 'true');
    }
} owners.push(ownerId); element.setAttribute(OWNER_ATTRIBUTE, owners.join(',')); if (overflowOnly) {
    element.style.overflowX = 'visible';
    element.style.overflowY = 'visible';
    return;
} if (bounds) {
    var width = String(bounds.width) + 'px';
    element.style.setProperty('min-width', width, 'important');
    element.style.setProperty('max-width', width, 'important');
    element.style.marginLeft = String(bounds.leftOffset) + 'px';
    element.style.marginRight = String(bounds.rightOffset) + 'px';
    return;
} element.style.width = '100%'; element.style.minWidth = '0'; element.style.maxWidth = 'none'; element.style.flexBasis = '100%'; element.style.flexGrow = '1'; element.style.flexShrink = '0'; element.style.setProperty('grid-column', '1 / -1'); element.style.boxSizing = 'border-box'; }
function removeOwner(element, ownerId) { var owners = getOwners(element).filter(function (owner) { return owner !== ownerId; }); if (owners.length > 0) {
    element.setAttribute(OWNER_ATTRIBUTE, owners.join(','));
    return;
} element.style.setProperty('width', element.getAttribute(ORIGINAL_WIDTH_ATTRIBUTE) || ''); element.style.setProperty('min-width', element.getAttribute(ORIGINAL_MIN_WIDTH_ATTRIBUTE) || ''); element.style.setProperty('max-width', element.getAttribute(ORIGINAL_MAX_WIDTH_ATTRIBUTE) || ''); element.style.setProperty('flex-basis', element.getAttribute(ORIGINAL_FLEX_BASIS_ATTRIBUTE) || ''); element.style.setProperty('flex-grow', element.getAttribute(ORIGINAL_FLEX_GROW_ATTRIBUTE) || ''); element.style.setProperty('flex-shrink', element.getAttribute(ORIGINAL_FLEX_SHRINK_ATTRIBUTE) || ''); element.style.setProperty('grid-column', element.getAttribute(ORIGINAL_GRID_COLUMN_ATTRIBUTE) || ''); element.style.boxSizing = element.getAttribute(ORIGINAL_BOX_SIZING_ATTRIBUTE) || ''; element.style.marginLeft = element.getAttribute(ORIGINAL_MARGIN_LEFT_ATTRIBUTE) || ''; element.style.marginRight = element.getAttribute(ORIGINAL_MARGIN_RIGHT_ATTRIBUTE) || ''; element.style.overflowX = element.getAttribute(ORIGINAL_OVERFLOW_X_ATTRIBUTE) || ''; element.style.overflowY = element.getAttribute(ORIGINAL_OVERFLOW_Y_ATTRIBUTE) || ''; if (element.getAttribute(ADDED_NATIVE_CLASS_ATTRIBUTE) === 'true') {
    element.classList.remove('CanvasZone--fullWidth');
} element.removeAttribute(OWNER_ATTRIBUTE); element.removeAttribute(ORIGINAL_WIDTH_ATTRIBUTE); element.removeAttribute(ORIGINAL_MIN_WIDTH_ATTRIBUTE); element.removeAttribute(ORIGINAL_MAX_WIDTH_ATTRIBUTE); element.removeAttribute(ORIGINAL_FLEX_BASIS_ATTRIBUTE); element.removeAttribute(ORIGINAL_FLEX_GROW_ATTRIBUTE); element.removeAttribute(ORIGINAL_FLEX_SHRINK_ATTRIBUTE); element.removeAttribute(ORIGINAL_GRID_COLUMN_ATTRIBUTE); element.removeAttribute(ORIGINAL_BOX_SIZING_ATTRIBUTE); element.removeAttribute(ORIGINAL_MARGIN_LEFT_ATTRIBUTE); element.removeAttribute(ORIGINAL_MARGIN_RIGHT_ATTRIBUTE); element.removeAttribute(ORIGINAL_OVERFLOW_X_ATTRIBUTE); element.removeAttribute(ORIGINAL_OVERFLOW_Y_ATTRIBUTE); element.removeAttribute(ADDED_NATIVE_CLASS_ATTRIBUTE); }
function adjustOptionalFullWidthForScrollbar(documentRef, ownerId) {
    var ownedElements = documentRef.querySelectorAll('[' + OWNER_ATTRIBUTE + ']');
    var scrollbarGutter = 0;
    for (var index = 0; index < ownedElements.length; index += 1) {
        var element = ownedElements[index];
        if (getOwners(element).indexOf(ownerId) < 0) {
            continue;
        }
        element.style.overflowY = element.getAttribute('data-sps-full-width-overflow-y') || '';
        var computedStyle = documentRef.defaultView ? documentRef.defaultView.getComputedStyle(element) : undefined;
        var borderWidth = computedStyle ? parseFloat(computedStyle.borderLeftWidth || '0') + parseFloat(computedStyle.borderRightWidth || '0') : 0;
        scrollbarGutter = Math.max(scrollbarGutter, Math.max(0, element.offsetWidth - element.clientWidth - borderWidth));
    }
    if (scrollbarGutter <= 0) {
        return;
    }
    for (var index = 0; index < ownedElements.length; index += 1) {
        var element = ownedElements[index];
        if (getOwners(element).indexOf(ownerId) < 0) {
            continue;
        }
        var appliedWidth = parseFloat(element.style.maxWidth || '');
        if (!isNaN(appliedWidth) && appliedWidth > scrollbarGutter) {
            var adjustedWidth = String(Math.floor(appliedWidth - scrollbarGutter)) + 'px';
            element.style.setProperty('min-width', adjustedWidth, 'important');
            element.style.setProperty('max-width', adjustedWidth, 'important');
        }
    }
}
exports.adjustOptionalFullWidthForScrollbar = adjustOptionalFullWidthForScrollbar;
function updateOptionalFullWidth(domElement, ownerId, enabled) { releaseOptionalFullWidth(domElement.ownerDocument, ownerId); if (!enabled) {
    return;
} var path = []; var containers = []; var ancestor = domElement.parentElement; var depth = 0; while (ancestor && ancestor !== domElement.ownerDocument.body && depth < 32) {
    path.push(ancestor);
    if (isLayoutContainer(ancestor)) {
        containers.push(ancestor);
    }
    ancestor = ancestor.parentElement;
    depth += 1;
} var nativeFullWidth = containers.some(function (container) { return hasLayoutName(container, 'CanvasZone') && container.classList.contains('CanvasZone--fullWidth'); }); if (nativeFullWidth) {
    return;
} var breakoutContainer; for (var index = 0; index < containers.length; index += 1) {
    if (hasLayoutName(containers[index], 'CanvasZoneSectionContainer')) {
        breakoutContainer = containers[index];
        break;
    }
} if (!breakoutContainer) {
    for (var index = 0; index < containers.length; index += 1) {
        if (hasLayoutName(containers[index], 'CanvasSection')) {
            breakoutContainer = containers[index];
            break;
        }
    }
} if (!breakoutContainer && containers.length > 0) {
    breakoutContainer = containers[containers.length - 1];
} if (!breakoutContainer) {
    return;
} var layout = getFullWidthLayout(breakoutContainer); var breakoutIndex = path.indexOf(breakoutContainer); var boundaryIndex = path.indexOf(layout.boundary); path.slice(0, breakoutIndex + 1).forEach(function (container) { return addOwner(container, ownerId, container === breakoutContainer ? layout.bounds : undefined); }); if (boundaryIndex > breakoutIndex) {
    path.slice(breakoutIndex + 1, boundaryIndex).forEach(function (container) { return addOwner(container, ownerId, undefined, true); });
} }
exports.updateOptionalFullWidth = updateOptionalFullWidth;
var RESPONSIVE_MONITORS = {};
function updateResponsiveOptionalFullWidth(domElement, ownerId, enabled) { updateOptionalFullWidth(domElement, ownerId, enabled); if (!enabled) {
    return;
} var documentRef = domElement.ownerDocument; var viewportWidth = documentRef.documentElement.clientWidth; var hostname = String(documentRef.location && documentRef.location.hostname || '').toLowerCase(); var isLocal = hostname === 'localhost' || hostname === '127.0.0.1'; var isSharePointOnline = /\.sharepoint(?:-mil)?\.[a-z.]+$/.test(hostname); var isSharePointServer = !isLocal && !isSharePointOnline; var serverRightGutter = isSharePointServer ? 20 : 0; var propertyPaneLeft = 0; var availableRight = viewportWidth; var ancestor = domElement.parentElement; while (ancestor && ancestor !== documentRef.body) {
    var rect = ancestor.getBoundingClientRect();
    var right = Math.min(viewportWidth, rect.left + ancestor.clientLeft + ancestor.clientWidth);
    if (rect.width >= viewportWidth * 0.45 && right > viewportWidth * 0.45 && right < availableRight) {
        availableRight = right;
    }
    ancestor = ancestor.parentElement;
} var pane = documentRef.querySelector('#spPropertyPaneContainer, .spPropertyPaneContainer, [data-automation-id="propertyPane"]'); if (pane) {
    var paneRect = pane.getBoundingClientRect();
    if (paneRect.width > 0 && paneRect.left > 0 && paneRect.left < viewportWidth) {
        availableRight = Math.min(availableRight, paneRect.left);
        propertyPaneLeft = paneRect.left;
    }
} var overlap = Math.max(0, viewportWidth - availableRight) + serverRightGutter; if (overlap > 0) {
    var ownedElements = documentRef.querySelectorAll('[' + OWNER_ATTRIBUTE + ']');
    for (var index = 0; index < ownedElements.length; index += 1) {
        var element = ownedElements[index];
        var appliedWidth = parseFloat(element.style.maxWidth || '');
        if (getOwners(element).indexOf(ownerId) >= 0 && !isNaN(appliedWidth) && appliedWidth > overlap) {
            var adjustedWidth = isSharePointServer && propertyPaneLeft > 0 ? Math.floor(propertyPaneLeft - element.getBoundingClientRect().left - serverRightGutter) : Math.floor(appliedWidth - overlap);
            var width = String(Math.max(1, adjustedWidth)) + 'px';
            element.style.setProperty('min-width', width, 'important');
            element.style.setProperty('max-width', width, 'important');
        }
    }
} adjustOptionalFullWidthForScrollbar(documentRef, ownerId); var MutationObserverConstructor = documentRef.defaultView.MutationObserver; if (MutationObserverConstructor && documentRef.body) {
    var monitor_1 = { observer: undefined, timer: 0 };
    monitor_1.observer = new MutationObserverConstructor(function () { if (monitor_1.timer) {
        window.clearTimeout(monitor_1.timer);
    } monitor_1.timer = window.setTimeout(function () { return updateResponsiveOptionalFullWidth(domElement, ownerId, true); }, 350); });
    monitor_1.observer.observe(documentRef.body, { childList: true, subtree: true });
    RESPONSIVE_MONITORS[ownerId] = monitor_1;
} }
exports.updateResponsiveOptionalFullWidth = updateResponsiveOptionalFullWidth;
function releaseOptionalFullWidth(documentRef, ownerId) { var monitor = RESPONSIVE_MONITORS[ownerId]; if (monitor) {
    if (monitor.timer) {
        window.clearTimeout(monitor.timer);
    }
    if (monitor.observer) {
        monitor.observer.disconnect();
    }
    delete RESPONSIVE_MONITORS[ownerId];
} var ownedElements = documentRef.querySelectorAll('[' + OWNER_ATTRIBUTE + ']'); for (var index = 0; index < ownedElements.length; index += 1) {
    removeOwner(ownedElements[index], ownerId);
} }
exports.releaseOptionalFullWidth = releaseOptionalFullWidth;

//# sourceMappingURL=optionalFullWidth.js.map
