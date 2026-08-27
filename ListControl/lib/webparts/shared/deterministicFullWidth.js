"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var OWNER_ATTRIBUTE = 'data-sps-full-width-owners';
var BREAKOUT_ATTRIBUTE = 'data-sps-full-width-breakout';
var ADDED_NATIVE_CLASS_ATTRIBUTE = 'data-sps-full-width-native-class';
var ORIGINAL_PREFIX = 'data-sps-full-width-';
var RESPONSIVE_MONITORS = {};
function hasLayoutName(element, name) {
    var automationId = String(element.getAttribute('data-automation-id') || '');
    if (automationId === name) {
        return true;
    }
    var classes = String(element.className || '').split(/\s+/);
    return classes.some(function (className) { return className === name || className.indexOf(name + '-') === 0 || className.indexOf(name + '_') === 0; });
}
function getOwners(element) {
    return String(element.getAttribute(OWNER_ATTRIBUTE) || '').split(',').filter(function (owner) { return !!owner; });
}
function saveStyle(element, propertyName) {
    element.setAttribute(ORIGINAL_PREFIX + propertyName, element.style.getPropertyValue(propertyName) || '');
    element.setAttribute(ORIGINAL_PREFIX + propertyName + '-priority', element.style.getPropertyPriority(propertyName) || '');
}
function restoreStyle(element, propertyName) {
    var value = element.getAttribute(ORIGINAL_PREFIX + propertyName) || '';
    var priority = element.getAttribute(ORIGINAL_PREFIX + propertyName + '-priority') || '';
    element.style.setProperty(propertyName, value, priority);
    element.removeAttribute(ORIGINAL_PREFIX + propertyName);
    element.removeAttribute(ORIGINAL_PREFIX + propertyName + '-priority');
}
function addOwner(element, ownerId, isBreakout) {
    var owners = getOwners(element);
    if (owners.length === 0) {
        ['width', 'min-width', 'max-width', 'flex-basis', 'flex-grow', 'flex-shrink', 'grid-column', 'box-sizing', 'margin-left', 'margin-right', 'overflow-x', 'overflow-y'].forEach(function (propertyName) { return saveStyle(element, propertyName); });
        if (hasLayoutName(element, 'CanvasZone') && !element.classList.contains('CanvasZone--fullWidth')) {
            element.classList.add('CanvasZone--fullWidth');
            element.setAttribute(ADDED_NATIVE_CLASS_ATTRIBUTE, 'true');
        }
    }
    if (owners.indexOf(ownerId) < 0) {
        owners.push(ownerId);
        element.setAttribute(OWNER_ATTRIBUTE, owners.join(','));
    }
    element.style.setProperty('width', '100%', 'important');
    element.style.setProperty('min-width', '0', 'important');
    element.style.setProperty('max-width', 'none', 'important');
    element.style.setProperty('flex-basis', '100%', 'important');
    element.style.setProperty('flex-grow', '1', 'important');
    element.style.setProperty('flex-shrink', '1', 'important');
    element.style.setProperty('grid-column', '1 / -1', 'important');
    element.style.setProperty('box-sizing', 'border-box', 'important');
    if (isBreakout) {
        element.setAttribute(BREAKOUT_ATTRIBUTE, ownerId);
        element.style.setProperty('margin-left', '10px', 'important');
        element.style.setProperty('margin-right', '0', 'important');
    }
}
function removeOwner(element, ownerId) {
    var owners = getOwners(element).filter(function (owner) { return owner !== ownerId; });
    if (owners.length > 0) {
        element.setAttribute(OWNER_ATTRIBUTE, owners.join(','));
        if (element.getAttribute(BREAKOUT_ATTRIBUTE) === ownerId) {
            element.removeAttribute(BREAKOUT_ATTRIBUTE);
        }
        return;
    }
    ['width', 'min-width', 'max-width', 'flex-basis', 'flex-grow', 'flex-shrink', 'grid-column', 'box-sizing', 'margin-left', 'margin-right', 'overflow-x', 'overflow-y'].forEach(function (propertyName) { return restoreStyle(element, propertyName); });
    if (element.getAttribute(ADDED_NATIVE_CLASS_ATTRIBUTE) === 'true') {
        element.classList.remove('CanvasZone--fullWidth');
    }
    element.removeAttribute(OWNER_ATTRIBUTE);
    element.removeAttribute(BREAKOUT_ATTRIBUTE);
    element.removeAttribute(ADDED_NATIVE_CLASS_ATTRIBUTE);
}
function getVisiblePropertyPaneLeft(documentRef, viewportWidth) {
    var pane = documentRef.querySelector('#spPropertyPaneContainer, .spPropertyPaneContainer, [data-automation-id="propertyPane"]');
    if (!pane) {
        return 0;
    }
    var paneRect = pane.getBoundingClientRect();
    return paneRect.width > 0 && paneRect.left > 0 && paneRect.left < viewportWidth ? paneRect.left : 0;
}
function getBreakoutContainer(path) {
    var preferredNames = ['CanvasZone', 'CanvasZoneContainer', 'CanvasZoneSectionContainer', 'CanvasSection'];
    for (var nameIndex = 0; nameIndex < preferredNames.length; nameIndex += 1) {
        for (var pathIndex = path.length - 1; pathIndex >= 0; pathIndex -= 1) {
            if (hasLayoutName(path[pathIndex], preferredNames[nameIndex])) {
                return path[pathIndex];
            }
        }
    }
    return undefined;
}
function applyAvailableWidth(domElement, ownerId) {
    var documentRef = domElement.ownerDocument;
    var viewportWidth = documentRef.documentElement.clientWidth;
    var paneLeft = getVisiblePropertyPaneLeft(documentRef, viewportWidth);
    var hostname = String(documentRef.location && documentRef.location.hostname || '').toLowerCase();
    var isLocal = hostname === 'localhost' || hostname === '127.0.0.1';
    var isSharePointOnline = /\.sharepoint(?:-mil)?\.[a-z.]+$/.test(hostname);
    var rightGutter = !isLocal && !isSharePointOnline ? 20 : 0;
    var availableRight = paneLeft > 0 ? paneLeft : viewportWidth;
    var ownedElements = documentRef.querySelectorAll('[' + OWNER_ATTRIBUTE + ']');
    for (var index = 0; index < ownedElements.length; index += 1) {
        var element = ownedElements[index];
        if (getOwners(element).indexOf(ownerId) < 0 || element.getAttribute(BREAKOUT_ATTRIBUTE) !== ownerId) {
            continue;
        }
        var targetWidth = Math.max(1, Math.floor(availableRight - element.getBoundingClientRect().left - rightGutter));
        var width = String(targetWidth) + 'px';
        element.style.setProperty('width', width, 'important');
        element.style.setProperty('min-width', width, 'important');
        element.style.setProperty('max-width', width, 'important');
    }
}
function scheduleUpdate(domElement, ownerId) {
    var monitor = RESPONSIVE_MONITORS[ownerId];
    if (!monitor) {
        return;
    }
    if (monitor.timer) {
        window.clearTimeout(monitor.timer);
    }
    monitor.timer = window.setTimeout(function () { return updateResponsiveOptionalFullWidth(domElement, ownerId, true); }, 350);
}
function updateResponsiveOptionalFullWidth(domElement, ownerId, enabled) {
    releaseOptionalFullWidth(domElement.ownerDocument, ownerId);
    if (!enabled) {
        return;
    }
    var path = [];
    var ancestor = domElement.parentElement;
    var depth = 0;
    while (ancestor && ancestor !== domElement.ownerDocument.body && depth < 32) {
        path.push(ancestor);
        ancestor = ancestor.parentElement;
        depth += 1;
    }
    var breakoutContainer = getBreakoutContainer(path);
    if (!breakoutContainer) {
        return;
    }
    var breakoutIndex = path.indexOf(breakoutContainer);
    path.slice(0, breakoutIndex + 1).forEach(function (element) { return addOwner(element, ownerId, element === breakoutContainer); });
    applyAvailableWidth(domElement, ownerId);
    var documentRef = domElement.ownerDocument;
    var MutationObserverConstructor = documentRef.defaultView && documentRef.defaultView.MutationObserver;
    var monitor = {
        observer: undefined,
        timer: 0,
        resizeHandler: function () { return scheduleUpdate(domElement, ownerId); }
    };
    if (MutationObserverConstructor && documentRef.body) {
        monitor.observer = new MutationObserverConstructor(function () { return scheduleUpdate(domElement, ownerId); });
        monitor.observer.observe(documentRef.body, { childList: true, subtree: true });
        var propertyPane = documentRef.querySelector('#spPropertyPaneContainer, .spPropertyPaneContainer, [data-automation-id="propertyPane"]');
        if (propertyPane) {
            monitor.observer.observe(propertyPane, { attributes: true, attributeFilter: ['class', 'style', 'aria-hidden'] });
        }
    }
    if (documentRef.defaultView) {
        documentRef.defaultView.addEventListener('resize', monitor.resizeHandler);
    }
    RESPONSIVE_MONITORS[ownerId] = monitor;
}
exports.updateResponsiveOptionalFullWidth = updateResponsiveOptionalFullWidth;
function releaseOptionalFullWidth(documentRef, ownerId) {
    var monitor = RESPONSIVE_MONITORS[ownerId];
    if (monitor) {
        if (monitor.timer) {
            window.clearTimeout(monitor.timer);
        }
        if (monitor.observer) {
            monitor.observer.disconnect();
        }
        if (documentRef.defaultView) {
            documentRef.defaultView.removeEventListener('resize', monitor.resizeHandler);
        }
        delete RESPONSIVE_MONITORS[ownerId];
    }
    var ownedElements = documentRef.querySelectorAll('[' + OWNER_ATTRIBUTE + ']');
    for (var index = 0; index < ownedElements.length; index += 1) {
        removeOwner(ownedElements[index], ownerId);
    }
}
exports.releaseOptionalFullWidth = releaseOptionalFullWidth;

//# sourceMappingURL=deterministicFullWidth.js.map
