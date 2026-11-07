"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_http_1 = require("@microsoft/sp-http");
function normalizeId(value) {
    return String(value || '').replace(/[{}]/g, '').toLowerCase();
}
function readControl(data) {
    var webPart = data && (data.webPartData || data);
    return {
        instanceId: normalizeId(webPart && (webPart.instanceId || data.id)),
        name: String(webPart && webPart.properties && webPart.properties.instanceName || '').trim()
    };
}
function readDocumentNames(ownerDocument) {
    var names = [];
    var controls = ownerDocument.querySelectorAll('[data-sp-controldata]');
    for (var index = 0; index < controls.length; index += 1) {
        names.push(readControl(JSON.parse(controls[index].getAttribute('data-sp-controldata'))));
    }
    return names;
}
function readCanvasNames(canvas, ownerDocument) {
    var source = String(canvas || '').trim();
    if (!source) {
        return [];
    }
    if (source.charAt(0) === '[') {
        var controls = JSON.parse(source);
        if (!Array.isArray(controls)) {
            throw new Error('Invalid page canvas metadata.');
        }
        return controls.map(readControl);
    }
    var snapshot = ownerDocument.implementation.createHTMLDocument('');
    snapshot.body.innerHTML = source;
    return readDocumentNames(snapshot);
}
function needsName(properties, prefix, instanceId) {
    var name = String(properties.instanceName || '').trim();
    return !name || (!!properties.instanceNameGeneratedFor
        && normalizeId(properties.instanceNameGeneratedFor) !== instanceId
        && new RegExp('^' + prefix + '[1-9][0-9]*$', 'i').test(name));
}
function ensureInstanceName(properties, prefix, context, element, getDisplayMode) {
    if (getDisplayMode() !== sp_core_library_1.DisplayMode.Edit) {
        return Promise.resolve(false);
    }
    var ownerDocument = element.ownerDocument;
    var pageContext = context.pageContext;
    var pageKey = pageContext.web.absoluteUrl + '|'
        + (pageContext.list && pageContext.list.id.toString() || '') + '|'
        + (pageContext.listItem && pageContext.listItem.id || ownerDocument.location.pathname);
    var registry = ownerDocument.__spseInstanceNames;
    if (!registry || registry.pageKey !== pageKey) {
        registry = { pageKey: pageKey, liveNames: {} };
        ownerDocument.__spseInstanceNames = registry;
    }
    var instanceId = normalizeId(context.instanceId);
    var registerName = function () {
        registry.liveNames[instanceId] = String(properties.instanceName || '').trim();
        element.setAttribute('data-spse-instance-id', instanceId);
        element.setAttribute('data-spse-instance-name', registry.liveNames[instanceId]);
    };
    if (!needsName(properties, prefix, instanceId)) {
        registerName();
        return Promise.resolve(false);
    }
    if (!registry.pageNames) {
        if (pageContext.list && pageContext.listItem) {
            var listId = normalizeId(pageContext.list.id.toString());
            if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(listId) || !(pageContext.listItem.id > 0)) {
                return Promise.reject(new Error('Cannot allocate a web part name: invalid page identity.'));
            }
            var url = pageContext.web.absoluteUrl + "/_api/web/lists(guid'" + listId + "')/items(" + pageContext.listItem.id + ')';
            registry.pageNames = context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                headers: { Accept: 'application/json;odata=verbose', 'OData-Version': '3.0' }
            }).then(function (response) {
                if (!response.ok) {
                    throw new Error('Cannot read page names: HTTP ' + response.status);
                }
                return response.json().then(function (data) {
                    var page = data && (data.d || data);
                    if (!page || typeof page !== 'object' || page.error) {
                        throw new Error('Invalid page metadata response.');
                    }
                    return readCanvasNames(page.CanvasContent1, ownerDocument);
                });
            }).catch(function (error) {
                registry.pageNames = undefined;
                throw error;
            });
        }
        else {
            registry.pageNames = Promise.resolve([]);
        }
    }
    return registry.pageNames.then(function (savedNames) {
        if (ownerDocument.__spseInstanceNames !== registry || getDisplayMode() !== sp_core_library_1.DisplayMode.Edit) {
            return false;
        }
        if (!needsName(properties, prefix, instanceId)) {
            registerName();
            return false;
        }
        var namesById = {};
        savedNames.concat(readDocumentNames(ownerDocument)).forEach(function (item, index) {
            if (item.name) {
                namesById[item.instanceId || 'unidentified-' + index] = item.name;
            }
        });
        Object.keys(registry.liveNames).forEach(function (id) { namesById[id] = registry.liveNames[id]; });
        var used = {};
        Object.keys(namesById).forEach(function (id) {
            if (id !== instanceId) {
                used[namesById[id].toLowerCase()] = true;
            }
        });
        var number = 1;
        while (used[(prefix + number).toLowerCase()]) {
            number += 1;
        }
        properties.instanceName = prefix + number;
        properties.instanceNameGeneratedFor = instanceId;
        registerName();
        return true;
    });
}
exports.ensureInstanceName = ensureInstanceName;

//# sourceMappingURL=instanceNames.js.map
