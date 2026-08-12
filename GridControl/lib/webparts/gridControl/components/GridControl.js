"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var sp_http_1 = require("@microsoft/sp-http");
var strings = require("GridControlWebPartStrings");
var GridValidation_1 = require("./GridValidation");
require("./GridControl.css");
function isCompatibleGridControlType(sharePointType, controlType) {
    var compatibleTypes = {
        Text: ['text', 'multiline', 'number'],
        Note: ['text', 'multiline', 'number'],
        Number: ['number'],
        Currency: ['number'],
        Integer: ['number'],
        Choice: ['dropdown'],
        MultiChoice: ['multiselect'],
        DateTime: ['datetime'],
        Boolean: ['boolean'],
        URL: ['url', 'text'],
        Hyperlink: ['url', 'text'],
        Lookup: ['lookup'],
        LookupMulti: ['lookup'],
        User: ['person'],
        UserMulti: ['person'],
        TaxonomyFieldType: ['taxonomy'],
        TaxonomyFieldTypeMulti: ['taxonomy'],
        Image: ['image'],
        Thumbnail: ['image'],
        Attachments: ['attachment']
    };
    var allowed = compatibleTypes[sharePointType];
    return !!allowed && allowed.indexOf(String(controlType || '').toLowerCase()) >= 0;
}
var GRID_CONTROL_REFRESH_EVENT = 'spse:gridcontrol-refresh';
function escapeODataText(value) {
    return value.replace(/'/g, "''");
}
function toPositiveInt(value) {
    var parsed = parseInt(String(value || ''), 10);
    return !isNaN(parsed) && parsed > 0 ? parsed : 0;
}
function formatString(template, value) {
    return template.replace('{0}', String(value));
}
function tryParseObject(value) {
    if (!value || typeof value !== 'string') {
        return value;
    }
    try {
        return JSON.parse(value);
    }
    catch (_error) {
        return value;
    }
}
function toArray(value) {
    var parsed = tryParseObject(value);
    if (!parsed) {
        return [];
    }
    if (Array.isArray(parsed)) {
        return parsed;
    }
    if (Array.isArray(parsed.results)) {
        return parsed.results;
    }
    return [];
}
function trimGuidBraces(value) {
    return String(value || '').replace(/^[{]/, '').replace(/[}]$/, '');
}
function appendQuery(url, query) {
    return url + (url.indexOf('?') >= 0 ? '&' : '?') + query;
}
function appendQueryParam(url, key, value) {
    var hash = '';
    var base = url;
    var hashIndex = url.indexOf('#');
    if (hashIndex >= 0) {
        base = url.substring(0, hashIndex);
        hash = url.substring(hashIndex);
    }
    var separator = base.indexOf('?') >= 0 ? '&' : '?';
    return base + separator + encodeURIComponent(key) + '=' + encodeURIComponent(value) + hash;
}
function isWorkbenchUrl(url) {
    return /\/_layouts\/15\/workbench\.aspx/i.test(String(url || ''));
}
function getQueryParam(url, paramName) {
    var queryIndex = url.indexOf('?');
    if (queryIndex < 0) {
        return '';
    }
    var hashIndex = url.indexOf('#', queryIndex);
    var query = hashIndex >= 0 ? url.substring(queryIndex + 1, hashIndex) : url.substring(queryIndex + 1);
    var parts = query.split('&');
    for (var i = 0; i < parts.length; i += 1) {
        var pair = parts[i].split('=');
        if (pair.length < 1) {
            continue;
        }
        var key = decodeURIComponent(String(pair[0] || '')).toLowerCase();
        if (key === paramName.toLowerCase()) {
            return decodeURIComponent(String(pair[1] || ''));
        }
    }
    return '';
}
function stripQueryAndHash(url) {
    var result = String(url || '');
    var hashIndex = result.indexOf('#');
    if (hashIndex >= 0) {
        result = result.substring(0, hashIndex);
    }
    var queryIndex = result.indexOf('?');
    if (queryIndex >= 0) {
        result = result.substring(0, queryIndex);
    }
    return result;
}
function resolveEmbeddedPageUrl(url) {
    var source = String(url || '').trim();
    if (!source) {
        return '';
    }
    // First choice: explicit Source/source parameter.
    var fromSource = getQueryParam(source, 'source') || getQueryParam(source, 'Source');
    if (fromSource && /\.aspx/i.test(fromSource)) {
        return fromSource;
    }
    // Second choice: id parameter used by SharePoint wrapper pages.
    var fromId = getQueryParam(source, 'id') || getQueryParam(source, 'Id') || getQueryParam(source, 'ID');
    if (fromId && /\.aspx/i.test(fromId)) {
        return fromId;
    }
    return '';
}
function isSharePointWrapperUrl(url) {
    return /\/_layouts\/15\/(sharepoint|onedrive|doc)\.aspx/i.test(String(url || ''));
}
function buildViewRequestUrls(baseEndpoint, selectedViewId) {
    var urls = [];
    var normalized = trimGuidBraces(selectedViewId);
    var candidates = [String(selectedViewId || ''), normalized, '{' + normalized + '}'];
    for (var i = 0; i < candidates.length; i += 1) {
        var candidate = String(candidates[i] || '');
        if (!candidate) {
            continue;
        }
        urls.push(appendQuery(baseEndpoint, 'View=' + encodeURIComponent(candidate)));
        urls.push(appendQuery(baseEndpoint, 'ViewId=' + encodeURIComponent(candidate)));
    }
    urls.push(baseEndpoint);
    var unique = [];
    for (var j = 0; j < urls.length; j += 1) {
        if (unique.indexOf(urls[j]) < 0) {
            unique.push(urls[j]);
        }
    }
    return unique;
}
function extractRenderRowsAndFields(data) {
    var schema = tryParseObject(data.ListSchema) || tryParseObject(data.Schema) || {};
    var listData = tryParseObject(data.ListData) || {};
    var rows = toArray(data.Row);
    if (rows.length === 0) {
        rows = toArray(data.Rows);
    }
    if (rows.length === 0) {
        rows = toArray(listData.Row);
    }
    if (rows.length === 0) {
        rows = toArray(listData.Rows);
    }
    if (rows.length === 0) {
        rows = toArray(data && data.d && data.d.Row);
    }
    var fields = toArray(schema.Field);
    if (fields.length === 0) {
        fields = toArray(data.Field);
    }
    return {
        rows: rows,
        fields: fields
    };
}
function isSystemItemKey(key) {
    if (!key) {
        return true;
    }
    return key.indexOf('__') === 0 || key === 'odata.editLink' || key === 'odata.id' || key === 'odata.type' || key === 'GUID';
}
function normalizeViewFieldNames(value) {
    var items = toArray(value);
    var names = [];
    for (var i = 0; i < items.length; i += 1) {
        var candidate = String(items[i] || '');
        if (!candidate) {
            continue;
        }
        if (names.indexOf(candidate) < 0) {
            names.push(candidate);
        }
    }
    return names;
}
function isExpandableFieldType(typeAsString) {
    var normalized = String(typeAsString || '').trim();
    return normalized === 'User' || normalized === 'UserMulti' || normalized === 'Lookup' || normalized === 'LookupMulti';
}
function joinClassNames(classes) {
    return classes.filter(function (className) {
        return !!className;
    }).join(' ');
}
function normalizeFilterOperator(value) {
    var normalized = String(value || '').trim().toLowerCase().replace(/\s+/g, '');
    if (normalized === 'eq' || normalized === 'ne' || normalized === 'contains' || normalized === 'notcontains' || normalized === 'startswith' || normalized === 'endswith' || normalized === 'gt' || normalized === 'ge' || normalized === 'lt' || normalized === 'le') {
        return normalized;
    }
    return 'contains';
}
function normalizeFilterLogical(value) {
    return String(value || '').trim().toLowerCase() === 'or' ? 'or' : 'and';
}
function formatLocalDate(value) {
    var month = String(value.getMonth() + 1);
    var day = String(value.getDate());
    return String(value.getFullYear()) + '-' + (month.length < 2 ? '0' + month : month) + '-' + (day.length < 2 ? '0' + day : day);
}
function toCssString(value) {
    return String(value === undefined || value === null ? '' : value).trim();
}
function mergeStyleObjects(baseStyle, overrideStyle) {
    var merged = {};
    var key;
    for (key in baseStyle) {
        if (Object.prototype.hasOwnProperty.call(baseStyle, key)) {
            merged[key] = baseStyle[key];
        }
    }
    for (key in overrideStyle) {
        if (Object.prototype.hasOwnProperty.call(overrideStyle, key)) {
            merged[key] = overrideStyle[key];
        }
    }
    return merged;
}
function toPriorityNumber(value, fallback) {
    var parsed = parseInt(String(value === undefined || value === null ? '' : value), 10);
    if (isNaN(parsed)) {
        return fallback;
    }
    return parsed;
}
var GridControl = (function (_super) {
    __extends(GridControl, _super);
    function GridControl(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            selectedViewId: props.defaultViewId || _this.getInitialViewId(props.views),
            fields: [],
            fieldMetadataByName: {},
            rows: [],
            loading: true,
            error: null,
            selectedItemId: 0,
            selectedItemIds: [],
            selectedMode: 'view',
            deleting: false,
            deleteMessage: '',
            sortFieldName: '',
            sortDirection: '',
            activeFilterFieldName: '',
            draftFilterOperator: 'contains',
            draftFilterValue: '',
            columnFilters: {},
            currentPage: 0,
            editingItemId: -1,
            editingValues: {},
            editingErrors: {},
            saving: false,
        };
        _this._refreshEventHandler = _this.handleExternalRefresh.bind(_this);
        return _this;
    }
    GridControl.prototype.componentDidMount = function () {
        this.logDiagnostic('Component mounted. listName=' + String(this.props.listName || '(none)') + ', defaultViewId=' + String(this.props.defaultViewId || '(none)'));
        if (typeof window !== 'undefined' && window.addEventListener) {
            window.addEventListener(GRID_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
        }
        this.loadRows();
    };
    GridControl.prototype.componentWillUnmount = function () {
        if (typeof window !== 'undefined' && window.removeEventListener) {
            window.removeEventListener(GRID_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
        }
    };
    GridControl.prototype.componentDidUpdate = function (prevProps, prevState) {
        var _this = this;
        if (prevProps.listName !== this.props.listName || prevProps.defaultViewId !== this.props.defaultViewId
            || prevProps.gridSchemaJson !== this.props.gridSchemaJson) {
            this.logDiagnostic('Props changed; resetting selection and reloading rows. listName=' + String(this.props.listName || '(none)') + ', viewId=' + String(this.props.defaultViewId || '(none)'));
            this.setState({
                selectedViewId: this.props.defaultViewId || this.getInitialViewId(this.props.views),
                selectedItemId: 0,
                selectedItemIds: [],
                selectedMode: 'view',
                deleteMessage: '',
                sortFieldName: '',
                sortDirection: '',
                activeFilterFieldName: '',
                draftFilterOperator: 'contains',
                draftFilterValue: '',
                columnFilters: {},
                currentPage: 0,
                editingItemId: -1,
                editingValues: {},
                editingErrors: {},
                saving: false,
            }, function () {
                _this.props.onSelectionChange(0, 'view');
                _this.loadRows();
            });
            return;
        }
        if (prevState.selectedViewId !== this.state.selectedViewId) {
            if (this.state.currentPage !== 0) {
                this.setState({ currentPage: 0 });
            }
            this.loadRows();
            return;
        }
        if (prevProps.pageSize !== this.props.pageSize || prevProps.filterJson !== this.props.filterJson) {
            this.setState({ currentPage: 0 });
        }
    };
    GridControl.prototype.getInitialViewId = function (views) {
        for (var i = 0; i < views.length; i += 1) {
            if (views[i].isDefault) {
                return String(views[i].key);
            }
        }
        return views.length > 0 ? String(views[0].key) : '';
    };
    GridControl.prototype.handleExternalRefresh = function (event) {
        var detail = event && event.detail ? event.detail : {};
        var sourceListName = String(detail.listName || '');
        if (!sourceListName || !this.props.listName) {
            return;
        }
        if (sourceListName.toLowerCase() !== String(this.props.listName).toLowerCase()) {
            return;
        }
        this.logDiagnostic('Received external refresh event for list: ' + sourceListName);
        this.loadRows();
    };
    GridControl.prototype.getWebUrl = function () {
        return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    };
    GridControl.prototype.getJsonWithFallback = function (url) {
        return __awaiter(this, void 0, void 0, function () {
            var response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.props.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.props.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=verbose'
                                }
                            })];
                    case 2:
                        response = _a.sent();
                        _a.label = 3;
                    case 3:
                        if (!!response.ok) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.props.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=minimalmetadata'
                                }
                            })];
                    case 4:
                        response = _a.sent();
                        _a.label = 5;
                    case 5:
                        if (!!response.ok) return [3 /*break*/, 7];
                        return [4 /*yield*/, this.props.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata'
                                }
                            })];
                    case 6:
                        response = _a.sent();
                        _a.label = 7;
                    case 7: return [2 /*return*/, response];
                }
            });
        });
    };
    GridControl.prototype.postJsonWithFallback = function (url, body) {
        return __awaiter(this, void 0, void 0, function () {
            var payload, response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        payload = JSON.stringify(body || {});
                        return [4 /*yield*/, this.props.context.spHttpClient.post(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    'Content-Type': 'application/json; charset=utf-8'
                                },
                                body: payload
                            })];
                    case 1:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.props.context.spHttpClient.post(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=verbose',
                                    'Content-Type': 'application/json;odata=verbose'
                                },
                                body: payload
                            })];
                    case 2:
                        response = _a.sent();
                        _a.label = 3;
                    case 3:
                        if (!!response.ok) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.props.context.spHttpClient.post(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=minimalmetadata',
                                    'Content-Type': 'application/json;odata=minimalmetadata'
                                },
                                body: payload
                            })];
                    case 4:
                        response = _a.sent();
                        _a.label = 5;
                    case 5:
                        if (!!response.ok) return [3 /*break*/, 7];
                        return [4 /*yield*/, this.props.context.spHttpClient.post(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata'
                                },
                                body: payload
                            })];
                    case 6:
                        response = _a.sent();
                        _a.label = 7;
                    case 7: return [2 /*return*/, response];
                }
            });
        });
    };
    GridControl.prototype.buildViewIdCandidates = function (selectedViewId) {
        var normalized = trimGuidBraces(selectedViewId);
        var candidates = [String(selectedViewId || ''), normalized, '{' + normalized + '}'];
        var unique = [];
        for (var i = 0; i < candidates.length; i += 1) {
            var value = String(candidates[i] || '');
            if (value && unique.indexOf(value) < 0) {
                unique.push(value);
            }
        }
        return unique;
    };
    GridControl.prototype.loadSelectedViewFieldNames = function (selectedViewId) {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, listPath, viewIds, urls, i, encoded, normalized, j, response, data, names, viewFieldsError_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!selectedViewId) {
                            return [2 /*return*/, []];
                        }
                        webUrl = this.getWebUrl();
                        listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
                        viewIds = this.buildViewIdCandidates(selectedViewId);
                        urls = [];
                        for (i = 0; i < viewIds.length; i += 1) {
                            encoded = encodeURIComponent(viewIds[i]);
                            normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
                            urls.push(webUrl + listPath + "/views/getById('" + encoded + "')/ViewFields");
                            urls.push(webUrl + listPath + "/views(guid'" + normalized + "')/ViewFields");
                        }
                        j = 0;
                        _a.label = 1;
                    case 1:
                        if (!(j < urls.length)) return [3 /*break*/, 7];
                        _a.label = 2;
                    case 2:
                        _a.trys.push([2, 5, , 6]);
                        return [4 /*yield*/, this.getJsonWithFallback(urls[j])];
                    case 3:
                        response = _a.sent();
                        if (!response.ok) {
                            return [3 /*break*/, 6];
                        }
                        return [4 /*yield*/, response.json()];
                    case 4:
                        data = _a.sent();
                        names = normalizeViewFieldNames(data.value);
                        if (names.length === 0) {
                            names = normalizeViewFieldNames(data.Items);
                        }
                        if (names.length === 0) {
                            names = normalizeViewFieldNames(data && data.d && data.d.Items);
                        }
                        if (names.length > 0) {
                            return [2 /*return*/, names];
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        viewFieldsError_1 = _a.sent();
                        this.logDiagnostic('loadSelectedViewFieldNames: Attempt failed for url=' + urls[j] + ': ' + (viewFieldsError_1 && viewFieldsError_1.message ? viewFieldsError_1.message : String(viewFieldsError_1)));
                        return [3 /*break*/, 6];
                    case 6:
                        j += 1;
                        return [3 /*break*/, 1];
                    case 7:
                        this.logDiagnostic('loadSelectedViewFieldNames: No view field names resolved for viewId=' + String(selectedViewId));
                        return [2 /*return*/, []];
                }
            });
        });
    };
    GridControl.prototype.getDisplayFields = function () {
        var baseFields = this.state.fields;
        if (this.state.selectedViewId === this.props.defaultViewId && this.props.viewColumns && this.props.viewColumns.length > 0) {
            var configuredByName = {};
            for (var i = 0; i < this.state.fields.length; i += 1) {
                var field = this.state.fields[i];
                var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
                var responseName = String(field.Name || '').toLowerCase();
                if (fieldName) {
                    configuredByName[fieldName] = field;
                }
                if (responseName) {
                    configuredByName[responseName] = field;
                }
            }
            var configured = [];
            for (var j = 0; j < this.props.viewColumns.length; j += 1) {
                var column = this.props.viewColumns[j];
                var configuredMatch = configuredByName[String(column.fieldName || '').toLowerCase()];
                if (configuredMatch) {
                    configured.push(Object.assign({}, configuredMatch, {
                        DisplayName: column.displayName || configuredMatch.DisplayName || configuredMatch.Name,
                        ConfiguredWidth: column.width || ''
                    }));
                }
            }
            baseFields = configured;
        }
        var schemaFields = this.getGridSchemaFields();
        if (schemaFields.length === 0) {
            return baseFields;
        }
        var byName = {};
        for (var baseIndex = 0; baseIndex < this.state.fields.length; baseIndex += 1) {
            var baseField = this.state.fields[baseIndex];
            byName[String(baseField.Name || '').toLowerCase()] = baseField;
            byName[String(baseField.RealFieldName || '').toLowerCase()] = baseField;
        }
        var schemaDisplayFields = [];
        for (var schemaIndex = 0; schemaIndex < schemaFields.length; schemaIndex += 1) {
            var schemaField = schemaFields[schemaIndex];
            if (schemaField.visible === false) {
                continue;
            }
            var match = byName[String(schemaField.fieldName || '').toLowerCase()];
            if (match) {
                schemaDisplayFields.push(Object.assign({}, match, {
                    DisplayName: schemaField.label || match.DisplayName || match.Name,
                    ConfiguredWidth: schemaField.gridWidth || ''
                }));
            }
        }
        return schemaDisplayFields;
    };
    GridControl.prototype.getGridSchemaFields = function () {
        var schema = this.getGridSchema();
        try {
            var steps = schema && Array.isArray(schema.steps) ? schema.steps : [];
            var fields = [];
            for (var stepIndex = 0; stepIndex < steps.length; stepIndex += 1) {
                var stepFields = steps[stepIndex] && Array.isArray(steps[stepIndex].fields) ? steps[stepIndex].fields : [];
                for (var fieldIndex = 0; fieldIndex < stepFields.length; fieldIndex += 1) {
                    var field = stepFields[fieldIndex];
                    if (field && field.fieldName && field.type !== 'newline' && field.type !== 'richtext') {
                        fields.push(field);
                    }
                }
            }
            return fields;
        }
        catch (_error) {
            return [];
        }
    };
    GridControl.prototype.getGridSchema = function () {
        var json = String(this.props.gridSchemaJson || '').trim();
        if (!json) {
            return {};
        }
        try {
            return JSON.parse(json);
        }
        catch (_error) {
            return {};
        }
    };
    GridControl.prototype.getAdvancedValidationRules = function () {
        var schema = this.getGridSchema();
        var advancedValidation = schema && schema.advancedValidation;
        return advancedValidation && advancedValidation.enabled === true && Array.isArray(advancedValidation.rules)
            ? advancedValidation.rules : [];
    };
    GridControl.prototype.getGridSchemaField = function (fieldName) {
        var normalized = String(fieldName || '').toLowerCase();
        var fields = this.getGridSchemaFields();
        for (var i = 0; i < fields.length; i += 1) {
            if (String(fields[i].fieldName || '').toLowerCase() === normalized) {
                return fields[i];
            }
        }
        return undefined;
    };
    GridControl.prototype.getConfiguredColumnStyle = function (field) {
        var configuredWidth = String(field.ConfiguredWidth || '').trim();
        if (!configuredWidth) {
            return {};
        }
        var width = /^\d+(?:\.\d+)?$/.test(configuredWidth) ? configuredWidth + 'px' : configuredWidth;
        if (!/^\d+(?:\.\d+)?(?:px|%|rem|em|vw)$/.test(width)) {
            return {};
        }
        return { width: width, minWidth: width, maxWidth: width };
    };
    GridControl.prototype.getFieldsForConsumption = function (rawFields, viewFieldNames) {
        if (viewFieldNames.length === 0) {
            return rawFields.filter(function (field) {
                return !(field.Hidden === true || field.Hidden === 'TRUE');
            });
        }
        var byName = {};
        for (var i = 0; i < rawFields.length; i += 1) {
            var field = rawFields[i];
            var key = String(field.Name || field.RealFieldName || '');
            var real = String(field.RealFieldName || '');
            if (key) {
                byName[key] = field;
            }
            if (real) {
                byName[real] = field;
            }
        }
        var mapped = [];
        for (var j = 0; j < viewFieldNames.length; j += 1) {
            var viewFieldName = String(viewFieldNames[j] || '');
            if (!viewFieldName) {
                continue;
            }
            var match = byName[viewFieldName];
            if (match) {
                mapped.push(match);
            }
            else if (!isSystemItemKey(viewFieldName)) {
                mapped.push({ Name: viewFieldName, DisplayName: viewFieldName });
            }
        }
        return mapped;
    };
    GridControl.prototype.loadListFieldTypeMap = function (viewFieldNames) {
        return __awaiter(this, void 0, void 0, function () {
            var endpoint, response, data, fields, requested, i, map, j, field, internalName, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!viewFieldNames || viewFieldNames.length === 0) {
                            return [2 /*return*/, {}];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,TypeAsString";
                        return [4 /*yield*/, this.getJsonWithFallback(endpoint)];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, {}];
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        fields = toArray(data.value);
                        if (fields.length === 0) {
                            fields = toArray(data && data.d && data.d.results);
                        }
                        requested = {};
                        for (i = 0; i < viewFieldNames.length; i += 1) {
                            requested[String(viewFieldNames[i] || '')] = true;
                        }
                        map = {};
                        for (j = 0; j < fields.length; j += 1) {
                            field = fields[j] || {};
                            internalName = String(field.InternalName || '');
                            if (!internalName || !requested[internalName]) {
                                continue;
                            }
                            map[internalName] = String(field.TypeAsString || '');
                        }
                        return [2 /*return*/, map];
                    case 4:
                        error_1 = _a.sent();
                        this.logDiagnostic('loadListFieldTypeMap failed: ' + (error_1 && error_1.message ? error_1.message : String(error_1)));
                        return [2 /*return*/, {}];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    GridControl.prototype.loadListFieldTitleMap = function () {
        return __awaiter(this, void 0, void 0, function () {
            var endpoint, response, data, fields, map, i, field, internalName, title, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,Title";
                        return [4 /*yield*/, this.getJsonWithFallback(endpoint)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, {}];
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        fields = toArray(data.value);
                        if (fields.length === 0) {
                            fields = toArray(data && data.d && data.d.results);
                        }
                        map = {};
                        for (i = 0; i < fields.length; i += 1) {
                            field = fields[i] || {};
                            internalName = String(field.InternalName || '');
                            title = String(field.Title || '');
                            if (internalName && title) {
                                map[internalName.toLowerCase()] = title;
                            }
                        }
                        return [2 /*return*/, map];
                    case 3:
                        error_2 = _a.sent();
                        this.logDiagnostic('loadListFieldTitleMap failed: ' + (error_2 && error_2.message ? error_2.message : String(error_2)));
                        return [2 /*return*/, {}];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    GridControl.prototype.loadGridFieldMetadata = function () {
        return __awaiter(this, void 0, void 0, function () {
            var endpoint, response, data, fields, metadataByName, i, source, internalName, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName)
                            + "')/fields?$select=InternalName,Title,TypeAsString,Required,ReadOnlyField,Hidden,Description,Choices,DisplayFormat";
                        return [4 /*yield*/, this.getJsonWithFallback(endpoint)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, {}];
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        fields = toArray(data.value);
                        if (fields.length === 0) {
                            fields = toArray(data && data.d && data.d.results);
                        }
                        metadataByName = {};
                        for (i = 0; i < fields.length; i += 1) {
                            source = fields[i] || {};
                            internalName = String(source.InternalName || '');
                            if (!internalName) {
                                continue;
                            }
                            metadataByName[internalName.toLowerCase()] = {
                                internalName: internalName,
                                title: String(source.Title || internalName),
                                typeAsString: String(source.TypeAsString || 'Text'),
                                required: source.Required === true,
                                readOnly: source.ReadOnlyField === true,
                                hidden: source.Hidden === true,
                                description: String(source.Description || ''),
                                choices: toArray(source.Choices).map(function (choice) { return String(choice); }),
                                displayFormat: parseInt(String(source.DisplayFormat || '0'), 10) || 0
                            };
                        }
                        return [2 /*return*/, metadataByName];
                    case 3:
                        error_3 = _a.sent();
                        this.logDiagnostic('loadGridFieldMetadata failed: ' + (error_3 && error_3.message ? error_3.message : String(error_3)));
                        return [2 /*return*/, {}];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    GridControl.prototype.applyFieldDisplayNames = function (fields, titleMap) {
        return fields.map(function (field) {
            var internalName = String(field.RealFieldName || field.Name || '');
            var displayName = titleMap[internalName.toLowerCase()] || field.DisplayName || internalName;
            return __assign({}, field, { DisplayName: displayName });
        });
    };
    GridControl.prototype.getRowFieldValue = function (row, field) {
        var fieldName = field.Name || field.RealFieldName || '';
        var value = row[fieldName];
        if ((value === undefined || value === null || value === '') && field.RealFieldName) {
            value = row[field.RealFieldName];
        }
        return value;
    };
    GridControl.prototype.stringifyCellValue = function (value) {
        if (value === undefined || value === null || value === '') {
            return '';
        }
        if (Array.isArray(value)) {
            var listValues = [];
            for (var i = 0; i < value.length; i += 1) {
                var itemText = this.stringifyCellValue(value[i]);
                if (itemText) {
                    listValues.push(itemText);
                }
            }
            return listValues.join('; ');
        }
        if (typeof value === 'object') {
            if (Array.isArray(value.results)) {
                return this.stringifyCellValue(value.results);
            }
            if (value.Title) {
                return String(value.Title);
            }
            if (value.LookupValue) {
                return String(value.LookupValue);
            }
            if (value.Name) {
                return String(value.Name);
            }
            if (value.Email) {
                return String(value.Email);
            }
            if (value.Url) {
                return String(value.Url);
            }
            if (value.Id !== undefined && value.Id !== null) {
                return String(value.Id);
            }
            try {
                return JSON.stringify(value);
            }
            catch (_jsonError) {
                return String(value);
            }
        }
        return String(value);
    };
    GridControl.prototype.isMeaningfulCellValue = function (value) {
        if (value === undefined || value === null) {
            return false;
        }
        var text = this.stringifyCellValue(value).replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, '').trim();
        return text.length > 0;
    };
    GridControl.prototype.filterRenderableRows = function (rows, visibleFields) {
        var filtered = [];
        for (var i = 0; i < rows.length; i += 1) {
            var row = rows[i] || {};
            var hasVisibleValue = false;
            for (var j = 0; j < visibleFields.length; j += 1) {
                var field = visibleFields[j];
                var fieldName = String(field.Name || field.RealFieldName || '');
                var realFieldName = String(field.RealFieldName || '');
                var candidateValue = row[fieldName];
                if ((candidateValue === undefined || candidateValue === null || candidateValue === '') && realFieldName) {
                    candidateValue = row[realFieldName];
                }
                if (this.isMeaningfulCellValue(candidateValue)) {
                    hasVisibleValue = true;
                    break;
                }
            }
            // Keep rows only when at least one visible field has meaningful content.
            if (hasVisibleValue) {
                filtered.push(row);
            }
        }
        return filtered;
    };
    GridControl.prototype.loadRowsFromItemsEndpoint = function (viewFieldNames) {
        return __awaiter(this, void 0, void 0, function () {
            var fieldTypeMap, selectFields, expandFields, i, fieldName, fieldType, endpoint, response, data, rows, fields, v, viewFieldName, firstRow, key;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.loadListFieldTypeMap(viewFieldNames)];
                    case 1:
                        fieldTypeMap = _a.sent();
                        selectFields = ['ID'];
                        expandFields = [];
                        for (i = 0; i < viewFieldNames.length; i += 1) {
                            fieldName = String(viewFieldNames[i] || '');
                            if (!fieldName) {
                                continue;
                            }
                            fieldType = String(fieldTypeMap[fieldName] || '');
                            if (isExpandableFieldType(fieldType)) {
                                if (selectFields.indexOf(fieldName + '/Id') < 0) {
                                    selectFields.push(fieldName + '/Id');
                                }
                                if (selectFields.indexOf(fieldName + '/Title') < 0) {
                                    selectFields.push(fieldName + '/Title');
                                }
                                if (expandFields.indexOf(fieldName) < 0) {
                                    expandFields.push(fieldName);
                                }
                            }
                            else if (selectFields.indexOf(fieldName) < 0) {
                                selectFields.push(fieldName);
                            }
                        }
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items?$top=200";
                        if (selectFields.length > 0) {
                            endpoint += '&$select=' + encodeURIComponent(selectFields.join(','));
                        }
                        if (expandFields.length > 0) {
                            endpoint += '&$expand=' + encodeURIComponent(expandFields.join(','));
                        }
                        return [4 /*yield*/, this.getJsonWithFallback(endpoint)];
                    case 2:
                        response = _a.sent();
                        this.logDiagnostic('Items endpoint fallback executed. Url=' + endpoint + ', ok=' + String(response.ok));
                        if (!response.ok) {
                            return [2 /*return*/, { rows: [], fields: [] }];
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        rows = toArray(data.value);
                        if (rows.length === 0) {
                            rows = toArray(data && data.d && data.d.results);
                        }
                        fields = [];
                        if (viewFieldNames.length > 0) {
                            for (v = 0; v < viewFieldNames.length; v += 1) {
                                viewFieldName = String(viewFieldNames[v] || '');
                                if (viewFieldName && !isSystemItemKey(viewFieldName)) {
                                    fields.push({ Name: viewFieldName, DisplayName: viewFieldName });
                                }
                            }
                        }
                        else if (rows.length > 0) {
                            firstRow = rows[0] || {};
                            for (key in firstRow) {
                                if (firstRow.hasOwnProperty(key) && !isSystemItemKey(key)) {
                                    fields.push({ Name: key, DisplayName: key });
                                }
                            }
                        }
                        return [2 /*return*/, {
                                rows: rows,
                                fields: fields
                            }];
                }
            });
        });
    };
    GridControl.prototype.loadRows = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var baseEndpoint, selectedViewId, body, requestUrls, selectedRows, selectedFields, lastError, hadSuccessfulResponse, requestIndex, requestUrl, response, errorText, _readError_1, data, extracted, rows, fields, viewFieldNames, schemaFields, schemaFieldNames, schemaItems, itemsFallback, visibleFields, fieldTitleMap, fieldMetadataByName, renderableRows, renderableItemIds, selectedItemIds, error_4, loadError;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.props.listName) {
                            this.setState({ loading: false, error: null, fields: [], rows: [] });
                            return [2 /*return*/];
                        }
                        this.logDiagnostic('Starting loadRows. listName=' + String(this.props.listName) + ', selectedViewId=' + String(this.state.selectedViewId || '(none)'));
                        this.setState({ loading: true, error: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 19, , 20]);
                        baseEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/RenderListDataAsStream";
                        selectedViewId = this.state.selectedViewId;
                        body = {
                            parameters: {
                                RenderOptions: 17
                            }
                        };
                        requestUrls = selectedViewId ? buildViewRequestUrls(baseEndpoint, selectedViewId) : [baseEndpoint];
                        selectedRows = [];
                        selectedFields = [];
                        lastError = '';
                        hadSuccessfulResponse = false;
                        requestIndex = 0;
                        _a.label = 2;
                    case 2:
                        if (!(requestIndex < requestUrls.length)) return [3 /*break*/, 11];
                        requestUrl = requestUrls[requestIndex];
                        return [4 /*yield*/, this.postJsonWithFallback(requestUrl, body)];
                    case 3:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 8];
                        errorText = '';
                        _a.label = 4;
                    case 4:
                        _a.trys.push([4, 6, , 7]);
                        return [4 /*yield*/, response.text()];
                    case 5:
                        errorText = _a.sent();
                        return [3 /*break*/, 7];
                    case 6:
                        _readError_1 = _a.sent();
                        errorText = '';
                        return [3 /*break*/, 7];
                    case 7:
                        lastError = 'Failed to load list view data. HTTP ' + String(response.status) + ' ' + response.statusText + (errorText ? (': ' + errorText) : '');
                        return [3 /*break*/, 10];
                    case 8:
                        hadSuccessfulResponse = true;
                        return [4 /*yield*/, response.json()];
                    case 9:
                        data = _a.sent();
                        extracted = extractRenderRowsAndFields(data);
                        if (selectedFields.length === 0 && extracted.fields.length > 0) {
                            selectedFields = extracted.fields;
                        }
                        if (extracted.rows.length > 0 || requestIndex === requestUrls.length - 1) {
                            selectedRows = extracted.rows;
                            if (selectedFields.length === 0) {
                                selectedFields = extracted.fields;
                            }
                            return [3 /*break*/, 11];
                        }
                        _a.label = 10;
                    case 10:
                        requestIndex += 1;
                        return [3 /*break*/, 2];
                    case 11:
                        if (!hadSuccessfulResponse) {
                            throw new Error(lastError || 'Failed to load list view data.');
                        }
                        rows = selectedRows;
                        fields = selectedFields;
                        return [4 /*yield*/, this.loadSelectedViewFieldNames(selectedViewId)];
                    case 12:
                        viewFieldNames = _a.sent();
                        schemaFields = this.getGridSchemaFields();
                        if (!(schemaFields.length > 0)) return [3 /*break*/, 14];
                        schemaFieldNames = schemaFields.filter(function (field) {
                            return field.visible !== false && !!field.fieldName;
                        }).map(function (field) {
                            return String(field.fieldName);
                        });
                        return [4 /*yield*/, this.loadRowsFromItemsEndpoint(schemaFieldNames)];
                    case 13:
                        schemaItems = _a.sent();
                        rows = schemaItems.rows;
                        fields = schemaItems.fields;
                        viewFieldNames = schemaFieldNames;
                        return [3 /*break*/, 16];
                    case 14:
                        if (!(rows.length === 0)) return [3 /*break*/, 16];
                        return [4 /*yield*/, this.loadRowsFromItemsEndpoint(viewFieldNames)];
                    case 15:
                        itemsFallback = _a.sent();
                        rows = itemsFallback.rows;
                        if (fields.length === 0) {
                            fields = itemsFallback.fields;
                        }
                        _a.label = 16;
                    case 16:
                        visibleFields = this.getFieldsForConsumption(fields, viewFieldNames);
                        return [4 /*yield*/, this.loadListFieldTitleMap()];
                    case 17:
                        fieldTitleMap = _a.sent();
                        return [4 /*yield*/, this.loadGridFieldMetadata()];
                    case 18:
                        fieldMetadataByName = _a.sent();
                        visibleFields = this.applyFieldDisplayNames(visibleFields, fieldTitleMap);
                        renderableRows = this.filterRenderableRows(rows, visibleFields);
                        renderableItemIds = renderableRows.map(function (row) { return _this.getRowItemId(row); });
                        selectedItemIds = this.state.selectedItemIds.filter(function (itemId) {
                            return renderableItemIds.indexOf(itemId) >= 0;
                        });
                        this.setState({
                            fields: visibleFields,
                            fieldMetadataByName: fieldMetadataByName,
                            rows: renderableRows,
                            selectedItemIds: selectedItemIds,
                            loading: false,
                            error: null
                        });
                        this.logDiagnostic('loadRows completed. visibleFields=' + String(visibleFields.length) + ', renderableRows=' + String(renderableRows.length));
                        return [3 /*break*/, 20];
                    case 19:
                        error_4 = _a.sent();
                        loadError = error_4;
                        this.setState({
                            loading: false,
                            error: loadError && loadError.message ? loadError.message : 'Failed to load data.',
                            fields: [],
                            fieldMetadataByName: {},
                            rows: []
                        });
                        this.logDiagnostic('loadRows failed: ' + (loadError && loadError.message ? loadError.message : String(loadError)));
                        return [3 /*break*/, 20];
                    case 20: return [2 /*return*/];
                }
            });
        });
    };
    GridControl.prototype.getRowItemId = function (row) {
        return toPositiveInt(row.ID || row.Id || row.id);
    };
    GridControl.prototype.selectRow = function (row) {
        var itemId = this.getRowItemId(row);
        this.setState({ selectedItemId: itemId, selectedMode: 'view' });
        this.logDiagnostic('Row selected. itemId=' + String(itemId));
        this.props.onSelectionChange(itemId, 'view');
    };
    GridControl.prototype.isItemChecked = function (itemId) {
        return this.state.selectedItemIds.indexOf(itemId) >= 0;
    };
    GridControl.prototype.toggleItemChecked = function (itemId) {
        if (itemId <= 0 || this.state.deleting) {
            return;
        }
        var selectedItemIds = this.state.selectedItemIds.slice(0);
        var selectedIndex = selectedItemIds.indexOf(itemId);
        if (selectedIndex >= 0) {
            selectedItemIds.splice(selectedIndex, 1);
        }
        else {
            selectedItemIds.push(itemId);
        }
        this.setState({ selectedItemIds: selectedItemIds });
    };
    GridControl.prototype.toggleVisibleItemsChecked = function (rows) {
        var _this = this;
        var visibleItemIds = rows.map(function (row) { return _this.getRowItemId(row); }).filter(function (itemId) {
            return itemId > 0;
        });
        var allVisibleSelected = visibleItemIds.length > 0 && visibleItemIds.every(function (itemId) { return _this.isItemChecked(itemId); });
        var selectedItemIds = this.state.selectedItemIds.filter(function (itemId) {
            return allVisibleSelected ? visibleItemIds.indexOf(itemId) < 0 : true;
        });
        if (!allVisibleSelected) {
            for (var itemIndex = 0; itemIndex < visibleItemIds.length; itemIndex += 1) {
                if (selectedItemIds.indexOf(visibleItemIds[itemIndex]) < 0) {
                    selectedItemIds.push(visibleItemIds[itemIndex]);
                }
            }
        }
        this.setState({ selectedItemIds: selectedItemIds });
    };
    GridControl.prototype.getGridFieldMetadata = function (field) {
        var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
        var metadata = this.state.fieldMetadataByName[fieldName];
        if (!metadata) {
            metadata = this.state.fieldMetadataByName[String(field.Name || '').toLowerCase()];
        }
        if (!metadata) {
            return undefined;
        }
        var schemaField = this.getGridSchemaField(metadata.internalName);
        if (!schemaField) {
            return metadata;
        }
        var typeMap = {
            text: 'Text',
            multiline: 'Note',
            number: 'Number',
            boolean: 'Boolean',
            dropdown: 'Choice',
            multiselect: 'MultiChoice',
            datetime: 'DateTime',
            url: 'URL'
        };
        var config = schemaField.config || {};
        var displayFormat = config.displayFormat === 'dateOnly' ? 0 : config.displayFormat === 'timeOnly' ? 2 : 1;
        var schemaControlType = String(schemaField.type || '').toLowerCase();
        var effectiveType = isCompatibleGridControlType(metadata.typeAsString, schemaControlType)
            ? typeMap[schemaControlType] : metadata.typeAsString;
        return Object.assign({}, metadata, {
            title: schemaField.label || metadata.title,
            typeAsString: effectiveType,
            required: metadata.required || schemaField.required === true,
            readOnly: schemaField.readOnly === true || schemaField.disabled === true || metadata.readOnly,
            description: metadata.description || schemaField.description || config.helpText,
            choices: Array.isArray(config.choices) ? config.choices.map(function (choice) { return String(choice); }) : metadata.choices,
            displayFormat: schemaField.type === 'datetime' ? displayFormat : metadata.displayFormat
        });
    };
    GridControl.prototype.isEditableGridField = function (field) {
        var metadata = this.getGridFieldMetadata(field);
        if (!metadata || metadata.readOnly || metadata.hidden) {
            return false;
        }
        var supportedTypes = ['Text', 'Note', 'Number', 'Currency', 'Integer', 'Boolean', 'Choice', 'MultiChoice', 'DateTime', 'URL'];
        return supportedTypes.indexOf(metadata.typeAsString) >= 0;
    };
    GridControl.prototype.normalizeEditingValue = function (value, metadata) {
        if (metadata.typeAsString === 'Boolean') {
            return value === true || value === 1 || String(value).toLowerCase() === 'yes' || String(value).toLowerCase() === 'true';
        }
        if (metadata.typeAsString === 'MultiChoice') {
            return toArray(value).map(function (entry) { return String(entry); });
        }
        if (metadata.typeAsString === 'DateTime') {
            if (!value) {
                return '';
            }
            var parsedDate = new Date(String(value));
            if (isNaN(parsedDate.getTime())) {
                return '';
            }
            var localDate = new Date(parsedDate.getTime() - (parsedDate.getTimezoneOffset() * 60000));
            var isoValue = localDate.toISOString();
            return metadata.displayFormat === 0 ? isoValue.substring(0, 10)
                : metadata.displayFormat === 2 ? isoValue.substring(11, 16) : isoValue.substring(0, 16);
        }
        if (metadata.typeAsString === 'URL' && value && typeof value === 'object') {
            return String(value.Url || '');
        }
        return value === undefined || value === null ? '' : this.stringifyCellValue(value);
    };
    GridControl.prototype.beginRowEdit = function (row) {
        var itemId = this.getRowItemId(row);
        if (itemId <= 0 || this.state.saving) {
            return;
        }
        var values = {};
        var fields = this.getDisplayFields();
        for (var i = 0; i < fields.length; i += 1) {
            var metadata = this.getGridFieldMetadata(fields[i]);
            if (metadata && this.isEditableGridField(fields[i])) {
                values[metadata.internalName] = this.normalizeEditingValue(this.getRowFieldValue(row, fields[i]), metadata);
            }
        }
        this.setState({
            selectedItemId: itemId,
            selectedMode: 'edit',
            editingItemId: itemId,
            editingValues: values,
            editingErrors: {},
            error: null
        });
        this.props.onSelectionChange(itemId, 'edit');
    };
    GridControl.prototype.beginNewRow = function () {
        if (this.state.saving) {
            return;
        }
        var values = {};
        var fields = this.getDisplayFields();
        for (var i = 0; i < fields.length; i += 1) {
            var metadata = this.getGridFieldMetadata(fields[i]);
            if (metadata && this.isEditableGridField(fields[i])) {
                var schemaField = this.getGridSchemaField(metadata.internalName);
                values[metadata.internalName] = schemaField && schemaField.defaultValue !== undefined
                    ? schemaField.defaultValue
                    : metadata.typeAsString === 'Boolean' ? false
                        : metadata.typeAsString === 'MultiChoice' ? [] : '';
            }
        }
        this.setState({
            selectedItemId: 0,
            selectedMode: 'new',
            editingItemId: 0,
            editingValues: values,
            editingErrors: {},
            error: null
        });
        this.props.onSelectionChange(0, 'new');
    };
    GridControl.prototype.cancelRowEdit = function () {
        this.setState({
            editingItemId: -1,
            editingValues: {},
            editingErrors: {},
            saving: false,
            selectedMode: 'view'
        });
        this.props.onSelectionChange(this.state.selectedItemId, 'view');
    };
    GridControl.prototype.updateEditingValue = function (fieldName, value) {
        var values = Object.assign({}, this.state.editingValues);
        var errors = Object.assign({}, this.state.editingErrors);
        values[fieldName] = value;
        delete errors[fieldName];
        this.setState({ editingValues: values, editingErrors: errors });
    };
    GridControl.prototype.validateEditingValues = function () {
        var errors = {};
        var fields = this.getDisplayFields();
        var schemaFields = this.getGridSchemaFields();
        var validationFields = [];
        var validationValues = {};
        for (var schemaIndex = 0; schemaIndex < schemaFields.length; schemaIndex += 1) {
            var validationField = schemaFields[schemaIndex];
            var validationFieldId = validationField.id || validationField.fieldName;
            validationFields.push({ id: validationFieldId, fieldName: validationField.fieldName, label: validationField.label });
            validationValues[validationFieldId] = this.state.editingValues[validationField.fieldName];
        }
        for (var i = 0; i < fields.length; i += 1) {
            var metadata = this.getGridFieldMetadata(fields[i]);
            if (!metadata || !this.isEditableGridField(fields[i])) {
                continue;
            }
            var value = this.state.editingValues[metadata.internalName];
            var isEmpty = value === undefined || value === null || value === ''
                || (Array.isArray(value) && value.length === 0);
            var schemaField = this.getGridSchemaField(metadata.internalName);
            if (metadata.required && isEmpty) {
                errors[metadata.internalName] = schemaField && schemaField.requiredMessage
                    ? schemaField.requiredMessage : strings.RuntimeFieldRequired;
                continue;
            }
            if (!schemaField || !Array.isArray(schemaField.validation)) {
                continue;
            }
            for (var ruleIndex = 0; ruleIndex < schemaField.validation.length; ruleIndex += 1) {
                var rule = schemaField.validation[ruleIndex] || {};
                if (rule.applyWhen) {
                    try {
                        if (!GridValidation_1.evaluateGridValidationExpression(String(rule.applyWhen), validationFields, validationValues)) {
                            continue;
                        }
                    }
                    catch (applyWhenError) {
                        errors[metadata.internalName] = applyWhenError && applyWhenError.message
                            ? 'Validation condition error: ' + applyWhenError.message : strings.RuntimeFieldInvalid;
                        break;
                    }
                }
                if (isEmpty && rule.type !== 'required') {
                    continue;
                }
                var textValue = String(value);
                var numericValue = Number(value);
                var ruleValue = Number(rule.value);
                var failed = rule.type === 'required' ? isEmpty
                    : rule.type === 'minLength' ? textValue.length < ruleValue
                        : rule.type === 'maxLength' ? textValue.length > ruleValue
                            : rule.type === 'min' ? numericValue < ruleValue
                                : rule.type === 'max' ? numericValue > ruleValue : false;
                if (rule.type === 'pattern') {
                    try {
                        failed = !(new RegExp(String(rule.value || ''))).test(textValue);
                    }
                    catch (patternError) {
                        errors[metadata.internalName] = patternError && patternError.message
                            ? 'Invalid validation pattern: ' + patternError.message : strings.RuntimeFieldInvalid;
                        break;
                    }
                }
                if (rule.type === 'custom') {
                    try {
                        failed = !GridValidation_1.evaluateGridValidationExpression(String(rule.value || ''), validationFields, validationValues);
                    }
                    catch (customError) {
                        errors[metadata.internalName] = customError && customError.message
                            ? 'Validation expression error: ' + customError.message : strings.RuntimeFieldInvalid;
                        break;
                    }
                }
                if (failed) {
                    errors[metadata.internalName] = String(rule.message || strings.RuntimeFieldInvalid);
                    break;
                }
            }
        }
        var advancedRules = this.getAdvancedValidationRules();
        for (var advancedIndex = 0; advancedIndex < advancedRules.length; advancedIndex += 1) {
            var advancedRule = advancedRules[advancedIndex];
            if (!advancedRule || !advancedRule.expression || !advancedRule.message) {
                continue;
            }
            try {
                if (GridValidation_1.evaluateGridValidationExpression(advancedRule.expression, validationFields, validationValues)) {
                    continue;
                }
                var targetFieldName = String(advancedRule.targetField || '').toLowerCase();
                var resolvedTarget = '';
                for (var targetIndex = 0; targetIndex < validationFields.length; targetIndex += 1) {
                    var target = validationFields[targetIndex];
                    if (String(target.id || '').toLowerCase() === targetFieldName
                        || String(target.fieldName || '').toLowerCase() === targetFieldName
                        || String(target.label || '').toLowerCase() === targetFieldName) {
                        resolvedTarget = target.fieldName;
                        break;
                    }
                }
                errors[resolvedTarget || '__form'] = advancedRule.message;
            }
            catch (advancedError) {
                errors['__form'] = advancedError && advancedError.message
                    ? 'Validation rule error: ' + advancedError.message : strings.RuntimeFieldInvalid;
            }
        }
        return errors;
    };
    GridControl.prototype.buildEditingPayload = function () {
        var payload = {};
        var fields = this.getDisplayFields();
        for (var i = 0; i < fields.length; i += 1) {
            var metadata = this.getGridFieldMetadata(fields[i]);
            if (!metadata || !this.isEditableGridField(fields[i])) {
                continue;
            }
            var value = this.state.editingValues[metadata.internalName];
            var storageMetadata = this.state.fieldMetadataByName[String(metadata.internalName || '').toLowerCase()] || metadata;
            var storageType = storageMetadata.typeAsString;
            if (storageType === 'Number' || storageType === 'Currency' || storageType === 'Integer') {
                payload[metadata.internalName] = value === '' ? null : Number(value);
            }
            else if (storageType === 'Boolean') {
                payload[metadata.internalName] = value === true;
            }
            else if (storageType === 'MultiChoice') {
                payload[metadata.internalName] = { results: Array.isArray(value) ? value : [] };
            }
            else if (storageType === 'DateTime') {
                var dateValue = String(value || '');
                payload[metadata.internalName] = !dateValue ? null
                    : metadata.displayFormat === 0 ? new Date(dateValue + 'T00:00:00Z').toISOString()
                        : metadata.displayFormat === 2 ? new Date('2000-01-01T' + dateValue + ':00Z').toISOString()
                            : new Date(dateValue + 'Z').toISOString();
            }
            else if (storageType === 'URL' || storageType === 'Hyperlink') {
                payload[metadata.internalName] = value ? { Url: String(value), Description: String(value) } : null;
            }
            else if (storageType === 'Text' || storageType === 'Note') {
                payload[metadata.internalName] = value === undefined || value === null ? '' : String(value);
            }
            else {
                payload[metadata.internalName] = value === undefined || value === null ? '' : value;
            }
        }
        return payload;
    };
    GridControl.prototype.saveEditingRow = function () {
        return __awaiter(this, void 0, void 0, function () {
            var validationErrors, listUrl, payload, response, responseText, error_5, saveError;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        validationErrors = this.validateEditingValues();
                        if (Object.keys(validationErrors).length > 0) {
                            this.setState({ editingErrors: validationErrors });
                            return [2 /*return*/];
                        }
                        this.setState({ saving: true, error: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 9, , 10]);
                        listUrl = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items";
                        payload = this.buildEditingPayload();
                        if (!(this.state.editingItemId > 0)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.props.context.spHttpClient.post(listUrl + '(' + this.state.editingItemId + ')', sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata',
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: JSON.stringify(payload)
                            })];
                    case 2:
                        response = _a.sent();
                        return [3 /*break*/, 5];
                    case 3: return [4 /*yield*/, this.props.context.spHttpClient.post(listUrl, sp_http_1.SPHttpClient.configurations.v1, {
                            headers: {
                                Accept: 'application/json;odata=nometadata',
                                'Content-Type': 'application/json;odata=nometadata'
                            },
                            body: JSON.stringify(payload)
                        })];
                    case 4:
                        response = _a.sent();
                        _a.label = 5;
                    case 5:
                        if (!!response.ok) return [3 /*break*/, 7];
                        return [4 /*yield*/, response.text()];
                    case 6:
                        responseText = _a.sent();
                        throw new Error(responseText || strings.RuntimeSaveFailed);
                    case 7:
                        this.setState({
                            editingItemId: -1,
                            editingValues: {},
                            editingErrors: {},
                            saving: false,
                            selectedMode: 'view'
                        });
                        return [4 /*yield*/, this.loadRows()];
                    case 8:
                        _a.sent();
                        return [3 /*break*/, 10];
                    case 9:
                        error_5 = _a.sent();
                        saveError = error_5;
                        this.setState({
                            saving: false,
                            error: saveError && saveError.message ? saveError.message : strings.RuntimeSaveFailed
                        });
                        return [3 /*break*/, 10];
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    GridControl.prototype.renderEditingControl = function (field) {
        var _this = this;
        var metadata = this.getGridFieldMetadata(field);
        if (!metadata || !this.isEditableGridField(field)) {
            return null;
        }
        var value = this.state.editingValues[metadata.internalName];
        var error = this.state.editingErrors[metadata.internalName];
        var schemaField = this.getGridSchemaField(metadata.internalName);
        var config = schemaField && schemaField.config ? schemaField.config : {};
        var control;
        if (metadata.typeAsString === 'Boolean') {
            control = React.createElement("input", { type: "checkbox", checked: value === true, title: metadata.description, onChange: function (ev) { return _this.updateEditingValue(metadata.internalName, ev.currentTarget.checked); } });
        }
        else if (metadata.typeAsString === 'Note') {
            control = React.createElement("textarea", { value: String(value || ''), title: metadata.description, placeholder: String(config.placeholder || ''), maxLength: config.maxLength, onChange: function (ev) { return _this.updateEditingValue(metadata.internalName, ev.currentTarget.value); } });
        }
        else if (metadata.typeAsString === 'Choice') {
            control = (React.createElement("select", { value: String(value || ''), title: metadata.description, onChange: function (ev) { return _this.updateEditingValue(metadata.internalName, ev.currentTarget.value); } },
                React.createElement("option", { value: "" }),
                metadata.choices.map(function (choice) { return React.createElement("option", { key: choice, value: choice }, choice); })));
        }
        else if (metadata.typeAsString === 'MultiChoice') {
            control = (React.createElement("select", { multiple: true, value: Array.isArray(value) ? value : [], title: metadata.description, onChange: function (ev) {
                    var selected = [];
                    var options = ev.currentTarget.options;
                    for (var optionIndex = 0; optionIndex < options.length; optionIndex += 1) {
                        if (options[optionIndex].selected) {
                            selected.push(options[optionIndex].value);
                        }
                    }
                    _this.updateEditingValue(metadata.internalName, selected);
                } }, metadata.choices.map(function (choice) { return React.createElement("option", { key: choice, value: choice }, choice); })));
        }
        else {
            var inputType = metadata.typeAsString === 'DateTime' ? (metadata.displayFormat === 0 ? 'date' : metadata.displayFormat === 2 ? 'time' : 'datetime-local')
                : (metadata.typeAsString === 'Number' || metadata.typeAsString === 'Currency' || metadata.typeAsString === 'Integer') ? 'number'
                    : metadata.typeAsString === 'URL' ? 'url' : 'text';
            var numberStep = config.decimals !== undefined && Number(config.decimals) > 0
                ? String(1 / Math.pow(10, Number(config.decimals))) : metadata.typeAsString === 'Integer' ? '1' : 'any';
            control = React.createElement("input", { type: inputType, value: String(value || ''), title: metadata.description, placeholder: String(config.placeholder || ''), maxLength: config.maxLength, min: config.min, max: config.max, step: inputType === 'number' ? numberStep : undefined, onChange: function (ev) { return _this.updateEditingValue(metadata.internalName, ev.currentTarget.value); } });
        }
        return (React.createElement("div", { className: joinClassNames(['gc-cell-editor', error ? 'gc-cell-editor-error' : '']) },
            control,
            error && React.createElement("div", { className: "gc-field-error" }, error)));
    };
    GridControl.prototype.renderEditingRow = function (displayFields) {
        var _this = this;
        var isNew = this.state.editingItemId === 0;
        return (React.createElement("tr", { className: "lc-row gc-row-editing", onClick: function (ev) { return ev.stopPropagation(); } },
            this.props.showDelete && React.createElement("td", { className: "gc-selection-cell" }),
            displayFields.map(function (field) {
                var editor = _this.renderEditingControl(field);
                return React.createElement("td", { key: field.Name, style: _this.getConfiguredColumnStyle(field) }, editor || (isNew ? null : strings.RuntimeReadOnlyCell));
            }),
            React.createElement("td", { className: "gc-row-actions" },
                React.createElement("button", { type: "button", disabled: this.state.saving, onClick: function () { return _this.saveEditingRow(); } }, this.state.saving ? strings.RuntimeSaving : strings.RuntimeSave),
                React.createElement("button", { type: "button", disabled: this.state.saving, onClick: function () { return _this.cancelRowEdit(); } }, strings.RuntimeCancel),
                this.state.editingErrors['__form'] && React.createElement("div", { className: "gc-field-error" }, this.state.editingErrors['__form']))));
    };
    GridControl.prototype.deleteSelected = function () {
        return __awaiter(this, void 0, void 0, function () {
            var itemIds, failedItemIds, itemIndex, itemId, endpoint, response, _itemDeleteError_1, remainingSelectedItemId, error_6, deleteError;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        itemIds = this.state.selectedItemIds.slice(0);
                        if (itemIds.length === 0) {
                            return [2 /*return*/];
                        }
                        if (typeof window !== 'undefined' && !window.confirm(formatString(strings.RuntimeDeleteConfirm, itemIds.length))) {
                            return [2 /*return*/];
                        }
                        this.logDiagnostic('Deleting items. itemIds=' + itemIds.join(','));
                        this.setState({ deleting: true, deleteMessage: '', error: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 9, , 10]);
                        failedItemIds = [];
                        itemIndex = 0;
                        _a.label = 2;
                    case 2:
                        if (!(itemIndex < itemIds.length)) return [3 /*break*/, 7];
                        itemId = itemIds[itemIndex];
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items(" + itemId + ")";
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, this.props.context.spHttpClient.post(endpoint, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'DELETE'
                                }
                            })];
                    case 4:
                        response = _a.sent();
                        if (!response.ok) {
                            failedItemIds.push(itemId);
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        _itemDeleteError_1 = _a.sent();
                        failedItemIds.push(itemId);
                        return [3 /*break*/, 6];
                    case 6:
                        itemIndex += 1;
                        return [3 /*break*/, 2];
                    case 7:
                        remainingSelectedItemId = failedItemIds.indexOf(this.state.selectedItemId) >= 0 ? this.state.selectedItemId : 0;
                        this.setState({
                            selectedItemId: remainingSelectedItemId,
                            selectedItemIds: failedItemIds,
                            selectedMode: 'view',
                            deleting: false
                        });
                        this.props.onSelectionChange(remainingSelectedItemId, 'view');
                        this.logDiagnostic('Bulk delete completed. deleted=' + String(itemIds.length - failedItemIds.length) + ', failed=' + String(failedItemIds.length));
                        return [4 /*yield*/, this.loadRows()];
                    case 8:
                        _a.sent();
                        if (failedItemIds.length > 0) {
                            this.setState({ deleteMessage: formatString(strings.RuntimeDeleteFailedCount, failedItemIds.length) });
                        }
                        return [3 /*break*/, 10];
                    case 9:
                        error_6 = _a.sent();
                        deleteError = error_6;
                        this.setState({
                            deleting: false,
                            deleteMessage: deleteError && deleteError.message ? deleteError.message : strings.RuntimeDeleteFailed,
                            error: deleteError && deleteError.message ? deleteError.message : strings.RuntimeDeleteFailed
                        });
                        this.logDiagnostic('Delete failed: ' + (deleteError && deleteError.message ? deleteError.message : String(deleteError)));
                        return [3 /*break*/, 10];
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    GridControl.prototype.logDiagnostic = function (message) {
        if (this.props.enableDiagnostics === false) {
            return;
        }
        console.log('[GridControl] ' + message);
    };
    GridControl.prototype.getCellMarkup = function (row, field) {
        var value = this.getRowFieldValue(row, field);
        if (value === undefined || value === null || value === '') {
            return null;
        }
        var text = this.stringifyCellValue(value);
        if (text.indexOf('<') >= 0 && text.indexOf('>') >= 0) {
            return { __html: text };
        }
        return { __html: text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') };
    };
    GridControl.prototype.isTitleField = function (field) {
        var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
        var displayName = String(field.DisplayName || '').toLowerCase();
        // RenderListDataAsStream often returns link-capable title/name aliases.
        return fieldName === 'title'
            || fieldName === 'linktitle'
            || fieldName === 'linktitlenomenu'
            || fieldName === 'linkfilename'
            || fieldName === 'linkfilename2'
            || fieldName === 'fileleafref'
            || displayName === 'title'
            || displayName === 'name';
    };
    GridControl.prototype.getCellPlainText = function (row, field) {
        var value = this.getRowFieldValue(row, field);
        if (value === undefined || value === null || value === '') {
            return '';
        }
        return this.stringifyCellValue(value).replace(/<[^>]*>/g, '').trim();
    };
    GridControl.prototype.getItemLinkUrl = function (row) {
        var itemId = this.getRowItemId(row);
        var targetPageUrl = String(this.props.linkTargetPageUrl || '').trim();
        var targetIdParam = String(this.props.linkTargetIdParam || 'itemid').trim() || 'itemid';
        if (targetPageUrl && itemId > 0) {
            var normalizedTargetUrl = targetPageUrl;
            if (!/^https?:\/\//i.test(normalizedTargetUrl)) {
                normalizedTargetUrl = normalizedTargetUrl.charAt(0) === '/'
                    ? this.getWebUrl() + normalizedTargetUrl
                    : this.getWebUrl() + '/' + normalizedTargetUrl;
            }
            var embeddedPageUrl = resolveEmbeddedPageUrl(normalizedTargetUrl);
            if (embeddedPageUrl) {
                if (/^https?:\/\//i.test(embeddedPageUrl)) {
                    normalizedTargetUrl = embeddedPageUrl;
                }
                else if (embeddedPageUrl.charAt(0) === '/') {
                    normalizedTargetUrl = this.getWebUrl() + embeddedPageUrl;
                }
                else {
                    normalizedTargetUrl = this.getWebUrl() + '/' + embeddedPageUrl;
                }
            }
            if (isWorkbenchUrl(normalizedTargetUrl) || isSharePointWrapperUrl(normalizedTargetUrl)) {
                var sourceUrl = resolveEmbeddedPageUrl(normalizedTargetUrl);
                if (sourceUrl && !isWorkbenchUrl(sourceUrl) && !isSharePointWrapperUrl(sourceUrl)) {
                    normalizedTargetUrl = sourceUrl;
                }
                else if (typeof window !== 'undefined' && window.location && !isWorkbenchUrl(window.location.pathname)) {
                    normalizedTargetUrl = window.location.protocol + '//' + window.location.host + stripQueryAndHash(window.location.pathname);
                }
                else {
                    normalizedTargetUrl = '';
                }
            }
            if (normalizedTargetUrl) {
                var targetUrlWithItem = appendQueryParam(normalizedTargetUrl, targetIdParam, String(itemId));
                if (this.props.includeReturnUrlParam === true && typeof window !== 'undefined' && window.location && window.location.href) {
                    return appendQueryParam(targetUrlWithItem, 'return', window.location.href);
                }
                return targetUrlWithItem;
            }
        }
        var fileRef = String(row.FileRef || '').trim();
        if (fileRef) {
            if (/^https?:\/\//i.test(fileRef)) {
                return fileRef;
            }
            // List item pseudo files usually end with _.000; avoid those and use form URL instead.
            if (fileRef.charAt(0) === '/' && !/\/\d+_\.000$/i.test(fileRef)) {
                return this.getWebUrl() + fileRef;
            }
        }
        if (itemId > 0) {
            return this.getWebUrl() + '/Lists/' + encodeURIComponent(this.props.listName) + '/DispForm.aspx?ID=' + itemId;
        }
        return '';
    };
    GridControl.prototype.getFieldKey = function (field) {
        return String(field.Name || field.RealFieldName || '').trim();
    };
    GridControl.prototype.getFilterOperatorOptions = function () {
        return [
            { key: 'eq', label: strings.RuntimeFilterOperatorEquals },
            { key: 'ne', label: strings.RuntimeFilterOperatorNotEquals },
            { key: 'contains', label: strings.RuntimeFilterOperatorContains },
            { key: 'notcontains', label: strings.RuntimeFilterOperatorNotContains },
            { key: 'startswith', label: strings.RuntimeFilterOperatorStartsWith },
            { key: 'endswith', label: strings.RuntimeFilterOperatorEndsWith },
            { key: 'gt', label: strings.RuntimeFilterOperatorGreaterThan },
            { key: 'ge', label: strings.RuntimeFilterOperatorGreaterThanOrEqual },
            { key: 'lt', label: strings.RuntimeFilterOperatorLessThan },
            { key: 'le', label: strings.RuntimeFilterOperatorLessThanOrEqual }
        ];
    };
    GridControl.prototype.toggleSort = function (field) {
        var fieldKey = this.getFieldKey(field);
        if (!fieldKey) {
            return;
        }
        var nextDirection = 'asc';
        if (this.state.sortFieldName === fieldKey && this.state.sortDirection === 'asc') {
            nextDirection = 'desc';
        }
        else if (this.state.sortFieldName === fieldKey && this.state.sortDirection === 'desc') {
            nextDirection = '';
        }
        this.setState({
            sortFieldName: nextDirection ? fieldKey : '',
            sortDirection: nextDirection,
            currentPage: 0
        });
    };
    GridControl.prototype.openFilter = function (field) {
        var fieldKey = this.getFieldKey(field);
        if (!fieldKey) {
            return;
        }
        var existing = this.state.columnFilters[fieldKey];
        this.setState({
            activeFilterFieldName: fieldKey,
            draftFilterOperator: existing ? existing.operator : 'contains',
            draftFilterValue: existing ? existing.value : ''
        });
    };
    GridControl.prototype.closeFilter = function () {
        this.setState({ activeFilterFieldName: '' });
    };
    GridControl.prototype.applyActiveFilter = function () {
        var fieldKey = String(this.state.activeFilterFieldName || '');
        if (!fieldKey) {
            return;
        }
        var nextFilters = {};
        var key;
        for (key in this.state.columnFilters) {
            if (Object.prototype.hasOwnProperty.call(this.state.columnFilters, key)) {
                nextFilters[key] = this.state.columnFilters[key];
            }
        }
        var value = String(this.state.draftFilterValue || '').trim();
        if (!value) {
            delete nextFilters[fieldKey];
        }
        else {
            nextFilters[fieldKey] = {
                operator: this.state.draftFilterOperator,
                value: value
            };
        }
        this.setState({
            columnFilters: nextFilters,
            activeFilterFieldName: '',
            currentPage: 0
        });
    };
    GridControl.prototype.clearActiveFilter = function () {
        var fieldKey = String(this.state.activeFilterFieldName || '');
        if (!fieldKey) {
            return;
        }
        var nextFilters = {};
        var key;
        for (key in this.state.columnFilters) {
            if (Object.prototype.hasOwnProperty.call(this.state.columnFilters, key) && key !== fieldKey) {
                nextFilters[key] = this.state.columnFilters[key];
            }
        }
        this.setState({
            columnFilters: nextFilters,
            draftFilterOperator: 'contains',
            draftFilterValue: '',
            activeFilterFieldName: '',
            currentPage: 0
        });
    };
    GridControl.prototype.compareComparableValues = function (leftValue, rightValue) {
        var leftNumber = Number(leftValue);
        var rightNumber = Number(rightValue);
        if (!isNaN(leftNumber) && !isNaN(rightNumber)) {
            if (leftNumber > rightNumber) {
                return 1;
            }
            if (leftNumber < rightNumber) {
                return -1;
            }
            return 0;
        }
        var leftDate = Date.parse(leftValue);
        var rightDate = Date.parse(rightValue);
        if (!isNaN(leftDate) && !isNaN(rightDate)) {
            if (leftDate > rightDate) {
                return 1;
            }
            if (leftDate < rightDate) {
                return -1;
            }
            return 0;
        }
        var leftText = leftValue.toLowerCase();
        var rightText = rightValue.toLowerCase();
        if (leftText > rightText) {
            return 1;
        }
        if (leftText < rightText) {
            return -1;
        }
        return 0;
    };
    GridControl.prototype.rowMatchesFilter = function (row, field, filter) {
        var valueText = this.getCellPlainText(row, field);
        var candidate = String(valueText || '').trim();
        var query = String(filter.value || '').trim();
        if (!query) {
            return true;
        }
        var normalizedCandidate = candidate.toLowerCase();
        var normalizedQuery = query.toLowerCase();
        var compareResult = this.compareComparableValues(candidate, query);
        if (filter.compareDateOnly) {
            var candidateDate = new Date(candidate);
            if (!isNaN(candidateDate.getTime())) {
                normalizedCandidate = formatLocalDate(candidateDate).toLowerCase();
                compareResult = this.compareComparableValues(normalizedCandidate, query);
            }
        }
        switch (filter.operator) {
            case 'eq':
                return normalizedCandidate === normalizedQuery;
            case 'ne':
                return normalizedCandidate !== normalizedQuery;
            case 'contains':
                return normalizedCandidate.indexOf(normalizedQuery) >= 0;
            case 'notcontains':
                return normalizedCandidate.indexOf(normalizedQuery) < 0;
            case 'startswith':
                return normalizedCandidate.indexOf(normalizedQuery) === 0;
            case 'endswith':
                return normalizedCandidate.lastIndexOf(normalizedQuery) === normalizedCandidate.length - normalizedQuery.length;
            case 'gt':
                return compareResult > 0;
            case 'ge':
                return compareResult >= 0;
            case 'lt':
                return compareResult < 0;
            case 'le':
                return compareResult <= 0;
            default:
                return true;
        }
    };
    GridControl.prototype.parsePresetFilterConditions = function () {
        var source = String(this.props.filterJson || '').trim();
        if (!source) {
            return [];
        }
        try {
            var parsed = JSON.parse(source);
            if (!Array.isArray(parsed)) {
                return [];
            }
            var conditions = [];
            for (var i = 0; i < parsed.length; i += 1) {
                var item = parsed[i] || {};
                var field = String(item.field || '').trim();
                if (!field) {
                    continue;
                }
                conditions.push({
                    field: field,
                    operator: normalizeFilterOperator(item.operator),
                    logical: normalizeFilterLogical(item.logical),
                    valueType: String(item.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static',
                    value: item.value
                });
            }
            return conditions;
        }
        catch (_parseError) {
            this.logDiagnostic('Preset filter JSON is invalid; skipping preset filters.');
            return [];
        }
    };
    GridControl.prototype.resolveFieldByReference = function (fieldsByKey, fieldRef) {
        var direct = fieldsByKey[fieldRef];
        if (direct) {
            return direct;
        }
        var normalizedRef = String(fieldRef || '').trim().toLowerCase();
        if (!normalizedRef) {
            return undefined;
        }
        var key;
        for (key in fieldsByKey) {
            if (!Object.prototype.hasOwnProperty.call(fieldsByKey, key)) {
                continue;
            }
            var candidate = fieldsByKey[key];
            if (!candidate) {
                continue;
            }
            var byName = String(candidate.Name || '').toLowerCase();
            var byRealName = String(candidate.RealFieldName || '').toLowerCase();
            var byDisplay = String(candidate.DisplayName || '').toLowerCase();
            if (normalizedRef === byName || normalizedRef === byRealName || normalizedRef === byDisplay) {
                return candidate;
            }
        }
        return undefined;
    };
    GridControl.prototype.rowMatchesPresetConditions = function (row, fieldsByKey, conditions) {
        if (conditions.length === 0) {
            return true;
        }
        var aggregate = null;
        for (var i = 0; i < conditions.length; i += 1) {
            var condition = conditions[i];
            var field = this.resolveFieldByReference(fieldsByKey, condition.field);
            if (!field) {
                continue;
            }
            var resolvedValue = this.resolvePresetFilterValue(condition);
            var matches = this.rowMatchesFilter(row, field, {
                operator: normalizeFilterOperator(condition.operator),
                value: resolvedValue.value,
                compareDateOnly: resolvedValue.compareDateOnly
            });
            if (aggregate === null) {
                aggregate = matches;
            }
            else if (normalizeFilterLogical(condition.logical) === 'or') {
                aggregate = aggregate || matches;
            }
            else {
                aggregate = aggregate && matches;
            }
        }
        return aggregate === null ? true : aggregate;
    };
    GridControl.prototype.resolvePresetFilterValue = function (condition) {
        var rawValue = condition.value === undefined || condition.value === null ? '' : String(condition.value).trim();
        if (String(condition.valueType || '').toLowerCase() !== 'expression') {
            return { value: rawValue, compareDateOnly: false };
        }
        var expression = rawValue.toLowerCase();
        var pageContext = this.props.context && this.props.context.pageContext;
        var user = pageContext && pageContext.user;
        var legacyContext = pageContext && pageContext.legacyPageContext;
        if (expression === 'today') {
            return { value: formatLocalDate(new Date()), compareDateOnly: true };
        }
        if (expression === 'now') {
            return { value: new Date().toISOString(), compareDateOnly: false };
        }
        var dateMatch = /^date\(\s*([+-]?\d+)\s*\)$/.exec(expression);
        if (dateMatch) {
            var dateValue = new Date();
            dateValue.setDate(dateValue.getDate() + parseInt(dateMatch[1], 10));
            return { value: formatLocalDate(dateValue), compareDateOnly: true };
        }
        if (expression === 'me.email') {
            return { value: String(user && user.email || ''), compareDateOnly: false };
        }
        if (expression === 'me.login') {
            return { value: String(user && user.loginName || ''), compareDateOnly: false };
        }
        if (expression === 'me.id') {
            return { value: String(legacyContext && legacyContext.userId || ''), compareDateOnly: false };
        }
        if (expression === 'me') {
            return { value: String(user && (user.displayName || user.loginName) || ''), compareDateOnly: false };
        }
        return { value: rawValue, compareDateOnly: false };
    };
    GridControl.prototype.getProcessedRows = function () {
        var _this = this;
        var rows = this.state.rows.slice(0);
        var fieldsByKey = {};
        for (var i = 0; i < this.state.fields.length; i += 1) {
            var field = this.state.fields[i];
            var fieldKey = this.getFieldKey(field);
            if (fieldKey) {
                fieldsByKey[fieldKey] = field;
            }
        }
        var presetConditions = this.parsePresetFilterConditions();
        if (presetConditions.length > 0) {
            rows = rows.filter(function (row) {
                return _this.rowMatchesPresetConditions(row, fieldsByKey, presetConditions);
            });
        }
        var filterKeys = Object.keys(this.state.columnFilters || {});
        if (filterKeys.length > 0) {
            rows = rows.filter(function (row) {
                for (var index = 0; index < filterKeys.length; index += 1) {
                    var key = filterKeys[index];
                    var filter = _this.state.columnFilters[key];
                    var field = fieldsByKey[key];
                    if (!filter || !field) {
                        continue;
                    }
                    if (!_this.rowMatchesFilter(row, field, filter)) {
                        return false;
                    }
                }
                return true;
            });
        }
        if (this.state.sortFieldName && this.state.sortDirection) {
            var sortField = fieldsByKey[this.state.sortFieldName];
            if (sortField) {
                rows.sort(function (leftRow, rightRow) {
                    var left = _this.getCellPlainText(leftRow, sortField);
                    var right = _this.getCellPlainText(rightRow, sortField);
                    var comparison = _this.compareComparableValues(left, right);
                    return _this.state.sortDirection === 'asc' ? comparison : (comparison * -1);
                });
            }
        }
        return rows;
    };
    GridControl.prototype.parseConditionalStyleRules = function () {
        var source = String(this.props.conditionalStyleJson || '').trim();
        if (!source) {
            return [];
        }
        try {
            var parsed = JSON.parse(source);
            if (!Array.isArray(parsed)) {
                return [];
            }
            var groupedRules = {};
            var orderedRuleKeys = [];
            for (var i = 0; i < parsed.length; i += 1) {
                var item = parsed[i] || {};
                var ruleKey = toCssString(item.rule || item.ruleName || 'rule-' + String(i + 1));
                if (!ruleKey) {
                    ruleKey = 'rule-' + String(i + 1);
                }
                var scope = String(item.scope || 'row').trim().toLowerCase() === 'column' ? 'column' : 'row';
                var applyField = toCssString(item.applyField);
                var conditionField = toCssString(item.conditionField || item.field);
                if (!conditionField) {
                    continue;
                }
                var conditionOperator = normalizeFilterOperator(item.operator || item.conditionOperator);
                var conditionLogical = normalizeFilterLogical(item.logical || item.conditionLogical);
                var conditionValue = toCssString(item.value || item.conditionValue);
                var styleDefinition = {
                    backgroundColor: toCssString(item.backgroundColor || (item.style && item.style.backgroundColor) || ''),
                    color: toCssString(item.color || item.foregroundColor || (item.style && item.style.color) || ''),
                    fontFamily: toCssString(item.fontFamily || (item.style && item.style.fontFamily) || ''),
                    fontSize: toCssString(item.fontSize || (item.style && item.style.fontSize) || ''),
                    fontStyle: toCssString(item.fontStyle || (item.style && item.style.fontStyle) || ''),
                    fontWeight: toCssString(item.fontWeight || (item.style && item.style.fontWeight) || ''),
                    textAlign: toCssString(item.textAlign || (item.style && item.style.textAlign) || '')
                };
                if (!groupedRules[ruleKey]) {
                    groupedRules[ruleKey] = {
                        key: ruleKey,
                        scope: scope,
                        applyField: applyField,
                        enabled: item.enabled !== false,
                        priority: toPriorityNumber(item.priority, 100),
                        sourceOrder: i,
                        style: styleDefinition,
                        conditions: []
                    };
                    orderedRuleKeys.push(ruleKey);
                }
                groupedRules[ruleKey].scope = scope;
                groupedRules[ruleKey].applyField = applyField;
                groupedRules[ruleKey].enabled = item.enabled !== false;
                groupedRules[ruleKey].priority = toPriorityNumber(item.priority, groupedRules[ruleKey].priority);
                groupedRules[ruleKey].style = styleDefinition;
                groupedRules[ruleKey].conditions.push({
                    field: conditionField,
                    operator: conditionOperator,
                    logical: conditionLogical,
                    valueType: String(item.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static',
                    value: conditionValue
                });
            }
            return orderedRuleKeys
                .map(function (ruleKey) { return groupedRules[ruleKey]; })
                .filter(function (rule) { return !!rule && rule.conditions.length > 0; })
                .sort(function (leftRule, rightRule) {
                if (leftRule.priority !== rightRule.priority) {
                    return leftRule.priority - rightRule.priority;
                }
                return leftRule.sourceOrder - rightRule.sourceOrder;
            });
        }
        catch (_parseError) {
            this.logDiagnostic('Conditional style JSON is invalid; skipping conditional styles.');
            return [];
        }
    };
    GridControl.prototype.toReactCssStyle = function (styleDefinition) {
        var style = {};
        var backgroundColor = toCssString(styleDefinition.backgroundColor);
        var color = toCssString(styleDefinition.color);
        var fontFamily = toCssString(styleDefinition.fontFamily);
        var fontSize = toCssString(styleDefinition.fontSize);
        var fontStyle = toCssString(styleDefinition.fontStyle);
        var fontWeight = toCssString(styleDefinition.fontWeight);
        var textAlign = toCssString(styleDefinition.textAlign);
        if (backgroundColor) {
            style.backgroundColor = backgroundColor;
        }
        if (color) {
            style.color = color;
        }
        if (fontFamily) {
            style.fontFamily = fontFamily;
        }
        if (fontSize) {
            style.fontSize = fontSize;
        }
        if (fontStyle) {
            style.fontStyle = fontStyle;
        }
        if (fontWeight) {
            style.fontWeight = fontWeight;
        }
        if (textAlign) {
            style.textAlign = textAlign;
        }
        return style;
    };
    GridControl.prototype.evaluateConditionalStyleRule = function (row, rule, fieldsByKey) {
        if (!rule || !rule.conditions || rule.conditions.length === 0) {
            return false;
        }
        var aggregate = null;
        for (var i = 0; i < rule.conditions.length; i += 1) {
            var condition = rule.conditions[i];
            var field = this.resolveFieldByReference(fieldsByKey, condition.field);
            if (!field) {
                continue;
            }
            var resolvedValue = this.resolvePresetFilterValue(condition);
            var matches = this.rowMatchesFilter(row, field, {
                operator: condition.operator,
                value: resolvedValue.value,
                compareDateOnly: resolvedValue.compareDateOnly
            });
            if (aggregate === null) {
                aggregate = matches;
            }
            else if (condition.logical === 'or') {
                aggregate = aggregate || matches;
            }
            else {
                aggregate = aggregate && matches;
            }
        }
        return aggregate === null ? false : aggregate;
    };
    GridControl.prototype.getConditionalStyleForRow = function (row, fieldsByKey, rules) {
        var rowStyle = {};
        var columnStylesByFieldKey = {};
        for (var i = 0; i < rules.length; i += 1) {
            var rule = rules[i];
            if (!rule.enabled) {
                continue;
            }
            if (!this.evaluateConditionalStyleRule(row, rule, fieldsByKey)) {
                continue;
            }
            var overrideStyle = this.toReactCssStyle(rule.style);
            if (Object.keys(overrideStyle).length === 0) {
                continue;
            }
            if (rule.scope === 'column') {
                var applyFieldDefinition = this.resolveFieldByReference(fieldsByKey, rule.applyField);
                if (!applyFieldDefinition) {
                    continue;
                }
                var applyFieldKey = this.getFieldKey(applyFieldDefinition);
                if (!applyFieldKey) {
                    continue;
                }
                var columnStyle = columnStylesByFieldKey[applyFieldKey] || {};
                columnStylesByFieldKey[applyFieldKey] = mergeStyleObjects(columnStyle, overrideStyle);
            }
            else {
                rowStyle = mergeStyleObjects(rowStyle, overrideStyle);
            }
        }
        return {
            rowStyle: rowStyle,
            columnStylesByFieldKey: columnStylesByFieldKey
        };
    };
    GridControl.prototype.getMatchedConditionalRuleNames = function (row, fieldsByKey, rules) {
        var names = [];
        for (var i = 0; i < rules.length; i += 1) {
            var rule = rules[i];
            if (!rule.enabled) {
                continue;
            }
            if (this.evaluateConditionalStyleRule(row, rule, fieldsByKey)) {
                names.push(rule.key);
            }
        }
        return names;
    };
    GridControl.prototype.render = function () {
        var _this = this;
        if (!this.props.listName) {
            return React.createElement("div", null, strings.RuntimeNoListSelected);
        }
        var isEditingRow = this.state.editingItemId >= 0;
        var processedRows = this.getProcessedRows();
        var pageSize = this.props.pageSize > 0 ? this.props.pageSize : 0;
        var pageCount = pageSize > 0 ? Math.max(1, Math.ceil(processedRows.length / pageSize)) : 1;
        var currentPage = Math.min(this.state.currentPage, pageCount - 1);
        var visibleRows = pageSize > 0
            ? processedRows.slice(currentPage * pageSize, (currentPage + 1) * pageSize)
            : processedRows;
        var allVisibleItemsChecked = visibleRows.length > 0 && visibleRows.every(function (row) {
            return _this.isItemChecked(_this.getRowItemId(row));
        });
        var conditionalRules = this.parseConditionalStyleRules();
        var hasActiveFilters = Object.keys(this.state.columnFilters || {}).length > 0;
        var displayFields = this.getDisplayFields();
        var fieldsByKey = {};
        for (var i = 0; i < displayFields.length; i += 1) {
            var fieldsByKeyField = displayFields[i];
            var fieldsByKeyValue = this.getFieldKey(fieldsByKeyField);
            if (fieldsByKeyValue) {
                fieldsByKey[fieldsByKeyValue] = fieldsByKeyField;
            }
        }
        var selectedRow = null;
        var selectedMatchedRuleNames = [];
        if (this.props.isEditMode && this.state.selectedItemId > 0) {
            for (var selectedIndex = 0; selectedIndex < processedRows.length; selectedIndex += 1) {
                var selectedCandidate = processedRows[selectedIndex];
                if (this.getRowItemId(selectedCandidate) === this.state.selectedItemId) {
                    selectedRow = selectedCandidate;
                    break;
                }
            }
            if (selectedRow) {
                selectedMatchedRuleNames = this.getMatchedConditionalRuleNames(selectedRow, fieldsByKey, conditionalRules);
            }
        }
        var containerStyle = {
            '--lc-body-text-color': this.props.bodyTextColor || 'inherit',
            '--lc-body-font-family': this.props.bodyFontFamily || 'inherit',
            '--lc-body-font-size': this.props.bodyFontSize || '14px',
            '--lc-body-font-style': this.props.bodyFontStyle || 'normal',
            '--lc-body-font-weight': this.props.bodyFontBold ? 'bold' : 'normal',
            '--lc-body-text-align': this.props.bodyTextAlign || 'left',
            '--lc-selected-text-color': this.props.selectedTextColor || this.props.bodyTextColor || 'inherit',
            '--lc-selected-bg-color': this.props.selectedBackgroundColor || '#eef6ff',
            '--lc-selected-font-style': this.props.selectedFontStyle || this.props.bodyFontStyle || 'normal',
            '--lc-selected-font-weight': this.props.selectedFontBold ? 'bold' : (this.props.bodyFontBold ? 'bold' : 'normal'),
            '--lc-header-text-color': this.props.headerTextColor || 'inherit',
            '--lc-header-bg-color': this.props.headerBackgroundColor || 'transparent',
            '--lc-header-font-family': this.props.headerFontFamily || 'inherit',
            '--lc-header-font-size': this.props.headerFontSize || '14px',
            '--lc-header-font-style': this.props.headerFontStyle || 'normal',
            '--lc-header-font-weight': this.props.headerFontBold ? 'bold' : 'normal',
            '--lc-header-text-align': this.props.headerTextAlign || 'left',
            '--lc-table-bg-color': this.props.tableBackgroundColor || '#ffffff',
            '--lc-table-border-color': this.props.tableBorderColor || '#ccc',
            '--lc-table-border-width': this.props.tableBorderWidth || '1px',
            '--lc-table-corner-radius': this.props.tableCornerStyle === 'rounded'
                ? (typeof this.props.tableCornerRadius === 'number' ? this.props.tableCornerRadius : 8) + 'px'
                : '0px',
            '--lc-row-line-width': (typeof this.props.tableRowLineWidth === 'number' ? this.props.tableRowLineWidth : 1) + 'px',
            '--lc-alternate-row-color': this.props.alternateRowShading
                ? (this.props.alternateRowShadingColor || '#f9f9f9')
                : 'transparent',
            '--lc-button-text-color': this.props.buttonTextColor || '#000000',
            '--lc-button-bg-color': this.props.buttonBackgroundColor || '#f0f0f0',
            '--lc-button-font-family': this.props.buttonFontFamily || 'inherit',
            '--lc-button-font-size': this.props.buttonFontSize || '14px',
            '--lc-button-font-style': this.props.buttonFontStyle || 'normal',
            '--lc-button-font-weight': this.props.buttonFontBold ? 'bold' : 'normal',
            '--lc-button-corner-radius': this.props.buttonCornerStyle === 'rounded'
                ? (typeof this.props.buttonCornerRadius === 'number' ? this.props.buttonCornerRadius : 4) + 'px'
                : '0px',
            '--lc-webpart-bg-color': this.props.webpartBackgroundColor || '#ffffff',
            '--lc-webpart-border-color': this.props.webpartBorderColor || '#ccc',
            '--lc-webpart-border-width': (this.props.webpartBorderWidth || 0) + 'px',
            '--lc-webpart-corner-style': this.props.cornerStyle === 'rounded'
                ? (typeof this.props.cornerRadius === 'number' ? this.props.cornerRadius : 12) + 'px'
                : '0px'
        };
        return (React.createElement("div", { className: "lc-root", style: containerStyle },
            React.createElement("div", { className: "lc-toolbar" },
                this.props.showViewSelector && (React.createElement("label", null,
                    React.createElement("span", { style: { marginRight: '6px' } }, strings.RuntimeViewLabel),
                    React.createElement("select", { value: this.state.selectedViewId, onChange: function (ev) { return _this.setState({ selectedViewId: ev.currentTarget.value }); } }, this.props.views.map(function (view) {
                        return React.createElement("option", { key: view.key, value: view.key }, view.text);
                    })))),
                this.props.showRefresh && React.createElement("button", { type: "button", onClick: function () { return _this.loadRows(); } }, strings.RuntimeRefresh),
                this.props.showAdd && (React.createElement("button", { type: "button", className: "gc-add-row", disabled: isEditingRow, title: strings.RuntimeAddRow, onClick: function () { return _this.beginNewRow(); } },
                    "+ ",
                    strings.RuntimeAddRow)),
                this.props.showDelete && (React.createElement("button", { type: "button", disabled: this.state.selectedItemIds.length === 0 || this.state.deleting || isEditingRow, onClick: function () { return _this.deleteSelected(); } }, this.state.deleting
                    ? strings.RuntimeDeleting
                    : formatString(strings.RuntimeDeleteSelected, this.state.selectedItemIds.length)))),
            this.props.isEditMode && (React.createElement("div", { className: "lc-status" },
                React.createElement("div", null, formatString(strings.RuntimeSelectedItem, this.state.selectedItemId || 0)),
                React.createElement("div", null, formatString(strings.RuntimeSelectedMode, this.state.selectedMode)),
                React.createElement("div", null, formatString(strings.RuntimeConditionalStyleMatches, selectedMatchedRuleNames.length > 0 ? selectedMatchedRuleNames.join(', ') : strings.RuntimeConditionalStyleMatchesNone)),
                React.createElement("div", null, formatString(strings.RuntimeDeleteSelectionCount, this.state.selectedItemIds.length)))),
            this.state.loading && React.createElement("div", null, strings.RuntimeLoading),
            !!this.state.deleteMessage && React.createElement("div", { className: "gc-delete-message" }, this.state.deleteMessage),
            !!this.state.error && React.createElement("div", null, this.state.error),
            !this.state.loading && !this.state.error && processedRows.length === 0 && !isEditingRow && (React.createElement("div", null, hasActiveFilters ? strings.RuntimeNoItemsAfterFilter : strings.RuntimeNoItems)),
            !this.state.loading && !this.state.error && (processedRows.length > 0 || isEditingRow) && (React.createElement("div", { className: "lc-table-wrap" },
                React.createElement("table", { className: "lc-table" },
                    React.createElement("thead", null,
                        React.createElement("tr", null,
                            this.props.showDelete && (React.createElement("th", { className: "gc-selection-header" },
                                React.createElement("input", { type: "checkbox", checked: allVisibleItemsChecked, disabled: visibleRows.length === 0 || this.state.deleting || isEditingRow, title: strings.RuntimeSelectVisibleRows, "aria-label": strings.RuntimeSelectVisibleRows, onChange: function () { return _this.toggleVisibleItemsChecked(visibleRows); } }))),
                            displayFields.map(function (field) {
                                var fieldKey = _this.getFieldKey(field);
                                var sortActive = _this.state.sortFieldName === fieldKey;
                                var filterActive = !!_this.state.columnFilters[fieldKey];
                                var sortIndicator = '';
                                if (sortActive && _this.state.sortDirection === 'asc') {
                                    sortIndicator = ' ▲';
                                }
                                else if (sortActive && _this.state.sortDirection === 'desc') {
                                    sortIndicator = ' ▼';
                                }
                                return (React.createElement("th", { key: field.Name, style: _this.getConfiguredColumnStyle(field), title: (_this.getGridFieldMetadata(field) || {}).description || '' },
                                    React.createElement("div", { className: "lc-header-cell" },
                                        React.createElement("button", { type: "button", className: "lc-header-title", onClick: function (ev) {
                                                ev.stopPropagation();
                                                _this.toggleSort(field);
                                            }, title: strings.RuntimeSortToggle },
                                            field.DisplayName || field.Name || strings.ColumnIdFallback,
                                            sortIndicator),
                                        React.createElement("div", { className: "lc-header-actions" },
                                            React.createElement("button", { type: "button", className: joinClassNames(['lc-header-action', filterActive ? 'lc-header-action-active' : '']), title: strings.RuntimeFilterTitle, onClick: function (ev) {
                                                    ev.preventDefault();
                                                    ev.stopPropagation();
                                                    if (_this.state.activeFilterFieldName === fieldKey) {
                                                        _this.closeFilter();
                                                    }
                                                    else {
                                                        _this.openFilter(field);
                                                    }
                                                } }, strings.RuntimeFilterIcon)),
                                        _this.state.activeFilterFieldName === fieldKey && (React.createElement("div", { className: "lc-filter-popover", onClick: function (ev) {
                                                ev.preventDefault();
                                                ev.stopPropagation();
                                            } },
                                            React.createElement("label", { className: "lc-filter-label" }, strings.RuntimeFilterOperatorLabel),
                                            React.createElement("select", { className: "lc-filter-select", value: _this.state.draftFilterOperator, onChange: function (ev) { return _this.setState({ draftFilterOperator: ev.currentTarget.value }); } }, _this.getFilterOperatorOptions().map(function (option) {
                                                return React.createElement("option", { key: option.key, value: option.key }, option.label);
                                            })),
                                            React.createElement("label", { className: "lc-filter-label" }, strings.RuntimeFilterValueLabel),
                                            React.createElement("input", { className: "lc-filter-input", type: "text", value: _this.state.draftFilterValue, placeholder: strings.RuntimeFilterValuePlaceholder, onChange: function (ev) { return _this.setState({ draftFilterValue: ev.currentTarget.value }); }, onKeyDown: function (ev) {
                                                    if (ev.key === 'Enter') {
                                                        _this.applyActiveFilter();
                                                    }
                                                } }),
                                            React.createElement("div", { className: "lc-filter-actions" },
                                                React.createElement("button", { type: "button", onClick: function () { return _this.applyActiveFilter(); } }, strings.RuntimeFilterApply),
                                                React.createElement("button", { type: "button", onClick: function () { return _this.clearActiveFilter(); } }, strings.RuntimeFilterClear),
                                                React.createElement("button", { type: "button", onClick: function () { return _this.closeFilter(); } }, strings.RuntimeFilterClose)))))));
                            }),
                            React.createElement("th", { className: "gc-actions-header" }, strings.RuntimeActions))),
                    React.createElement("tbody", null,
                        this.state.editingItemId === 0 && this.renderEditingRow(displayFields),
                        visibleRows.map(function (row, index) {
                            var rowItemId = _this.getRowItemId(row);
                            var isSelected = rowItemId > 0 && rowItemId === _this.state.selectedItemId;
                            var conditionalStyle = _this.getConditionalStyleForRow(row, fieldsByKey, conditionalRules);
                            if (_this.state.editingItemId === rowItemId) {
                                return React.cloneElement(_this.renderEditingRow(displayFields), { key: String(rowItemId) });
                            }
                            return (React.createElement("tr", { key: rowItemId > 0 ? String(rowItemId) : String(index), onClick: function () { return _this.selectRow(row); }, className: joinClassNames(['lc-row', isSelected ? 'lc-row-selected' : '', _this.isItemChecked(rowItemId) ? 'gc-row-delete-selected' : '']) },
                                _this.props.showDelete && (React.createElement("td", { className: "gc-selection-cell" },
                                    React.createElement("input", { type: "checkbox", checked: _this.isItemChecked(rowItemId), disabled: _this.state.deleting || isEditingRow, title: strings.RuntimeSelectRowForDelete, "aria-label": strings.RuntimeSelectRowForDelete, onClick: function (ev) { return ev.stopPropagation(); }, onChange: function () { return _this.toggleItemChecked(rowItemId); } }))),
                                displayFields.map(function (field) {
                                    var markup = _this.getCellMarkup(row, field);
                                    var showItemLink = _this.props.showLinkToItem && _this.isTitleField(field);
                                    var itemLinkUrl = showItemLink ? _this.getItemLinkUrl(row) : '';
                                    var itemLinkText = showItemLink ? _this.getCellPlainText(row, field) : '';
                                    var cellFieldKey = _this.getFieldKey(field);
                                    var columnStyle = conditionalStyle.columnStylesByFieldKey[cellFieldKey] || {};
                                    var mergedCellStyle = mergeStyleObjects(mergeStyleObjects(conditionalStyle.rowStyle, columnStyle), _this.getConfiguredColumnStyle(field));
                                    return (React.createElement("td", { key: field.Name, style: mergedCellStyle, title: (_this.getGridFieldMetadata(field) || {}).description || '' }, showItemLink && itemLinkUrl ? (React.createElement("a", { className: "lc-item-link", href: itemLinkUrl, onClick: function (ev) { return ev.stopPropagation(); } }, itemLinkText || strings.RuntimeView)) : (markup ? React.createElement("span", { dangerouslySetInnerHTML: markup }) : null)));
                                }),
                                React.createElement("td", { className: "gc-row-actions" },
                                    React.createElement("button", { type: "button", disabled: isEditingRow || _this.state.deleting, onClick: function (ev) {
                                            ev.stopPropagation();
                                            _this.beginRowEdit(row);
                                        } }, strings.RuntimeEdit))));
                        }))))),
            !this.state.loading && !this.state.error && pageSize > 0 && pageCount > 1 && (React.createElement("div", { className: "lc-pagination" },
                React.createElement("button", { type: "button", disabled: currentPage === 0, onClick: function () { return _this.setState({ currentPage: Math.max(0, currentPage - 1) }); } }, strings.RuntimePreviousPage),
                React.createElement("span", null, strings.RuntimePageStatus.replace('{0}', String(currentPage + 1)).replace('{1}', String(pageCount))),
                React.createElement("button", { type: "button", disabled: currentPage >= pageCount - 1, onClick: function () { return _this.setState({ currentPage: Math.min(pageCount - 1, currentPage + 1) }); } }, strings.RuntimeNextPage)))));
    };
    return GridControl;
}(React.Component));
exports.GridControl = GridControl;

//# sourceMappingURL=GridControl.js.map
