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
var strings = require("ListControlWebPartStrings");
require("./ListControl.css");
var LIST_CONTROL_REFRESH_EVENT = 'spse:listcontrol-refresh';
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
var ListControl = (function (_super) {
    __extends(ListControl, _super);
    function ListControl(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            selectedViewId: props.defaultViewId || _this.getInitialViewId(props.views),
            fields: [],
            rows: [],
            loading: true,
            error: null,
            selectedItemId: 0,
            selectedMode: 'view',
            deleting: false,
            sortFieldName: '',
            sortDirection: '',
            activeFilterFieldName: '',
            draftFilterOperator: 'contains',
            draftFilterValue: '',
            columnFilters: {},
            currentPage: 0,
        };
        _this._refreshEventHandler = _this.handleExternalRefresh.bind(_this);
        return _this;
    }
    ListControl.prototype.componentDidMount = function () {
        this.logDiagnostic('Component mounted. listName=' + String(this.props.listName || '(none)') + ', defaultViewId=' + String(this.props.defaultViewId || '(none)'));
        if (typeof window !== 'undefined' && window.addEventListener) {
            window.addEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
        }
        this.loadRows();
    };
    ListControl.prototype.componentWillUnmount = function () {
        if (typeof window !== 'undefined' && window.removeEventListener) {
            window.removeEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
        }
    };
    ListControl.prototype.componentDidUpdate = function (prevProps, prevState) {
        var _this = this;
        if (prevProps.listName !== this.props.listName || prevProps.defaultViewId !== this.props.defaultViewId) {
            this.logDiagnostic('Props changed; resetting selection and reloading rows. listName=' + String(this.props.listName || '(none)') + ', viewId=' + String(this.props.defaultViewId || '(none)'));
            this.setState({
                selectedViewId: this.props.defaultViewId || this.getInitialViewId(this.props.views),
                selectedItemId: 0,
                selectedMode: 'view',
                sortFieldName: '',
                sortDirection: '',
                activeFilterFieldName: '',
                draftFilterOperator: 'contains',
                draftFilterValue: '',
                columnFilters: {},
                currentPage: 0,
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
    ListControl.prototype.getInitialViewId = function (views) {
        for (var i = 0; i < views.length; i += 1) {
            if (views[i].isDefault) {
                return String(views[i].key);
            }
        }
        return views.length > 0 ? String(views[0].key) : '';
    };
    ListControl.prototype.handleExternalRefresh = function (event) {
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
    ListControl.prototype.getWebUrl = function () {
        return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    };
    ListControl.prototype.getJsonWithFallback = function (url) {
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
    ListControl.prototype.postJsonWithFallback = function (url, body) {
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
    ListControl.prototype.buildViewIdCandidates = function (selectedViewId) {
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
    ListControl.prototype.loadSelectedViewFieldNames = function (selectedViewId) {
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
    ListControl.prototype.getDisplayFields = function () {
        if (this.state.selectedViewId !== this.props.defaultViewId || !this.props.viewColumns || this.props.viewColumns.length === 0) {
            return this.state.fields;
        }
        var byName = {};
        for (var i = 0; i < this.state.fields.length; i += 1) {
            var field = this.state.fields[i];
            var fieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
            var responseName = String(field.Name || '').toLowerCase();
            if (fieldName) {
                byName[fieldName] = field;
            }
            if (responseName) {
                byName[responseName] = field;
            }
        }
        var configured = [];
        for (var j = 0; j < this.props.viewColumns.length; j += 1) {
            var column = this.props.viewColumns[j];
            var match = byName[String(column.fieldName || '').toLowerCase()];
            if (match) {
                configured.push(Object.assign({}, match, {
                    DisplayName: column.displayName || match.DisplayName || match.Name,
                    ConfiguredWidth: column.width || ''
                }));
            }
        }
        return configured;
    };
    ListControl.prototype.getConfiguredColumnStyle = function (field) {
        var width = parseInt(String(field.ConfiguredWidth || ''), 10);
        if (!isNaN(width) && width > 0) {
            return { width: width + 'px', minWidth: width + 'px', maxWidth: width + 'px' };
        }
        return {};
    };
    ListControl.prototype.getFieldsForConsumption = function (rawFields, viewFieldNames) {
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
    ListControl.prototype.loadListFieldTypeMap = function (viewFieldNames) {
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
    ListControl.prototype.loadListFieldTitleMap = function () {
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
    ListControl.prototype.applyFieldDisplayNames = function (fields, titleMap) {
        return fields.map(function (field) {
            var internalName = String(field.RealFieldName || field.Name || '');
            var displayName = titleMap[internalName.toLowerCase()] || field.DisplayName || internalName;
            return __assign({}, field, { DisplayName: displayName });
        });
    };
    ListControl.prototype.getRowFieldValue = function (row, field) {
        var fieldName = field.Name || field.RealFieldName || '';
        var value = row[fieldName];
        if ((value === undefined || value === null || value === '') && field.RealFieldName) {
            value = row[field.RealFieldName];
        }
        return value;
    };
    ListControl.prototype.stringifyCellValue = function (value) {
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
    ListControl.prototype.isMeaningfulCellValue = function (value) {
        if (value === undefined || value === null) {
            return false;
        }
        var text = this.stringifyCellValue(value).replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, '').trim();
        return text.length > 0;
    };
    ListControl.prototype.filterRenderableRows = function (rows, visibleFields) {
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
    ListControl.prototype.loadRowsFromItemsEndpoint = function (viewFieldNames) {
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
    ListControl.prototype.loadRows = function () {
        return __awaiter(this, void 0, void 0, function () {
            var baseEndpoint, selectedViewId, body, requestUrls, selectedRows, selectedFields, lastError, hadSuccessfulResponse, requestIndex, requestUrl, response, errorText, _readError_1, data, extracted, rows, fields, viewFieldNames, itemsFallback, visibleFields, fieldTitleMap, renderableRows, error_3, loadError;
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
                        _a.trys.push([1, 16, , 17]);
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
                        if (!(rows.length === 0)) return [3 /*break*/, 14];
                        return [4 /*yield*/, this.loadRowsFromItemsEndpoint(viewFieldNames)];
                    case 13:
                        itemsFallback = _a.sent();
                        rows = itemsFallback.rows;
                        if (fields.length === 0) {
                            fields = itemsFallback.fields;
                        }
                        _a.label = 14;
                    case 14:
                        visibleFields = this.getFieldsForConsumption(fields, viewFieldNames);
                        return [4 /*yield*/, this.loadListFieldTitleMap()];
                    case 15:
                        fieldTitleMap = _a.sent();
                        visibleFields = this.applyFieldDisplayNames(visibleFields, fieldTitleMap);
                        renderableRows = this.filterRenderableRows(rows, visibleFields);
                        this.setState({
                            fields: visibleFields,
                            rows: renderableRows,
                            loading: false,
                            error: null
                        });
                        this.logDiagnostic('loadRows completed. visibleFields=' + String(visibleFields.length) + ', renderableRows=' + String(renderableRows.length));
                        return [3 /*break*/, 17];
                    case 16:
                        error_3 = _a.sent();
                        loadError = error_3;
                        this.setState({
                            loading: false,
                            error: loadError && loadError.message ? loadError.message : 'Failed to load data.',
                            fields: [],
                            rows: []
                        });
                        this.logDiagnostic('loadRows failed: ' + (loadError && loadError.message ? loadError.message : String(loadError)));
                        return [3 /*break*/, 17];
                    case 17: return [2 /*return*/];
                }
            });
        });
    };
    ListControl.prototype.getRowItemId = function (row) {
        return toPositiveInt(row.ID || row.Id || row.id);
    };
    ListControl.prototype.selectRow = function (row) {
        var itemId = this.getRowItemId(row);
        this.setState({ selectedItemId: itemId, selectedMode: 'view' });
        this.logDiagnostic('Row selected. itemId=' + String(itemId));
        this.props.onSelectionChange(itemId, 'view');
    };
    ListControl.prototype.deleteSelected = function () {
        return __awaiter(this, void 0, void 0, function () {
            var endpoint, response, error_4, deleteError;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (this.state.selectedItemId <= 0) {
                            return [2 /*return*/];
                        }
                        if (typeof window !== 'undefined' && !window.confirm(strings.RuntimeDeleteConfirm)) {
                            return [2 /*return*/];
                        }
                        this.logDiagnostic('Deleting item. itemId=' + String(this.state.selectedItemId));
                        this.setState({ deleting: true, error: null });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items(" + this.state.selectedItemId + ")";
                        return [4 /*yield*/, this.props.context.spHttpClient.post(endpoint, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'DELETE'
                                }
                            })];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error(strings.RuntimeDeleteFailed);
                        }
                        this.setState({ selectedItemId: 0, selectedMode: 'view', deleting: false });
                        this.props.onSelectionChange(0, 'view');
                        this.logDiagnostic('Item deleted successfully. itemId=' + String(this.state.selectedItemId));
                        return [4 /*yield*/, this.loadRows()];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 5];
                    case 4:
                        error_4 = _a.sent();
                        deleteError = error_4;
                        this.setState({
                            deleting: false,
                            error: deleteError && deleteError.message ? deleteError.message : strings.RuntimeDeleteFailed
                        });
                        this.logDiagnostic('Delete failed: ' + (deleteError && deleteError.message ? deleteError.message : String(deleteError)));
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ListControl.prototype.logDiagnostic = function (message) {
        if (this.props.enableDiagnostics === false) {
            return;
        }
        console.log('[ListControl] ' + message);
    };
    ListControl.prototype.getCellMarkup = function (row, field) {
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
    ListControl.prototype.isTitleField = function (field) {
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
    ListControl.prototype.getCellPlainText = function (row, field) {
        var value = this.getRowFieldValue(row, field);
        if (value === undefined || value === null || value === '') {
            return '';
        }
        return this.stringifyCellValue(value).replace(/<[^>]*>/g, '').trim();
    };
    ListControl.prototype.getItemLinkUrl = function (row) {
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
    ListControl.prototype.getFieldKey = function (field) {
        return String(field.Name || field.RealFieldName || '').trim();
    };
    ListControl.prototype.getFilterOperatorOptions = function () {
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
    ListControl.prototype.toggleSort = function (field) {
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
    ListControl.prototype.openFilter = function (field) {
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
    ListControl.prototype.closeFilter = function () {
        this.setState({ activeFilterFieldName: '' });
    };
    ListControl.prototype.applyActiveFilter = function () {
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
    ListControl.prototype.clearActiveFilter = function () {
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
    ListControl.prototype.compareComparableValues = function (leftValue, rightValue) {
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
    ListControl.prototype.rowMatchesFilter = function (row, field, filter) {
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
    ListControl.prototype.parsePresetFilterConditions = function () {
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
    ListControl.prototype.resolveFieldByReference = function (fieldsByKey, fieldRef) {
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
    ListControl.prototype.rowMatchesPresetConditions = function (row, fieldsByKey, conditions) {
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
    ListControl.prototype.resolvePresetFilterValue = function (condition) {
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
    ListControl.prototype.getProcessedRows = function () {
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
    ListControl.prototype.parseConditionalStyleRules = function () {
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
    ListControl.prototype.toReactCssStyle = function (styleDefinition) {
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
    ListControl.prototype.evaluateConditionalStyleRule = function (row, rule, fieldsByKey) {
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
    ListControl.prototype.getConditionalStyleForRow = function (row, fieldsByKey, rules) {
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
    ListControl.prototype.getMatchedConditionalRuleNames = function (row, fieldsByKey, rules) {
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
    ListControl.prototype.render = function () {
        var _this = this;
        if (!this.props.listName) {
            return React.createElement("div", null, strings.RuntimeNoListSelected);
        }
        var canOperateOnSelection = this.state.selectedItemId > 0;
        var processedRows = this.getProcessedRows();
        var pageSize = this.props.pageSize > 0 ? this.props.pageSize : 0;
        var pageCount = pageSize > 0 ? Math.max(1, Math.ceil(processedRows.length / pageSize)) : 1;
        var currentPage = Math.min(this.state.currentPage, pageCount - 1);
        var visibleRows = pageSize > 0
            ? processedRows.slice(currentPage * pageSize, (currentPage + 1) * pageSize)
            : processedRows;
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
                this.props.showAdd && (React.createElement("button", { type: "button", onClick: function () {
                        _this.setState({ selectedItemId: 0, selectedMode: 'new' });
                        _this.props.onSelectionChange(0, 'new');
                    } }, strings.RuntimeNew)),
                this.props.showEdit && (React.createElement("button", { type: "button", disabled: !canOperateOnSelection, onClick: function () {
                        _this.setState({ selectedMode: 'edit' });
                        _this.props.onSelectionChange(_this.state.selectedItemId, 'edit');
                    } }, strings.RuntimeEdit)),
                this.props.showView && (React.createElement("button", { type: "button", disabled: !canOperateOnSelection, onClick: function () {
                        _this.setState({ selectedMode: 'view' });
                        _this.props.onSelectionChange(_this.state.selectedItemId, 'view');
                    } }, strings.RuntimeView)),
                this.props.showDelete && (React.createElement("button", { type: "button", disabled: !canOperateOnSelection || this.state.deleting, onClick: function () { return _this.deleteSelected(); } }, strings.RuntimeDelete))),
            this.props.isEditMode && (React.createElement("div", { className: "lc-status" },
                React.createElement("div", null, formatString(strings.RuntimeSelectedItem, this.state.selectedItemId || 0)),
                React.createElement("div", null, formatString(strings.RuntimeSelectedMode, this.state.selectedMode)),
                React.createElement("div", null, formatString(strings.RuntimeConditionalStyleMatches, selectedMatchedRuleNames.length > 0 ? selectedMatchedRuleNames.join(', ') : strings.RuntimeConditionalStyleMatchesNone)),
                !canOperateOnSelection && React.createElement("div", null, strings.RuntimeSelectRowPrompt))),
            this.state.loading && React.createElement("div", null, strings.RuntimeLoading),
            !!this.state.error && React.createElement("div", null, this.state.error),
            !this.state.loading && !this.state.error && processedRows.length === 0 && (React.createElement("div", null, hasActiveFilters ? strings.RuntimeNoItemsAfterFilter : strings.RuntimeNoItems)),
            !this.state.loading && !this.state.error && processedRows.length > 0 && (React.createElement("div", { className: "lc-table-wrap" },
                React.createElement("table", { className: "lc-table" },
                    React.createElement("thead", null,
                        React.createElement("tr", null, displayFields.map(function (field) {
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
                            return (React.createElement("th", { key: field.Name, style: _this.getConfiguredColumnStyle(field) },
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
                        }))),
                    React.createElement("tbody", null, visibleRows.map(function (row, index) {
                        var rowItemId = _this.getRowItemId(row);
                        var isSelected = rowItemId > 0 && rowItemId === _this.state.selectedItemId;
                        var conditionalStyle = _this.getConditionalStyleForRow(row, fieldsByKey, conditionalRules);
                        return (React.createElement("tr", { key: rowItemId > 0 ? String(rowItemId) : String(index), onClick: function () { return _this.selectRow(row); }, className: joinClassNames(['lc-row', isSelected ? 'lc-row-selected' : '']) }, displayFields.map(function (field) {
                            var markup = _this.getCellMarkup(row, field);
                            var showItemLink = _this.props.showLinkToItem && _this.isTitleField(field);
                            var itemLinkUrl = showItemLink ? _this.getItemLinkUrl(row) : '';
                            var itemLinkText = showItemLink ? _this.getCellPlainText(row, field) : '';
                            var cellFieldKey = _this.getFieldKey(field);
                            var columnStyle = conditionalStyle.columnStylesByFieldKey[cellFieldKey] || {};
                            var mergedCellStyle = mergeStyleObjects(mergeStyleObjects(conditionalStyle.rowStyle, columnStyle), _this.getConfiguredColumnStyle(field));
                            return (React.createElement("td", { key: field.Name, style: mergedCellStyle }, showItemLink && itemLinkUrl ? (React.createElement("a", { className: "lc-item-link", href: itemLinkUrl, onClick: function (ev) { return ev.stopPropagation(); } }, itemLinkText || strings.RuntimeView)) : (markup ? React.createElement("span", { dangerouslySetInnerHTML: markup }) : null)));
                        })));
                    }))))),
            !this.state.loading && !this.state.error && pageSize > 0 && pageCount > 1 && (React.createElement("div", { className: "lc-pagination" },
                React.createElement("button", { type: "button", disabled: currentPage === 0, onClick: function () { return _this.setState({ currentPage: Math.max(0, currentPage - 1) }); } }, strings.RuntimePreviousPage),
                React.createElement("span", null, strings.RuntimePageStatus.replace('{0}', String(currentPage + 1)).replace('{1}', String(pageCount))),
                React.createElement("button", { type: "button", disabled: currentPage >= pageCount - 1, onClick: function () { return _this.setState({ currentPage: Math.min(pageCount - 1, currentPage + 1) }); } }, strings.RuntimeNextPage)))));
    };
    return ListControl;
}(React.Component));
exports.ListControl = ListControl;

//# sourceMappingURL=ListControl.js.map
