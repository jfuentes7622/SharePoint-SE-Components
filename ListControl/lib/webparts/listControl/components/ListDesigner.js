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
require("./ListDesigner.css");
var SYSTEM_FIELDS = {
    ID: true,
    GUID: true,
    ContentType: true,
    AppAuthor: true,
    AppEditor: true,
    Edit: true,
    ItemChildCount: true,
    FolderChildCount: true,
    ComplianceAssetId: true
};
function createId() {
    return 'listfield_' + new Date().getTime().toString(36) + '_' + Math.floor(Math.random() * 100000).toString(36);
}
function escapeODataText(value) {
    return value.replace(/'/g, "''");
}
function toArray(value) {
    if (Array.isArray(value)) {
        return value;
    }
    if (value && Array.isArray(value.results)) {
        return value.results;
    }
    return [];
}
function trimGuidBraces(value) {
    return String(value || '').replace(/^[{]/, '').replace(/[}]$/, '');
}
function buildViewIdCandidates(selectedViewId) {
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
}
function parseGroupByFieldNames(viewQuery) {
    var queryText = String(viewQuery || '').trim();
    if (!queryText) {
        return { fieldNames: [], collapsed: false };
    }
    try {
        var xmlDocument = new DOMParser().parseFromString(/^<Query(?:\s|>)/i.test(queryText) ? queryText : '<Query>' + queryText + '</Query>', 'text/xml');
        if (xmlDocument.getElementsByTagName('parsererror').length > 0) {
            return { fieldNames: [], collapsed: false };
        }
        var groupByElements = xmlDocument.getElementsByTagName('GroupBy');
        if (groupByElements.length === 0) {
            return { fieldNames: [], collapsed: false };
        }
        var groupByElement = groupByElements[0];
        var fieldRefs = groupByElement.getElementsByTagName('FieldRef');
        var fieldNames = [];
        for (var i = 0; i < fieldRefs.length && fieldNames.length < 2; i += 1) {
            var name = fieldRefs[i].getAttribute('Name');
            if (name) {
                fieldNames.push(name);
            }
        }
        var collapsed = String(groupByElement.getAttribute('Collapse') || '').toUpperCase() === 'TRUE';
        return { fieldNames: fieldNames, collapsed: collapsed };
    }
    catch (_error) {
        return { fieldNames: [], collapsed: false };
    }
}
function copyFields(fields) {
    return JSON.parse(JSON.stringify(fields || []));
}
function parseGrouping(groupingJson) {
    var defaultGrouping = { field1: '', field2: '', collapsedByDefault: false, showCount: true };
    if (!String(groupingJson || '').trim()) {
        return defaultGrouping;
    }
    try {
        var grouping = JSON.parse(groupingJson);
        if (!grouping) {
            return defaultGrouping;
        }
        return {
            field1: String(grouping.field1 || ''),
            field2: String(grouping.field2 || ''),
            collapsedByDefault: grouping.collapsedByDefault === true,
            showCount: grouping.showCount !== false
        };
    }
    catch (_error) {
        return defaultGrouping;
    }
}
var ListDesigner = (function (_super) {
    __extends(ListDesigner, _super);
    function ListDesigner(props) {
        var _this = _super.call(this, props) || this;
        var initialFields = (props.viewColumns || []).filter(function (column) {
            return !!(column && column.fieldName);
        }).map(function (column) {
            return {
                id: createId(),
                fieldName: column.fieldName,
                displayName: column.displayName || column.fieldName,
                width: column.width || ''
            };
        });
        _this.state = {
            fields: initialFields,
            sharePointFields: [],
            selectedFieldId: initialFields.length > 0 ? initialFields[0].id : '',
            loading: false,
            error: '',
            grouping: parseGrouping(props.groupingJson),
            viewLoading: false,
            viewLoadMessage: ''
        };
        return _this;
    }
    ListDesigner.prototype.componentDidMount = function () {
        this.loadFields();
    };
    ListDesigner.prototype.getWebUrl = function () {
        return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    };
    ListDesigner.prototype.getFieldsResponse = function (url) {
        return __awaiter(this, void 0, void 0, function () {
            var acceptHeaders, response, i;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        acceptHeaders = [
                            '',
                            'application/json;odata=verbose',
                            'application/json;odata=minimalmetadata',
                            'application/json;odata=nometadata'
                        ];
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!(i < acceptHeaders.length)) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.props.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, acceptHeaders[i] ? { headers: { Accept: acceptHeaders[i] } } : undefined)];
                    case 2:
                        response = _a.sent();
                        if (response.ok) {
                            return [2 /*return*/, response];
                        }
                        _a.label = 3;
                    case 3:
                        i += 1;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/, response];
                }
            });
        });
    };
    ListDesigner.prototype.loadFields = function () {
        return __awaiter(this, void 0, void 0, function () {
            var url, response, data, sourceFields, fields, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.props.listName) {
                            this.setState({ error: 'Select a SharePoint list before opening List Designer.' });
                            return [2 /*return*/];
                        }
                        this.setState({ loading: true, error: '' });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        url = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,Title,Hidden,FromBaseType";
                        return [4 /*yield*/, this.getFieldsResponse(url)];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error('SharePoint fields could not be loaded. HTTP ' + String(response.status) + ' ' + String(response.statusText || ''));
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        sourceFields = toArray(data && data.value).length > 0 ? toArray(data.value) : toArray(data && data.d && data.d.results);
                        fields = sourceFields.filter(function (field) {
                            var isAttachment = field.InternalName === 'Attachments';
                            var isCommonSystemField = field.InternalName === 'Author' || field.InternalName === 'Editor'
                                || field.InternalName === 'Created' || field.InternalName === 'Modified';
                            return !field.Hidden && !SYSTEM_FIELDS[field.InternalName]
                                && (!field.FromBaseType || field.InternalName === 'Title' || isAttachment || isCommonSystemField);
                        }).map(function (field) {
                            return {
                                internalName: String(field.InternalName || ''),
                                title: String(field.Title || field.InternalName || '')
                            };
                        });
                        this.setState({ sharePointFields: fields, loading: false });
                        return [3 /*break*/, 5];
                    case 4:
                        error_1 = _a.sent();
                        this.setState({ loading: false, error: error_1 && error_1.message ? error_1.message : 'SharePoint fields could not be loaded.' });
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    ListDesigner.prototype.loadFromView = function () {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, listPath, viewIds, viewQuery, viewFieldNames, i, encoded, normalized, queryUrls, q, queryResponse, queryData, queryContainer, fieldsUrls, f, fieldsResponse, fieldsData, names, groupByResult, nextFields, addedFieldNames, v, viewFieldName, matchedSharePointField, s, groupField1, groupField2, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.props.listName || !this.props.viewId) {
                            this.setState({ viewLoadMessage: 'Select a SharePoint view in the web part properties first.' });
                            return [2 /*return*/];
                        }
                        if (this.state.fields.length > 0 && typeof window !== 'undefined'
                            && !window.confirm('Replace the current columns and grouping with the selected view\'s fields and grouping?')) {
                            return [2 /*return*/];
                        }
                        this.setState({ viewLoading: true, viewLoadMessage: '' });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 14, , 15]);
                        webUrl = this.getWebUrl();
                        listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
                        viewIds = buildViewIdCandidates(this.props.viewId);
                        viewQuery = '';
                        viewFieldNames = [];
                        i = 0;
                        _a.label = 2;
                    case 2:
                        if (!(i < viewIds.length && !viewQuery)) return [3 /*break*/, 13];
                        encoded = encodeURIComponent(viewIds[i]);
                        normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
                        queryUrls = [
                            webUrl + listPath + "/views/getById('" + encoded + "')?$select=ViewQuery",
                            webUrl + listPath + "/views(guid'" + normalized + "')?$select=ViewQuery"
                        ];
                        q = 0;
                        _a.label = 3;
                    case 3:
                        if (!(q < queryUrls.length && !viewQuery)) return [3 /*break*/, 7];
                        return [4 /*yield*/, this.getFieldsResponse(queryUrls[q])];
                    case 4:
                        queryResponse = _a.sent();
                        if (!queryResponse.ok) {
                            return [3 /*break*/, 6];
                        }
                        return [4 /*yield*/, queryResponse.json()];
                    case 5:
                        queryData = _a.sent();
                        queryContainer = queryData && queryData.d ? queryData.d : queryData;
                        if (queryContainer && queryContainer.ViewQuery !== undefined && queryContainer.ViewQuery !== null) {
                            viewQuery = String(queryContainer.ViewQuery);
                        }
                        _a.label = 6;
                    case 6:
                        q += 1;
                        return [3 /*break*/, 3];
                    case 7:
                        fieldsUrls = [
                            webUrl + listPath + "/views/getById('" + encoded + "')/ViewFields",
                            webUrl + listPath + "/views(guid'" + normalized + "')/ViewFields"
                        ];
                        f = 0;
                        _a.label = 8;
                    case 8:
                        if (!(f < fieldsUrls.length && viewFieldNames.length === 0)) return [3 /*break*/, 12];
                        return [4 /*yield*/, this.getFieldsResponse(fieldsUrls[f])];
                    case 9:
                        fieldsResponse = _a.sent();
                        if (!fieldsResponse.ok) {
                            return [3 /*break*/, 11];
                        }
                        return [4 /*yield*/, fieldsResponse.json()];
                    case 10:
                        fieldsData = _a.sent();
                        names = toArray(fieldsData.value);
                        if (names.length === 0) {
                            names = toArray(fieldsData.Items);
                        }
                        if (names.length === 0) {
                            names = toArray(fieldsData && fieldsData.d && fieldsData.d.Items);
                        }
                        if (names.length > 0) {
                            viewFieldNames = names.map(function (name) { return String(name || ''); }).filter(function (name) { return !!name; });
                        }
                        _a.label = 11;
                    case 11:
                        f += 1;
                        return [3 /*break*/, 8];
                    case 12:
                        i += 1;
                        return [3 /*break*/, 2];
                    case 13:
                        groupByResult = parseGroupByFieldNames(viewQuery);
                        if (viewFieldNames.length > 0) {
                            nextFields = [];
                            addedFieldNames = {};
                            for (v = 0; v < viewFieldNames.length; v += 1) {
                                viewFieldName = viewFieldNames[v];
                                if (SYSTEM_FIELDS[viewFieldName]) {
                                    continue;
                                }
                                for (s = 0; s < this.state.sharePointFields.length; s += 1) {
                                    if (this.state.sharePointFields[s].internalName.toLowerCase() === viewFieldName.toLowerCase()) {
                                        matchedSharePointField = this.state.sharePointFields[s];
                                        break;
                                    }
                                }
                                if (!matchedSharePointField) {
                                    continue;
                                }
                                // A SharePoint view can list the same field more than once; only add it once.
                                if (addedFieldNames[matchedSharePointField.internalName.toLowerCase()]) {
                                    continue;
                                }
                                addedFieldNames[matchedSharePointField.internalName.toLowerCase()] = true;
                                nextFields.push({
                                    id: createId(),
                                    fieldName: matchedSharePointField.internalName,
                                    displayName: matchedSharePointField.title,
                                    width: ''
                                });
                            }
                            if (nextFields.length > 0) {
                                this.setState({ fields: nextFields, selectedFieldId: nextFields[0].id });
                            }
                        }
                        groupField1 = groupByResult.fieldNames.length > 0 ? groupByResult.fieldNames[0] : '';
                        groupField2 = groupByResult.fieldNames.length > 1 ? groupByResult.fieldNames[1] : '';
                        this.setState({
                            viewLoading: false,
                            grouping: {
                                field1: groupField1,
                                field2: groupField2,
                                collapsedByDefault: groupByResult.collapsed,
                                showCount: this.state.grouping.showCount
                            },
                            viewLoadMessage: viewFieldNames.length > 0
                                ? (groupField1 ? 'Loaded columns and grouping from the selected view.' : 'Loaded columns from the selected view. The view has no grouping configured.')
                                : (groupField1 ? 'Loaded grouping from the selected view.' : 'The selected view has no fields or grouping to load.')
                        });
                        return [3 /*break*/, 15];
                    case 14:
                        error_2 = _a.sent();
                        this.setState({ viewLoading: false, viewLoadMessage: error_2 && error_2.message ? error_2.message : 'Failed to load the selected view.' });
                        return [3 /*break*/, 15];
                    case 15: return [2 /*return*/];
                }
            });
        });
    };
    ListDesigner.prototype.getAvailableFields = function () {
        var used = {};
        this.state.fields.forEach(function (field) { used[String(field.fieldName || '').toLowerCase()] = true; });
        return this.state.sharePointFields.filter(function (field) { return !used[field.internalName.toLowerCase()]; });
    };
    ListDesigner.prototype.addField = function (source) {
        var field = {
            id: createId(),
            fieldName: source.internalName,
            displayName: source.title,
            width: ''
        };
        var fields = copyFields(this.state.fields);
        fields.push(field);
        this.setState({ fields: fields, selectedFieldId: field.id });
    };
    ListDesigner.prototype.getSelectedField = function () {
        for (var i = 0; i < this.state.fields.length; i += 1) {
            if (this.state.fields[i].id === this.state.selectedFieldId) {
                return this.state.fields[i];
            }
        }
        return undefined;
    };
    ListDesigner.prototype.updateSelectedField = function (mutator) {
        var fields = copyFields(this.state.fields);
        for (var i = 0; i < fields.length; i += 1) {
            if (fields[i].id === this.state.selectedFieldId) {
                mutator(fields[i]);
                break;
            }
        }
        this.setState({ fields: fields });
    };
    ListDesigner.prototype.moveField = function (fieldId, direction) {
        var fields = copyFields(this.state.fields);
        var index = -1;
        for (var i = 0; i < fields.length; i += 1) {
            if (fields[i].id === fieldId) {
                index = i;
                break;
            }
        }
        var target = index + direction;
        if (index < 0 || target < 0 || target >= fields.length) {
            return;
        }
        var moved = fields[index];
        fields[index] = fields[target];
        fields[target] = moved;
        this.setState({ fields: fields });
    };
    ListDesigner.prototype.removeField = function (fieldId) {
        var fields = this.state.fields.filter(function (field) { return field.id !== fieldId; });
        this.setState({
            fields: fields,
            selectedFieldId: this.state.selectedFieldId === fieldId ? (fields.length > 0 ? fields[0].id : '') : this.state.selectedFieldId
        });
    };
    ListDesigner.prototype.updateGrouping = function (mutator) {
        var grouping = {
            field1: this.state.grouping.field1,
            field2: this.state.grouping.field2,
            collapsedByDefault: this.state.grouping.collapsedByDefault,
            showCount: this.state.grouping.showCount
        };
        mutator(grouping);
        if (!grouping.field1) {
            grouping.field2 = '';
        }
        this.setState({ grouping: grouping });
    };
    ListDesigner.prototype.getFieldLabel = function (fieldName) {
        if (!fieldName) {
            return '';
        }
        for (var i = 0; i < this.state.fields.length; i += 1) {
            if (this.state.fields[i].fieldName.toLowerCase() === fieldName.toLowerCase()) {
                return this.state.fields[i].displayName || this.state.fields[i].fieldName;
            }
        }
        for (var j = 0; j < this.state.sharePointFields.length; j += 1) {
            if (this.state.sharePointFields[j].internalName.toLowerCase() === fieldName.toLowerCase()) {
                return this.state.sharePointFields[j].title || fieldName;
            }
        }
        return fieldName;
    };
    ListDesigner.prototype.save = function () {
        if (this.state.loading) {
            return;
        }
        var columns = this.state.fields.map(function (field) {
            return { fieldName: field.fieldName, displayName: field.displayName, width: field.width || '' };
        });
        var grouping = this.state.grouping;
        var groupingJson = JSON.stringify({
            enabled: !!grouping.field1,
            field1: grouping.field1,
            field2: grouping.field1 ? grouping.field2 : '',
            collapsedByDefault: grouping.collapsedByDefault,
            showCount: grouping.showCount
        });
        this.props.onSave(columns, groupingJson);
    };
    ListDesigner.prototype.renderPalette = function () {
        var _this = this;
        var available = this.getAvailableFields();
        return (React.createElement("aside", { className: "gd-panel gd-palette" },
            React.createElement("h3", null, "SharePoint fields"),
            React.createElement("p", null, "Add fields as list columns."),
            this.state.loading && React.createElement("div", { className: "gd-empty" }, "Loading fields..."),
            this.state.error && React.createElement("div", { className: "gd-error" }, this.state.error),
            !this.state.loading && !this.state.error && available.length === 0 && React.createElement("div", { className: "gd-empty" }, "All available fields are in the list."),
            available.map(function (field) { return (React.createElement("button", { key: field.internalName, type: "button", className: "gd-palette-item", onClick: function () { return _this.addField(field); } },
                React.createElement("span", null, field.title),
                React.createElement("small", null, field.internalName))); })));
    };
    ListDesigner.prototype.renderGroupingSettings = function () {
        var _this = this;
        var grouping = this.state.grouping;
        return (React.createElement("div", { className: "gd-group-settings" },
            React.createElement("div", { className: "gd-canvas-header" },
                React.createElement("div", null,
                    React.createElement("h3", null, "Grouping"),
                    React.createElement("p", null, "Group list rows by up to two columns, similar to a SharePoint view's Group By."))),
            this.props.viewId && (React.createElement("div", { className: "gd-group-view-load" },
                React.createElement("button", { type: "button", disabled: this.state.viewLoading, onClick: function () { return _this.loadFromView(); } }, this.state.viewLoading ? 'Loading from view…' : 'Load columns & grouping from selected view'),
                this.state.viewLoadMessage && React.createElement("div", { className: "gd-hint" }, this.state.viewLoadMessage))),
            React.createElement("div", { className: "gd-inline-fields gd-inline-fields-two" },
                React.createElement("label", null,
                    "Group by",
                    React.createElement("select", { value: grouping.field1, onChange: function (ev) { var value = ev.currentTarget.value; _this.updateGrouping(function (next) { next.field1 = value; }); } },
                        React.createElement("option", { value: "" }, "None"),
                        this.state.fields.map(function (field) { return React.createElement("option", { key: field.id, value: field.fieldName }, field.displayName || field.fieldName); }))),
                React.createElement("label", null,
                    "Then by",
                    React.createElement("select", { value: grouping.field2, disabled: !grouping.field1, onChange: function (ev) { var value = ev.currentTarget.value; _this.updateGrouping(function (next) { next.field2 = value; }); } },
                        React.createElement("option", { value: "" }, "None"),
                        this.state.fields.filter(function (field) { return field.fieldName !== grouping.field1; }).map(function (field) { return React.createElement("option", { key: field.id, value: field.fieldName }, field.displayName || field.fieldName); })))),
            React.createElement("label", { className: "gd-check" },
                React.createElement("input", { type: "checkbox", checked: grouping.collapsedByDefault, disabled: !grouping.field1, onChange: function (ev) { var checked = ev.currentTarget.checked; _this.updateGrouping(function (next) { next.collapsedByDefault = checked; }); } }),
                " Collapse groups by default"),
            React.createElement("label", { className: "gd-check" },
                React.createElement("input", { type: "checkbox", checked: grouping.showCount, disabled: !grouping.field1, onChange: function (ev) { var checked = ev.currentTarget.checked; _this.updateGrouping(function (next) { next.showCount = checked; }); } }),
                " Show item count per group")));
    };
    ListDesigner.prototype.renderColumnsList = function () {
        var _this = this;
        return (React.createElement("div", null,
            this.state.fields.length === 0 && React.createElement("div", { className: "gd-empty gd-empty-canvas" }, "Add SharePoint fields from the left panel."),
            this.state.fields.map(function (field, index) { return (React.createElement("div", { key: field.id, className: 'gd-column ' + (field.id === _this.state.selectedFieldId ? 'gd-column-selected' : ''), onClick: function () { return _this.setState({ selectedFieldId: field.id }); } },
                React.createElement("div", { className: "gd-column-order" }, index + 1),
                React.createElement("div", { className: "gd-column-main" },
                    React.createElement("strong", null, field.displayName || field.fieldName),
                    React.createElement("span", null, field.fieldName)),
                React.createElement("div", { className: "gd-column-actions" },
                    React.createElement("button", { type: "button", title: "Move up", disabled: index === 0, onClick: function (ev) { ev.stopPropagation(); _this.moveField(field.id, -1); } }, "\u2191"),
                    React.createElement("button", { type: "button", title: "Move down", disabled: index === _this.state.fields.length - 1, onClick: function (ev) { ev.stopPropagation(); _this.moveField(field.id, 1); } }, "\u2193"),
                    React.createElement("button", { type: "button", title: "Remove column", onClick: function (ev) { ev.stopPropagation(); _this.removeField(field.id); } }, "\u00D7")))); })));
    };
    ListDesigner.prototype.renderCanvas = function () {
        var grouping = this.state.grouping;
        var columnsPreview = this.renderColumnsList();
        if (grouping.field1) {
            var innerPreview = grouping.field2 ? (React.createElement("div", { className: "gd-group-container gd-group-container-nested" },
                React.createElement("div", { className: "gd-group-container-header" },
                    React.createElement("span", { className: "gd-group-container-icon" }, "\u25BE"),
                    React.createElement("span", null,
                        "Then by: ",
                        this.getFieldLabel(grouping.field2),
                        " = ",
                        React.createElement("em", null, "(example value)")),
                    grouping.showCount && React.createElement("span", { className: "gd-group-container-count" }, "3 items")),
                React.createElement("div", { className: "gd-group-container-body" }, columnsPreview))) : columnsPreview;
            columnsPreview = (React.createElement("div", { className: "gd-group-container" },
                React.createElement("div", { className: "gd-group-container-header" },
                    React.createElement("span", { className: "gd-group-container-icon" }, grouping.collapsedByDefault ? '▸' : '▾'),
                    React.createElement("span", null,
                        "Grouped by: ",
                        this.getFieldLabel(grouping.field1),
                        " = ",
                        React.createElement("em", null, "(example value)")),
                    grouping.showCount && React.createElement("span", { className: "gd-group-container-count" }, grouping.field2 ? '' : '5 items')),
                React.createElement("div", { className: "gd-group-container-body" }, innerPreview)));
        }
        return (React.createElement("main", { className: "gd-canvas" },
            this.renderGroupingSettings(),
            React.createElement("div", { className: "gd-canvas-header" },
                React.createElement("div", null,
                    React.createElement("h3", null, "List columns"),
                    React.createElement("p", null, "Column order follows this list from top to bottom.")),
                React.createElement("span", null,
                    this.state.fields.length,
                    " columns")),
            columnsPreview));
    };
    ListDesigner.prototype.renderFieldEditor = function () {
        var _this = this;
        var field = this.getSelectedField();
        if (!field) {
            return React.createElement("aside", { className: "gd-panel gd-properties" },
                React.createElement("div", { className: "gd-empty" }, "Select a list column to configure it."));
        }
        return (React.createElement("aside", { className: "gd-panel gd-properties" },
            React.createElement("h3", null, "Column settings"),
            React.createElement("label", null,
                "Column label",
                React.createElement("input", { type: "text", value: field.displayName || '', onChange: function (ev) { return _this.updateSelectedField(function (next) { next.displayName = ev.currentTarget.value; }); } })),
            React.createElement("label", null,
                "SharePoint field",
                React.createElement("div", { className: "gd-readonly-value" }, field.fieldName)),
            React.createElement("label", null,
                "Column width",
                React.createElement("input", { type: "text", value: field.width || '', placeholder: "180px or 20%", onChange: function (ev) { return _this.updateSelectedField(function (next) { next.width = ev.currentTarget.value; }); } }))));
    };
    ListDesigner.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: "gd-overlay" },
            React.createElement("header", { className: "gd-toolbar" },
                React.createElement("div", null,
                    React.createElement("h2", null, "List Designer"),
                    React.createElement("span", null, this.props.listName || 'No list selected')),
                React.createElement("div", { className: "gd-toolbar-actions" },
                    React.createElement("button", { type: "button", onClick: this.props.onCancel }, "Cancel"),
                    React.createElement("button", { type: "button", className: "gd-primary", onClick: function () { return _this.save(); }, disabled: !this.props.listName || this.state.loading }, "Save list"))),
            React.createElement("div", { className: "gd-workspace" },
                this.renderPalette(),
                this.renderCanvas(),
                this.renderFieldEditor())));
    };
    return ListDesigner;
}(React.Component));
exports.ListDesigner = ListDesigner;

//# sourceMappingURL=ListDesigner.js.map
