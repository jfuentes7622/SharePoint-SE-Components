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
var ReactDom = require("react-dom");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_http_1 = require("@microsoft/sp-http");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var PropertyPaneCustomField_1 = require("@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneCustomField/PropertyPaneCustomField");
var PropertyFieldColorPicker_1 = require("@pnp/spfx-property-controls/lib/PropertyFieldColorPicker");
var strings = require("GridControlWebPartStrings");
var GridControl_1 = require("./components/GridControl");
var GridDesigner_1 = require("./components/GridDesigner");
var deterministicFullWidth_1 = require("../shared/deterministicFullWidth");
var packageSolutionConfig = require('../../../config/package-solution.json');
function escapeODataText(value) {
    return value.replace(/'/g, "''");
}
function normalizeQueryParamName(value, fallback) {
    var normalized = String(value || '').trim();
    if (!normalized || normalized.toLowerCase() === 'id') {
        return fallback;
    }
    return normalized;
}
function normalizePageSize(value) {
    var parsed = parseInt(String(value === undefined || value === null ? '' : value), 10);
    return !isNaN(parsed) && parsed > 0 ? parsed : 0;
}
function normalizeStringCollection(value) {
    var source = value && value.results ? value.results : value;
    if (!Array.isArray(source)) {
        return [];
    }
    return source.map(function (item) {
        return String(item && item.Name !== undefined ? item.Name : item || '');
    }).filter(function (item) {
        return !!item;
    });
}
var GridControlWebPart = (function (_super) {
    __extends(GridControlWebPart, _super);
    function GridControlWebPart() {
        var _this = _super.call(this) || this;
        _this._lists = [];
        _this._sitePages = [];
        _this._views = [];
        _this._listFields = [];
        _this._selectedItemId = 0;
        _this._selectedMode = 'view';
        _this._filterJsonValidationMessage = '';
        _this._filterDesignerMessage = '';
        _this._showFilterExpressionHelp = false;
        _this._conditionalStyleJsonValidationMessage = '';
        _this._conditionalStyleDesignerMessage = '';
        _this._showConditionalStyleExpressionHelp = false;
        _this._isEditingConditionalStyle = false;
        _this._conditionalStyleDesignerRevision = 0;
        _this._fieldTypeByInternalName = {};
        _this._fieldChoicesByInternalName = {};
        _this._fieldLookupListByInternalName = {};
        _this._filterLookupItemOptions = [];
        _this._conditionalStyleLookupItemOptions = [];
        _this._filterLookupMessage = '';
        _this._conditionalStyleLookupMessage = '';
        _this._isGridDesignerOpen = false;
        _this.onPropertyPaneFieldChanged = _this.onPropertyPaneFieldChanged.bind(_this);
        return _this;
    }
    Object.defineProperty(GridControlWebPart.prototype, "id", {
        get: function () {
            return this.context.instanceId;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(GridControlWebPart.prototype, "metadata", {
        get: function () {
            var instanceName = String(this.properties.instanceName || '').trim();
            return {
                title: instanceName || strings.DynamicSourceTitle,
                description: strings.DynamicSourceDescription,
                alias: this.context.manifest.alias,
                componentId: this.context.manifest.id,
                instanceId: this.context.instanceId,
            };
        },
        enumerable: true,
        configurable: true
    });
    GridControlWebPart.prototype.getPropertyDefinitions = function () {
        return [
            {
                id: 'instanceName',
                title: strings.PropInstanceNameLabel,
                description: strings.PropInstanceNameDescription,
            },
            {
                id: 'listName',
                title: 'Configured list name',
                description: 'The SharePoint list configured for this Grid Control.',
            },
            {
                id: 'selectedItemId',
                title: strings.DynamicPropertySelectedItemIdTitle,
                description: strings.DynamicPropertySelectedItemIdDescription,
            },
            {
                id: 'selectedMode',
                title: strings.DynamicPropertySelectedModeTitle,
                description: strings.DynamicPropertySelectedModeDescription,
            }
        ];
    };
    GridControlWebPart.prototype.getPropertyValue = function (propertyId) {
        if (propertyId === 'instanceName') {
            return String(this.properties.instanceName || '').trim();
        }
        if (propertyId === 'listName') {
            return String(this.properties.listName || '').trim();
        }
        if (propertyId === 'selectedItemId') {
            return this._selectedItemId;
        }
        if (propertyId === 'selectedMode') {
            return this._selectedMode;
        }
        throw new Error('Bad property id');
    };
    GridControlWebPart.prototype.render = function () {
        var _this = this;
        deterministicFullWidth_1.updateResponsiveOptionalFullWidth(this.domElement, this.context.instanceId, this.properties.forceFullWidth === true);
        if (this._isGridDesignerOpen) {
            var designerElement = React.createElement(GridDesigner_1.GridDesigner, {
                context: this.context,
                listName: this.properties.listName || '',
                schemaJson: this.properties.gridSchemaJson || '',
                onSave: function (schemaJson) { return _this.saveGridDesign(schemaJson); },
                onCancel: function () { return _this.closeGridDesigner(); }
            });
            ReactDom.render(designerElement, this.domElement);
            return;
        }
        var element = React.createElement(GridControl_1.GridControl, {
            context: this.context,
            listName: this.properties.listName || '',
            defaultViewId: this.properties.viewId || '',
            views: this._views,
            viewColumns: this.properties.viewColumns || [],
            gridSchemaJson: this.properties.gridSchemaJson || '',
            pageSize: normalizePageSize(this.properties.pageSize),
            isEditMode: this.displayMode === sp_core_library_1.DisplayMode.Edit,
            showViewSelector: this.properties.showViewSelector !== false,
            showRefresh: this.properties.showRefresh !== false,
            showAdd: this.properties.showAdd !== false,
            showDelete: this.properties.showDelete !== false,
            showLinkToItem: this.properties.showLinkToItem === true,
            linkTargetPageUrl: this.properties.linkTargetPageUrl === '__defaultForm__'
                ? '' : (this.properties.linkTargetPageUrl || ''),
            linkTargetIdParam: this.properties.linkTargetIdParam || 'itemid',
            includeReturnUrlParam: this.properties.includeReturnUrlParam === true,
            enableDiagnostics: this.properties.enableDiagnostics !== false,
            bodyTextColor: this.properties.bodyTextColor || '',
            bodyFontFamily: this.properties.bodyFontFamily || '',
            bodyFontSize: this.properties.bodyFontSize || '14px',
            bodyFontStyle: this.properties.bodyFontStyle || 'normal',
            bodyFontBold: this.properties.bodyFontBold === true,
            bodyTextAlign: this.properties.bodyTextAlign || 'left',
            dateDisplayFormat: this.properties.dateDisplayFormat || 'mdy',
            timeDisplayFormat: this.properties.timeDisplayFormat || '24hour',
            selectedTextColor: this.properties.selectedTextColor || '',
            selectedBackgroundColor: this.properties.selectedBackgroundColor || '',
            selectedFontStyle: this.properties.selectedFontStyle || 'normal',
            selectedFontBold: this.properties.selectedFontBold === true,
            headerTextColor: this.properties.headerTextColor || '',
            headerBackgroundColor: this.properties.headerBackgroundColor || '',
            headerFontFamily: this.properties.headerFontFamily || '',
            headerFontSize: this.properties.headerFontSize || '14px',
            headerFontStyle: this.properties.headerFontStyle || 'normal',
            headerFontBold: this.properties.headerFontBold === true,
            headerTextAlign: this.properties.headerTextAlign || 'left',
            tableBackgroundColor: this.properties.tableBackgroundColor || '#ffffff',
            tableBorderColor: this.properties.tableBorderColor || '#cccccc',
            tableBorderWidth: this.properties.tableBorderWidth || '1px',
            tableCornerStyle: this.properties.tableCornerStyle || 'square',
            tableCornerRadius: typeof this.properties.tableCornerRadius === 'number' ? this.properties.tableCornerRadius : 0,
            tableRowLineWidth: typeof this.properties.tableRowLineWidth === 'number' ? this.properties.tableRowLineWidth : 1,
            cornerStyle: this.properties.cornerStyle || 'square',
            cornerRadius: typeof this.properties.cornerRadius === 'number' ? this.properties.cornerRadius : 12,
            alternateRowShading: this.properties.alternateRowShading === true,
            alternateRowShadingColor: this.properties.alternateRowShadingColor || '#f9f9f9',
            buttonTextColor: this.properties.buttonTextColor || '#000000',
            buttonBackgroundColor: this.properties.buttonBackgroundColor || '#f0f0f0',
            buttonFontFamily: this.properties.buttonFontFamily || 'inherit',
            buttonFontSize: this.properties.buttonFontSize || '14px',
            buttonFontStyle: this.properties.buttonFontStyle || 'normal',
            buttonFontBold: this.properties.buttonFontBold === true,
            buttonCornerStyle: this.properties.buttonCornerStyle || 'square',
            buttonCornerRadius: typeof this.properties.buttonCornerRadius === 'number' ? this.properties.buttonCornerRadius : 4,
            webpartBackgroundColor: this.properties.webpartBackgroundColor || 'transparent',
            webpartBorderColor: this.properties.webpartBorderColor || '#ccc',
            webpartBorderWidth: this.properties.webpartBorderWidth || 0,
            filterJson: this.properties.filterJson || '',
            conditionalStyleJson: this.properties.conditionalStyleJson || '',
            onSelectionChange: function (itemId, mode) { return _this.handleSelectionChange(itemId, mode); }
        });
        ReactDom.render(element, this.domElement);
    };
    GridControlWebPart.prototype.onInit = function () {
        var _this = this;
        this.logDiagnostic('onInit started. listName=' + String(this.properties.listName || '(none)'));
        this.properties.linkTargetIdParam = normalizeQueryParamName(this.properties.linkTargetIdParam, 'itemid');
        this.initializeDynamicDataSource();
        return Promise.all([this.loadLists(), this.loadSitePages()]).then(function () {
            _this.logDiagnostic('List metadata loaded. Count=' + String(_this._lists.length));
            if (_this.properties.listName) {
                _this.logDiagnostic('Loading views for configured list: ' + String(_this.properties.listName));
                return _this.loadViews(_this.properties.listName).then(function () {
                    return _this.loadListFields(_this.properties.listName).then(function () {
                        if ((!_this.properties.viewColumns || _this.properties.viewColumns.length === 0) && _this.properties.viewId) {
                            return _this.loadViewColumns(_this.properties.listName, _this.properties.viewId);
                        }
                        return Promise.resolve();
                    });
                });
            }
            return Promise.resolve();
        });
    };
    GridControlWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        this.logDiagnostic('Property pane opened. listName=' + String(this.properties.listName || '(none)'));
        if (this._lists.length === 0) {
            this.loadLists();
        }
        if (this._sitePages.length === 0) {
            this.loadSitePages();
        }
        if (this.properties.listName && this._views.length === 0) {
            this.loadViews(this.properties.listName);
        }
        if (this.properties.listName && this._listFields.length === 0) {
            this.loadListFields(this.properties.listName);
        }
    };
    GridControlWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath, oldValue, newValue) {
        this.logDiagnostic('Property changed: ' + propertyPath + ', old=' + String(oldValue) + ', new=' + String(newValue));
        if (propertyPath === 'instanceName' && oldValue !== newValue) {
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
            this.notifyDynamicData('instanceName');
            this.notifyDynamicSourceChanged();
            return;
        }
        if (propertyPath === 'listName' && oldValue !== newValue) {
            this.properties.viewId = '';
            this.properties.viewColumns = [];
            this.properties.gridSchemaJson = '';
            this.properties.filterDesignerField = '';
            this.properties.conditionalStyleApplyField = '';
            this.properties.conditionalStyleConditionField = '';
            this._views = [];
            this._listFields = [];
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
            if (newValue) {
                this.loadViews(String(newValue));
                this.loadListFields(String(newValue));
            }
            this.notifyDynamicData('listName');
            this.notifyDynamicSourceChanged();
            return;
        }
        if (propertyPath === 'viewId' && oldValue !== newValue) {
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
            this.context.propertyPane.refresh();
            this.render();
            return;
        }
        if (propertyPath === 'filterDesignerField'
            || propertyPath === 'filterDesignerOperator'
            || propertyPath === 'filterDesignerLogical'
            || propertyPath === 'filterDesignerValueType'
            || propertyPath === 'filterDesignerSelectedIndex'
            || propertyPath === 'filterDesignerLookupPick'
            || propertyPath === 'conditionalStyleScope'
            || propertyPath === 'conditionalStyleApplyField'
            || propertyPath === 'conditionalStyleConditionField'
            || propertyPath === 'conditionalStyleConditionOperator'
            || propertyPath === 'conditionalStyleConditionLogical'
            || propertyPath === 'conditionalStyleConditionValueType'
            || propertyPath === 'conditionalStyleEnabled'
            || propertyPath === 'conditionalStylePriority'
            || propertyPath === 'conditionalStyleSelectedIndex'
            || propertyPath === 'conditionalStyleLookupPick') {
            if (propertyPath === 'filterDesignerField') {
                this.properties.filterDesignerValue = '';
                this._filterLookupItemOptions = [];
                this.properties.filterDesignerLookupPick = '';
                this._filterLookupMessage = '';
                if (this.isLookupTypeField(String(newValue || ''))) {
                    this.handleLoadFilterLookupItems();
                }
            }
            if (propertyPath === 'filterDesignerValueType') {
                this.properties.filterDesignerValue = '';
                this.properties.filterDesignerLookupPick = '';
            }
            if (propertyPath === 'conditionalStyleConditionField') {
                this._conditionalStyleLookupItemOptions = [];
                this.properties.conditionalStyleLookupPick = '';
                this._conditionalStyleLookupMessage = '';
                if (this.isLookupTypeField(String(newValue || ''))) {
                    this.handleLoadConditionalStyleLookupItems();
                }
            }
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
            if (propertyPath === 'filterDesignerLookupPick' && this.properties.filterDesignerValueType === 'fieldValue') {
                this.properties.filterDesignerValue = String(newValue || '');
                this._filterLookupMessage = newValue ? 'Selected item ID ' + String(newValue) + '.' : '';
            }
            this.context.propertyPane.refresh();
            return;
        }
        _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
    };
    GridControlWebPart.prototype.loadListFields = function (listName) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var endpoint, data, fields, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!listName) {
                            this._listFields = [];
                            this._fieldTypeByInternalName = {};
                            this._fieldChoicesByInternalName = {};
                            this._fieldLookupListByInternalName = {};
                            this.context.propertyPane.refresh();
                            return [2 /*return*/];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        endpoint = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '')
                            + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/fields?$select=InternalName,Title,TypeAsString,LookupList,Choices,Hidden,ReadOnlyField,Sealed&$filter=Hidden eq false and ((ReadOnlyField eq false and Sealed eq false) or InternalName eq 'ID')";
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(endpoint)];
                    case 2:
                        data = _a.sent();
                        fields = data && data.value ? data.value : [];
                        if (!fields || fields.length === 0) {
                            fields = data && data.d && data.d.results ? data.d.results : [];
                        }
                        this._fieldTypeByInternalName = {};
                        this._fieldChoicesByInternalName = {};
                        this._fieldLookupListByInternalName = {};
                        fields.forEach(function (field) {
                            var fieldInternalName = String(field.InternalName || '');
                            if (!fieldInternalName) {
                                return;
                            }
                            _this._fieldTypeByInternalName[fieldInternalName] = String(field.TypeAsString || '').toLowerCase();
                            var rawChoices = field.Choices;
                            _this._fieldChoicesByInternalName[fieldInternalName] = Array.isArray(rawChoices)
                                ? rawChoices.map(function (choice) { return String(choice); })
                                : (rawChoices && Array.isArray(rawChoices.results) ? rawChoices.results.map(function (choice) { return String(choice); }) : []);
                            var lookupListId = field.LookupList ? String(field.LookupList).replace(/[{}]/g, '') : '';
                            if (lookupListId) {
                                _this._fieldLookupListByInternalName[fieldInternalName] = lookupListId;
                            }
                        });
                        this._listFields = fields.map(function (field) {
                            var internalName = String(field.InternalName || '');
                            var title = String(field.Title || internalName);
                            return {
                                key: internalName,
                                text: title + (title !== internalName ? ' (' + internalName + ')' : '')
                            };
                        }).sort(function (a, b) {
                            return String(a.text).localeCompare(String(b.text));
                        });
                        this.context.propertyPane.refresh();
                        return [3 /*break*/, 4];
                    case 3:
                        error_1 = _a.sent();
                        this._listFields = [];
                        this._fieldTypeByInternalName = {};
                        this._fieldChoicesByInternalName = {};
                        this._fieldLookupListByInternalName = {};
                        this.context.propertyPane.refresh();
                        this.logDiagnostic('Failed to load fields for list "' + listName + '": ' + (error_1 && error_1.message ? error_1.message : String(error_1)));
                        return [3 /*break*/, 4];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.isLookupTypeField = function (internalName) {
        var typeName = this._fieldTypeByInternalName[internalName || ''];
        return typeName === 'lookup' || typeName === 'lookupmulti';
    };
    GridControlWebPart.prototype.supportsFilterFieldValue = function (internalName) {
        var typeName = String(this._fieldTypeByInternalName[internalName || ''] || '').toLowerCase();
        return typeName === 'choice' || typeName === 'multichoice' || typeName === 'lookup' || typeName === 'lookupmulti';
    };
    GridControlWebPart.prototype.isMultiFilterField = function (internalName) {
        var typeName = String(this._fieldTypeByInternalName[internalName || ''] || '').toLowerCase();
        return typeName === 'multichoice' || typeName === 'lookupmulti';
    };
    GridControlWebPart.prototype.getSelectedFilterFieldValues = function () {
        var rawValue = String(this.properties.filterDesignerValue || '').trim();
        if (!rawValue) {
            return [];
        }
        try {
            var parsed = JSON.parse(rawValue);
            return Array.isArray(parsed) ? parsed.map(function (value) { return String(value); }) : [String(parsed)];
        }
        catch (_parseError) {
            return [rawValue];
        }
    };
    GridControlWebPart.prototype.createFilterMultiValuePicker = function (options) {
        var _this = this;
        return PropertyPaneCustomField_1.PropertyPaneCustomField({
            key: 'filterDesignerFieldValueMultiPicker',
            onRender: function (domElement) {
                while (domElement.firstChild) {
                    domElement.removeChild(domElement.firstChild);
                }
                var selected = _this.getSelectedFilterFieldValues();
                options.forEach(function (option) {
                    var value = String(option.key);
                    var label = document.createElement('label');
                    label.style.display = 'block';
                    label.style.marginBottom = '6px';
                    var checkbox = document.createElement('input');
                    checkbox.type = 'checkbox';
                    checkbox.checked = selected.indexOf(value) >= 0;
                    checkbox.style.marginRight = '6px';
                    checkbox.addEventListener('change', function () {
                        var nextValues = _this.getSelectedFilterFieldValues();
                        var index = nextValues.indexOf(value);
                        if (checkbox.checked && index < 0) {
                            nextValues.push(value);
                        }
                        if (!checkbox.checked && index >= 0) {
                            nextValues.splice(index, 1);
                        }
                        _this.properties.filterDesignerValue = JSON.stringify(nextValues);
                    });
                    label.appendChild(checkbox);
                    label.appendChild(document.createTextNode(option.text));
                    domElement.appendChild(label);
                });
            }
        });
    };
    GridControlWebPart.prototype.loadLookupListItems = function (listId) {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, url, data, items, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
                        url = webUrl + "/_api/web/lists(guid'" + listId + "')/items?$select=Id,Title&$top=500&$orderby=Title";
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(url)];
                    case 1:
                        data = _a.sent();
                        items = data && data.value ? data.value : (data && data.d && data.d.results ? data.d.results : []);
                        return [2 /*return*/, items.map(function (item) {
                                var title = item.Title ? String(item.Title) : '(no title)';
                                return { key: String(item.Id), text: title + ' (ID: ' + item.Id + ')' };
                            })];
                    case 2:
                        error_2 = _a.sent();
                        this.logDiagnostic('Failed to load lookup list items for listId=' + listId + ': ' + (error_2 && error_2.message ? error_2.message : String(error_2)));
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.handleLoadFilterLookupItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var fieldName, listId, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        fieldName = String(this.properties.filterDesignerField || '');
                        listId = this._fieldLookupListByInternalName[fieldName];
                        if (!listId) {
                            this._filterLookupMessage = 'Selected field is not a lookup column, or its target list could not be resolved.';
                            this._filterLookupItemOptions = [];
                            this.context.propertyPane.refresh();
                            return [2 /*return*/];
                        }
                        _a = this;
                        return [4 /*yield*/, this.loadLookupListItems(listId)];
                    case 1:
                        _a._filterLookupItemOptions = _b.sent();
                        this._filterLookupMessage = this._filterLookupItemOptions.length > 0
                            ? 'Loaded ' + this._filterLookupItemOptions.length + ' item(s). Pick one and click Use selected value.'
                            : 'The target list has no items.';
                        this.context.propertyPane.refresh();
                        return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.handleApplyFilterLookupPick = function () {
        var picked = String(this.properties.filterDesignerLookupPick || '');
        if (!picked) {
            this._filterLookupMessage = 'Select an item from the dropdown before applying.';
            this.context.propertyPane.refresh();
            return;
        }
        this.properties.filterDesignerValue = picked;
        this._filterLookupMessage = 'Applied "' + picked + '" to the filter value.';
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleLoadConditionalStyleLookupItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var fieldName, listId, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        fieldName = String(this.properties.conditionalStyleConditionField || '');
                        listId = this._fieldLookupListByInternalName[fieldName];
                        if (!listId) {
                            this._conditionalStyleLookupMessage = 'Selected field is not a lookup column, or its target list could not be resolved.';
                            this._conditionalStyleLookupItemOptions = [];
                            this.context.propertyPane.refresh();
                            return [2 /*return*/];
                        }
                        _a = this;
                        return [4 /*yield*/, this.loadLookupListItems(listId)];
                    case 1:
                        _a._conditionalStyleLookupItemOptions = _b.sent();
                        this._conditionalStyleLookupMessage = this._conditionalStyleLookupItemOptions.length > 0
                            ? 'Loaded ' + this._conditionalStyleLookupItemOptions.length + ' item(s). Pick one and click Use selected value.'
                            : 'The target list has no items.';
                        this.context.propertyPane.refresh();
                        return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.handleApplyConditionalStyleLookupPick = function () {
        var picked = String(this.properties.conditionalStyleLookupPick || '');
        if (!picked) {
            this._conditionalStyleLookupMessage = 'Select an item from the results before applying.';
            this.context.propertyPane.refresh();
            return;
        }
        this.properties.conditionalStyleConditionValue = picked;
        this._conditionalStyleLookupMessage = 'Applied "' + picked + '" to the condition value.';
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.createPropertyPanePageTitle = function (key, title) {
        return PropertyPaneCustomField_1.PropertyPaneCustomField({
            key: key,
            onRender: function (domElement) {
                domElement.textContent = title;
                domElement.style.fontSize = '20px';
                domElement.style.fontWeight = '600';
                domElement.style.lineHeight = '28px';
                domElement.style.marginBottom = '12px';
            }
        });
    };
    GridControlWebPart.prototype.createFilterExpressionHelpField = function () {
        var _this = this;
        var isExpanded = this._showFilterExpressionHelp;
        return PropertyPaneCustomField_1.PropertyPaneCustomField({
            key: 'filterExpressionHelp-' + String(isExpanded),
            onRender: function (domElement) {
                domElement.innerHTML = '';
                var button = document.createElement('button');
                button.type = 'button';
                button.title = isExpanded ? strings.PropFilterDesignerExpressionHelpHide : strings.PropFilterDesignerExpressionHelpShow;
                button.setAttribute('aria-label', button.title);
                button.setAttribute('aria-expanded', String(isExpanded));
                button.style.background = 'transparent';
                button.style.border = '0';
                button.style.color = '#0078d4';
                button.style.cursor = 'pointer';
                button.style.padding = '2px';
                var icon = document.createElement('span');
                icon.className = 'ms-Icon ms-Icon--Info';
                icon.setAttribute('aria-hidden', 'true');
                button.appendChild(icon);
                button.onclick = function () {
                    _this._showFilterExpressionHelp = !_this._showFilterExpressionHelp;
                    _this.context.propertyPane.refresh();
                };
                domElement.appendChild(button);
                if (isExpanded) {
                    var help = document.createElement('div');
                    help.textContent = strings.PropFilterDesignerExpressionHelp;
                    help.style.fontSize = '12px';
                    help.style.lineHeight = '18px';
                    help.style.marginTop = '4px';
                    domElement.appendChild(help);
                }
            }
        });
    };
    GridControlWebPart.prototype.createConditionalStyleExpressionHelpField = function () {
        var _this = this;
        var isExpanded = this._showConditionalStyleExpressionHelp;
        return PropertyPaneCustomField_1.PropertyPaneCustomField({
            key: 'conditionalStyleExpressionHelp-' + String(isExpanded),
            onRender: function (domElement) {
                domElement.innerHTML = '';
                var button = document.createElement('button');
                button.type = 'button';
                button.title = isExpanded ? strings.PropFilterDesignerExpressionHelpHide : strings.PropFilterDesignerExpressionHelpShow;
                button.setAttribute('aria-label', button.title);
                button.setAttribute('aria-expanded', String(isExpanded));
                button.style.background = 'transparent';
                button.style.border = '0';
                button.style.color = '#0078d4';
                button.style.cursor = 'pointer';
                button.style.padding = '2px';
                var icon = document.createElement('span');
                icon.className = 'ms-Icon ms-Icon--Info';
                icon.setAttribute('aria-hidden', 'true');
                button.appendChild(icon);
                button.onclick = function () {
                    _this._showConditionalStyleExpressionHelp = !_this._showConditionalStyleExpressionHelp;
                    _this.context.propertyPane.refresh();
                };
                domElement.appendChild(button);
                if (isExpanded) {
                    var help = document.createElement('div');
                    help.textContent = strings.PropFilterDesignerExpressionHelp;
                    help.style.fontSize = '12px';
                    help.style.lineHeight = '18px';
                    help.style.marginTop = '4px';
                    domElement.appendChild(help);
                }
            }
        });
    };
    GridControlWebPart.prototype.handleColorPropertyChange = function (propertyPath, oldValue, newValue) {
        this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        this.context.propertyPane.refresh();
        this.render();
    };
    GridControlWebPart.prototype.onDispose = function () {
        deterministicFullWidth_1.releaseOptionalFullWidth(this.domElement.ownerDocument, this.context.instanceId);
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(GridControlWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    GridControlWebPart.prototype.getPropertyPaneConfiguration = function () {
        var fontFamilyOptions = this.getFontFamilyOptions();
        var fontStyleOptions = this.getFontStyleOptions();
        var fullVersionLabel = 'Version: ' + this.getWebPartVersion();
        var filterFieldOptions = this._listFields.length > 0 ? this._listFields : [{ key: '', text: strings.PropFilterDesignerFieldNone }];
        var filterConditionOptions = [{ key: '', text: strings.PropFilterDesignerExistingNone }];
        var conditionalStyleFieldOptions = this._listFields.length > 0 ? this._listFields : [{ key: '', text: strings.PropConditionalStyleFieldNone }];
        var conditionalStyleConditionOptions = [{ key: '', text: strings.PropConditionalStyleExistingNone }];
        var conditionalStyleRuleSummaryText = strings.PropConditionalStyleSummaryNone;
        try {
            var existingConditions = this.parseFilterJsonArray(this.properties.filterJson);
            if (existingConditions.length > 0) {
                filterConditionOptions = existingConditions.map(function (condition, index) {
                    var field = String((condition && condition.field) || '(field)');
                    var operator = String((condition && condition.operator) || 'contains').toLowerCase();
                    var logical = String((condition && condition.logical) || 'and').toUpperCase();
                    var valueType = String((condition && condition.valueType) || 'static').toLowerCase();
                    var value = condition && condition.value !== undefined ? String(condition.value) : '';
                    var logicalPrefix = index === 0 ? '' : logical + ' ';
                    return {
                        key: String(index),
                        text: '#' + String(index + 1) + ': ' + logicalPrefix + field + ' ' + operator + ' ' + (valueType === 'expression' ? '[' + value + ']' : value)
                    };
                });
            }
        }
        catch (_parseFilterError) {
            filterConditionOptions = [{ key: '', text: strings.PropFilterDesignerExistingNone }];
        }
        try {
            var existingConditionalStyles = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            if (existingConditionalStyles.length > 0) {
                conditionalStyleConditionOptions = existingConditionalStyles.map(function (condition, index) {
                    var rule = String((condition && condition.rule) || '(rule)');
                    var enabled = condition && condition.enabled !== false;
                    var priority = parseInt(String(condition && condition.priority !== undefined ? condition.priority : '100'), 10);
                    var scope = String((condition && condition.scope) || 'row').toLowerCase();
                    var applyField = String((condition && condition.applyField) || '');
                    var field = String((condition && condition.conditionField) || '(field)');
                    var operator = String((condition && condition.operator) || 'contains').toLowerCase();
                    var logical = String((condition && condition.logical) || 'and').toUpperCase();
                    var valueType = String((condition && condition.valueType) || 'static').toLowerCase();
                    var value = condition && condition.value !== undefined ? String(condition.value) : '';
                    var scopeInfo = scope === 'column' ? ' -> ' + (applyField || '(column)') : ' -> row';
                    var statusInfo = enabled ? 'enabled' : 'disabled';
                    var priorityInfo = isNaN(priority) ? 'P100' : ('P' + String(priority));
                    var logicalPrefix = index === 0 ? '' : logical + ' ';
                    return {
                        key: String(index),
                        text: '#' + String(index + 1) + ': [' + priorityInfo + '|' + statusInfo + '] ' + rule + scopeInfo + ' | ' + logicalPrefix + field + ' ' + operator + ' ' + (valueType === 'expression' ? '[' + value + ']' : value)
                    };
                });
                var groupedSummaries = {};
                var groupedOrder = [];
                for (var summaryIndex = 0; summaryIndex < existingConditionalStyles.length; summaryIndex += 1) {
                    var summaryCondition = existingConditionalStyles[summaryIndex] || {};
                    var summaryRuleName = String(summaryCondition.rule || '(rule)').trim() || '(rule)';
                    if (!groupedSummaries[summaryRuleName]) {
                        groupedSummaries[summaryRuleName] = {
                            name: summaryRuleName,
                            enabled: summaryCondition.enabled !== false,
                            priority: parseInt(String(summaryCondition.priority !== undefined ? summaryCondition.priority : '100'), 10),
                            scope: String(summaryCondition.scope || 'row').toLowerCase(),
                            applyField: String(summaryCondition.applyField || ''),
                            count: 0,
                            firstIndex: summaryIndex
                        };
                        groupedOrder.push(summaryRuleName);
                    }
                    groupedSummaries[summaryRuleName].count += 1;
                }
                var groupedItems = groupedOrder.map(function (ruleName) { return groupedSummaries[ruleName]; });
                groupedItems.sort(function (left, right) {
                    var leftPriority = isNaN(left.priority) ? 100 : left.priority;
                    var rightPriority = isNaN(right.priority) ? 100 : right.priority;
                    if (leftPriority !== rightPriority) {
                        return leftPriority - rightPriority;
                    }
                    return left.firstIndex - right.firstIndex;
                });
                conditionalStyleRuleSummaryText = groupedItems.map(function (item) {
                    var enabledText = item.enabled ? strings.PropConditionalStyleSummaryEnabled : strings.PropConditionalStyleSummaryDisabled;
                    var priorityText = strings.PropConditionalStyleSummaryPriorityPrefix + String(isNaN(item.priority) ? 100 : item.priority);
                    var scopeText = item.scope === 'column'
                        ? strings.PropConditionalStyleSummaryScopeColumn + ' ' + (item.applyField || '(column)')
                        : strings.PropConditionalStyleSummaryScopeRow;
                    var conditionText = strings.PropConditionalStyleSummaryConditionCountPrefix + String(item.count);
                    return item.name + ' | ' + enabledText + ' | ' + priorityText + ' | ' + scopeText + ' | ' + conditionText;
                }).join('\n');
            }
        }
        catch (_parseConditionalStyleError) {
            conditionalStyleConditionOptions = [{ key: '', text: strings.PropConditionalStyleExistingNone }];
            conditionalStyleRuleSummaryText = strings.PropConditionalStyleSummaryInvalid;
        }
        return {
            pages: [
                {
                    header: {
                        description: ''
                    },
                    groups: [
                        {
                            groupName: fullVersionLabel,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneLabel('propertyPaneVersionInfo', {
                                    text: ' '
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyPanePageGuideTitle,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneLabel('propertyPanePageGuideFilters', {
                                    text: strings.PropertyPanePageGuideFilters
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('propertyPanePageGuideStyles', {
                                    text: strings.PropertyPanePageGuideStyles
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyGroupData,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('instanceName', {
                                    label: strings.PropInstanceNameLabel,
                                    description: strings.PropInstanceNameDescription,
                                    placeholder: strings.DynamicSourceTitle,
                                    value: this.properties.instanceName || ''
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('forceFullWidth', { text: 'Force Full Width' }),
                                sp_webpart_base_1.PropertyPaneDropdown('listName', {
                                    label: strings.PropListLabel,
                                    options: this._lists,
                                    selectedKey: this.properties.listName
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('viewId', {
                                    label: strings.PropViewLabel,
                                    options: this._views.length > 0 ? this._views : [{ key: '', text: strings.PropNoViews }],
                                    selectedKey: this.properties.viewId || ''
                                }),
                                sp_webpart_base_1.PropertyPaneButton('openGridDesigner', {
                                    text: strings.PropGridDesignerButton,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Primary,
                                    disabled: !this.properties.listName,
                                    onClick: this.openGridDesigner.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('gridDesignerStatus', {
                                    text: this.getGridDesignerStatus()
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('pageSize', {
                                    label: strings.PropPageSizeLabel,
                                    description: strings.PropPageSizeDescription,
                                    value: String(this.properties.pageSize || '')
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('showViewSelector', {
                                    text: strings.PropShowViewSelectorLabel,
                                    checked: this.properties.showViewSelector !== false
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('showLinkToItem', {
                                    text: strings.PropShowLinkToItemLabel,
                                    checked: this.properties.showLinkToItem === true
                                })
                            ].concat((this.properties.showLinkToItem === true ? [
                                sp_webpart_base_1.PropertyPaneDropdown('linkTargetPageUrl', {
                                    label: strings.PropLinkTargetPageUrlLabel,
                                    options: this.getTargetPageOptions(),
                                    selectedKey: this.properties.linkTargetPageUrl || '__defaultForm__'
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('linkTargetIdParam', {
                                    label: strings.PropLinkTargetIdParamLabel,
                                    placeholder: strings.PropLinkTargetIdParamPlaceholder,
                                    value: this.properties.linkTargetIdParam || 'itemid'
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('includeReturnUrlParam', {
                                    text: strings.PropIncludeReturnUrlParamLabel,
                                    checked: this.properties.includeReturnUrlParam === true
                                })
                            ] : []))
                        },
                        {
                            groupName: strings.PropertyGroupButtons,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneCheckbox('showRefresh', {
                                    text: strings.PropShowRefreshLabel,
                                    checked: this.properties.showRefresh !== false
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('showAdd', {
                                    text: strings.PropShowAddLabel,
                                    checked: this.properties.showAdd !== false
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('showDelete', {
                                    text: strings.PropShowDeleteLabel,
                                    checked: this.properties.showDelete !== false
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyGroupBodyStyle,
                            groupFields: [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('bodyTextColor', {
                                    label: strings.PropBodyTextColorLabel,
                                    selectedColor: this.properties.bodyTextColor || '',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'bodyTextColorField'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('bodyFontFamily', {
                                    label: strings.PropBodyFontFamilyLabel,
                                    options: fontFamilyOptions,
                                    selectedKey: this.properties.bodyFontFamily || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('bodyFontSize', {
                                    label: strings.PropBodyFontSizeLabel,
                                    options: [
                                        { key: '12px', text: '12px' },
                                        { key: '14px', text: '14px' },
                                        { key: '16px', text: '16px' },
                                        { key: '18px', text: '18px' },
                                        { key: '20px', text: '20px' }
                                    ],
                                    selectedKey: this.properties.bodyFontSize || '14px'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('bodyFontStyle', {
                                    label: strings.PropBodyFontStyleLabel,
                                    options: fontStyleOptions,
                                    selectedKey: this.properties.bodyFontStyle || 'normal'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('bodyTextAlign', {
                                    label: strings.PropBodyTextAlignLabel,
                                    options: this.getTextAlignOptions(),
                                    selectedKey: this.properties.bodyTextAlign || 'left'
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('bodyFontBold', {
                                    text: strings.PropBodyBoldLabel,
                                    checked: this.properties.bodyFontBold === true
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyGroupDateDisplay,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneDropdown('dateDisplayFormat', {
                                    label: strings.PropDateDisplayFormatLabel,
                                    options: [
                                        { key: 'mdy', text: 'MM/DD/YYYY' },
                                        { key: 'dmy', text: 'DD/MM/YYYY' },
                                        { key: 'ymd', text: 'YYYY-MM-DD' }
                                    ],
                                    selectedKey: this.properties.dateDisplayFormat || 'mdy'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('timeDisplayFormat', {
                                    label: strings.PropTimeDisplayFormatLabel,
                                    options: [
                                        { key: '24hour', text: '24-hour (HH:mm)' },
                                        { key: '12hour', text: '12-hour (hh:mm AM/PM)' }
                                    ],
                                    selectedKey: this.properties.timeDisplayFormat || '24hour'
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyGroupSelectedStyle,
                            groupFields: [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('selectedTextColor', {
                                    label: strings.PropSelectedTextColorLabel,
                                    selectedColor: this.properties.selectedTextColor || '',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'selectedTextColorField'
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('selectedBackgroundColor', {
                                    label: strings.PropSelectedBackgroundColorLabel,
                                    selectedColor: this.properties.selectedBackgroundColor || '#eef6ff',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'selectedBackgroundColorField'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('selectedFontStyle', {
                                    label: strings.PropSelectedFontStyleLabel,
                                    options: fontStyleOptions,
                                    selectedKey: this.properties.selectedFontStyle || 'normal'
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('selectedFontBold', {
                                    text: strings.PropSelectedBoldLabel,
                                    checked: this.properties.selectedFontBold === true
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyGroupHeaderStyle,
                            groupFields: [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('headerTextColor', {
                                    label: strings.PropHeaderTextColorLabel,
                                    selectedColor: this.properties.headerTextColor || '',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'headerTextColorField'
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('headerBackgroundColor', {
                                    label: strings.PropHeaderBackgroundColorLabel,
                                    selectedColor: this.properties.headerBackgroundColor || '',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'headerBackgroundColorField'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('headerFontFamily', {
                                    label: strings.PropHeaderFontFamilyLabel,
                                    options: fontFamilyOptions,
                                    selectedKey: this.properties.headerFontFamily || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('headerFontSize', {
                                    label: strings.PropHeaderFontSizeLabel,
                                    options: [
                                        { key: '12px', text: '12px' },
                                        { key: '14px', text: '14px' },
                                        { key: '16px', text: '16px' },
                                        { key: '18px', text: '18px' },
                                        { key: '20px', text: '20px' }
                                    ],
                                    selectedKey: this.properties.headerFontSize || '14px'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('headerFontStyle', {
                                    label: strings.PropHeaderFontStyleLabel,
                                    options: fontStyleOptions,
                                    selectedKey: this.properties.headerFontStyle || 'normal'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('headerTextAlign', {
                                    label: strings.PropHeaderTextAlignLabel,
                                    options: this.getTextAlignOptions(),
                                    selectedKey: this.properties.headerTextAlign || 'left'
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('headerFontBold', {
                                    text: strings.PropHeaderBoldLabel,
                                    checked: this.properties.headerFontBold === true
                                })
                            ]
                        },
                        {
                            groupName: strings.PropertyGroupTableStyle,
                            groupFields: [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('tableBackgroundColor', {
                                    label: strings.PropTableBackgroundColorLabel,
                                    selectedColor: this.properties.tableBackgroundColor || '#ffffff',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'tableBackgroundColorField'
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('tableBorderColor', {
                                    label: strings.PropTableBorderColorLabel,
                                    selectedColor: this.properties.tableBorderColor || '#cccccc',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'tableBorderColorField'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('tableBorderWidth', {
                                    label: strings.PropTableBorderWidthLabel,
                                    options: [
                                        { key: '0px', text: 'None' },
                                        { key: '1px', text: '1px' },
                                        { key: '2px', text: '2px' },
                                        { key: '3px', text: '3px' },
                                        { key: '4px', text: '4px' },
                                        { key: '5px', text: '5px' }
                                    ],
                                    selectedKey: this.properties.tableBorderWidth || '1px'
                                }),
                                sp_webpart_base_1.PropertyPaneSlider('tableRowLineWidth', {
                                    label: strings.PropTableRowLineWidthLabel,
                                    min: 0,
                                    max: 10,
                                    step: 1,
                                    value: typeof this.properties.tableRowLineWidth === 'number' ? this.properties.tableRowLineWidth : 1
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('tableCornerStyle', {
                                    label: strings.PropTableCornerStyleLabel,
                                    options: [
                                        { key: 'square', text: strings.PropCornerStyleSquare },
                                        { key: 'rounded', text: strings.PropCornerStyleRounded }
                                    ],
                                    selectedKey: this.properties.tableCornerStyle || 'square'
                                })
                            ].concat(((this.properties.tableCornerStyle || 'square') === 'rounded' ? [
                                sp_webpart_base_1.PropertyPaneSlider('tableCornerRadius', {
                                    label: strings.PropTableCornerRadiusLabel,
                                    min: 0,
                                    max: 40,
                                    step: 1,
                                    value: typeof this.properties.tableCornerRadius === 'number' ? this.properties.tableCornerRadius : 8
                                })
                            ] : []), [
                                sp_webpart_base_1.PropertyPaneCheckbox('alternateRowShading', {
                                    text: strings.PropAlternateRowShadingLabel,
                                    checked: this.properties.alternateRowShading === true
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('alternateRowShadingColor', {
                                    label: strings.PropAlternateRowShadingColorLabel,
                                    selectedColor: this.properties.alternateRowShadingColor || '#f9f9f9',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'alternateRowShadingColorField'
                                })
                            ])
                        },
                        { groupName: strings.PropertyGroupWebpartStyle,
                            groupFields: [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('webpartBackgroundColor', {
                                    label: strings.PropWebpartBackgroundColorLabel,
                                    selectedColor: this.properties.webpartBackgroundColor || '#ffffff',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'webpartBackgroundColorField'
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('webpartBorderColor', {
                                    label: strings.PropWebpartBorderColorLabel,
                                    selectedColor: this.properties.webpartBorderColor || '#ccc',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'webpartBorderColorField'
                                }),
                                sp_webpart_base_1.PropertyPaneSlider('webpartBorderWidth', {
                                    label: strings.PropWebpartBorderWidthLabel,
                                    min: 0,
                                    max: 20,
                                    step: 1,
                                    value: this.properties.webpartBorderWidth || 0
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('cornerStyle', {
                                    label: strings.PropCornerStyleLabel,
                                    options: [
                                        { key: 'square', text: strings.PropCornerStyleSquare },
                                        { key: 'rounded', text: strings.PropCornerStyleRounded }
                                    ],
                                    selectedKey: this.properties.cornerStyle || 'square'
                                })
                            ].concat(((this.properties.cornerStyle || 'square') === 'rounded' ? [
                                sp_webpart_base_1.PropertyPaneSlider('cornerRadius', {
                                    label: strings.PropCornerRadiusLabel,
                                    min: 0,
                                    max: 40,
                                    step: 1,
                                    value: typeof this.properties.cornerRadius === 'number' ? this.properties.cornerRadius : 12
                                })
                            ] : []))
                        },
                        { groupName: strings.PropertyGroupButtonStyle,
                            groupFields: [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('buttonTextColor', {
                                    label: strings.PropButtonTextColorLabel,
                                    selectedColor: this.properties.buttonTextColor || '#000000',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'buttonTextColorField'
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('buttonBackgroundColor', {
                                    label: strings.PropButtonBackgroundColorLabel,
                                    selectedColor: this.properties.buttonBackgroundColor || '#f0f0f0',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'buttonBackgroundColorField'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('buttonFontFamily', {
                                    label: strings.PropButtonFontFamilyLabel,
                                    options: fontFamilyOptions,
                                    selectedKey: this.properties.buttonFontFamily || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('buttonFontSize', {
                                    label: strings.PropButtonFontSizeLabel,
                                    options: [
                                        { key: '12px', text: '12px' },
                                        { key: '14px', text: '14px' },
                                        { key: '16px', text: '16px' },
                                        { key: '18px', text: '18px' },
                                        { key: '20px', text: '20px' }
                                    ],
                                    selectedKey: this.properties.buttonFontSize || '14px'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('buttonFontStyle', {
                                    label: strings.PropButtonFontStyleLabel,
                                    options: fontStyleOptions,
                                    selectedKey: this.properties.buttonFontStyle || 'normal'
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('buttonFontBold', {
                                    text: strings.PropButtonBoldLabel,
                                    checked: this.properties.buttonFontBold === true
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('buttonCornerStyle', {
                                    label: strings.PropButtonCornerStyleLabel,
                                    options: [
                                        { key: 'square', text: strings.PropCornerStyleSquare },
                                        { key: 'rounded', text: strings.PropCornerStyleRounded }
                                    ],
                                    selectedKey: this.properties.buttonCornerStyle || 'square'
                                })
                            ].concat(((this.properties.buttonCornerStyle || 'square') === 'rounded' ? [
                                sp_webpart_base_1.PropertyPaneSlider('buttonCornerRadius', {
                                    label: strings.PropButtonCornerRadiusLabel,
                                    min: 0,
                                    max: 40,
                                    step: 1,
                                    value: typeof this.properties.buttonCornerRadius === 'number' ? this.properties.buttonCornerRadius : 4
                                })
                            ] : []))
                        },
                        {
                            groupName: 'Diagnostics',
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneCheckbox('enableDiagnostics', {
                                    text: strings.PropEnableDiagnosticsLabel,
                                    checked: this.properties.enableDiagnostics !== false
                                })
                            ]
                        },
                        {
                            groupName: '',
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneLabel('propertyPaneBottomSpacer', {
                                    text: ' '
                                })
                            ]
                        }
                    ]
                },
                {
                    header: {
                        description: ''
                    },
                    groups: [
                        {
                            groupName: '',
                            groupFields: [
                                this.createPropertyPanePageTitle('presetFilterPageTitle', strings.PresetFilterPageTitle)
                            ]
                        },
                        {
                            groupName: strings.PresetFilterDesignerGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerField', {
                                    label: strings.PropFilterDesignerFieldLabel,
                                    options: filterFieldOptions,
                                    selectedKey: this.properties.filterDesignerField || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerOperator', {
                                    label: strings.PropFilterDesignerOperatorLabel,
                                    options: [
                                        { key: 'eq', text: strings.RuntimeFilterOperatorEquals },
                                        { key: 'ne', text: strings.RuntimeFilterOperatorNotEquals },
                                        { key: 'contains', text: strings.RuntimeFilterOperatorContains },
                                        { key: 'notcontains', text: strings.RuntimeFilterOperatorNotContains },
                                        { key: 'startswith', text: strings.RuntimeFilterOperatorStartsWith },
                                        { key: 'endswith', text: strings.RuntimeFilterOperatorEndsWith },
                                        { key: 'gt', text: strings.RuntimeFilterOperatorGreaterThan },
                                        { key: 'ge', text: strings.RuntimeFilterOperatorGreaterThanOrEqual },
                                        { key: 'lt', text: strings.RuntimeFilterOperatorLessThan },
                                        { key: 'le', text: strings.RuntimeFilterOperatorLessThanOrEqual }
                                    ],
                                    selectedKey: this.properties.filterDesignerOperator || 'contains'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerLogical', {
                                    label: strings.PropFilterDesignerLogicalLabel,
                                    options: [
                                        { key: 'and', text: strings.PropFilterDesignerLogicalAnd },
                                        { key: 'or', text: strings.PropFilterDesignerLogicalOr }
                                    ],
                                    selectedKey: this.properties.filterDesignerLogical || 'and'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerValueType', {
                                    label: strings.PropFilterDesignerValueTypeLabel,
                                    options: [
                                        { key: 'static', text: strings.PropFilterDesignerValueTypeStatic },
                                        { key: 'expression', text: strings.PropFilterDesignerValueTypeExpression }
                                    ].concat((this.supportsFilterFieldValue(this.properties.filterDesignerField || '') ? [{ key: 'fieldValue', text: 'Field value' }] : [])),
                                    selectedKey: this.properties.filterDesignerValueType || 'static'
                                })
                            ].concat((this.properties.filterDesignerValueType !== 'fieldValue' ? [sp_webpart_base_1.PropertyPaneTextField('filterDesignerValue', {
                                    label: this.properties.filterDesignerValueType === 'expression' ? strings.PropFilterDesignerExpressionLabel : strings.PropFilterDesignerValueLabel,
                                    placeholder: this.properties.filterDesignerValueType === 'expression' ? strings.PropFilterDesignerExpressionPlaceholder : strings.PropFilterDesignerValuePlaceholder,
                                    value: this.properties.filterDesignerValue || ''
                                })] : []), (this.properties.filterDesignerValueType === 'expression' ? [
                                this.createFilterExpressionHelpField()
                            ] : []), (this.properties.filterDesignerValueType === 'fieldValue' && !this.isMultiFilterField(this.properties.filterDesignerField || '') && (this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).length > 0 ? [
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerValue', {
                                    label: 'Select field value',
                                    options: (this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).map(function (choice) { return { key: choice, text: choice }; }),
                                    selectedKey: this.properties.filterDesignerValue || ''
                                })
                            ] : []), (this.properties.filterDesignerValueType === 'fieldValue' && this.isMultiFilterField(this.properties.filterDesignerField || '') ? [
                                this.createFilterMultiValuePicker(this.isLookupTypeField(this.properties.filterDesignerField || '') ? this._filterLookupItemOptions : (this._fieldChoicesByInternalName[this.properties.filterDesignerField || ''] || []).map(function (choice) { return { key: choice, text: choice }; }))
                            ] : []), (this.properties.filterDesignerValueType !== 'expression' && this.isLookupTypeField(this.properties.filterDesignerField || '') && !this.isMultiFilterField(this.properties.filterDesignerField || '') ? [
                                sp_webpart_base_1.PropertyPaneLabel('filterLookupHelperTitle', {
                                    text: 'Lookup item picker: choose an item from the target list to fill in its value.'
                                }),
                                sp_webpart_base_1.PropertyPaneButton('loadFilterLookupItems', {
                                    text: 'Refresh lookup items',
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleLoadFilterLookupItems.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerLookupPick', {
                                    label: 'Select item',
                                    options: this._filterLookupItemOptions.length > 0 ? this._filterLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                                    selectedKey: this.properties.filterDesignerLookupPick || ''
                                }),
                                sp_webpart_base_1.PropertyPaneButton('applyFilterLookupPick', {
                                    text: 'Use selected value',
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleApplyFilterLookupPick.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('filterLookupHelperResult', {
                                    text: this._filterLookupMessage || 'Lookup items load automatically when you select the column.'
                                }),
                            ] : []), [
                                sp_webpart_base_1.PropertyPaneButton('addFilterDesignerCondition', {
                                    text: strings.PropFilterDesignerAddButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Primary,
                                    onClick: this.handleAddFilterDesignerCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('filterDesignerSelectedIndex', {
                                    label: strings.PropFilterDesignerExistingLabel,
                                    options: filterConditionOptions,
                                    selectedKey: this.properties.filterDesignerSelectedIndex || ''
                                }),
                                sp_webpart_base_1.PropertyPaneButton('loadFilterDesignerCondition', {
                                    text: strings.PropFilterDesignerLoadButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleLoadSelectedFilterCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('updateFilterDesignerCondition', {
                                    text: strings.PropFilterDesignerUpdateButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleUpdateSelectedFilterCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('removeFilterDesignerCondition', {
                                    text: strings.PropFilterDesignerRemoveButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleRemoveSelectedFilterCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('resetFilterDesignerJson', {
                                    text: strings.PropFilterDesignerResetButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleResetFilterJson.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('filterDesignerResult', {
                                    text: this._filterDesignerMessage || strings.PropFilterDesignerStatusPlaceholder
                                })
                            ])
                        },
                        {
                            groupName: strings.PresetFilterJsonGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('filterJson', {
                                    label: strings.PropFilterJsonLabel,
                                    placeholder: strings.PropFilterJsonPlaceholder,
                                    value: this.properties.filterJson || '',
                                    multiline: true,
                                    resizable: true,
                                    rows: 10,
                                    description: strings.PropFilterJsonHelp
                                }),
                                sp_webpart_base_1.PropertyPaneButton('validateFilterJson', {
                                    text: strings.PropValidateFilterJsonButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleValidateFilterJson.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('validateFilterJsonResult', {
                                    text: this._filterJsonValidationMessage || strings.JsonValidationResultPlaceholder
                                })
                            ]
                        }
                    ]
                },
                {
                    header: {
                        description: ''
                    },
                    groups: [
                        {
                            groupName: '',
                            groupFields: [
                                this.createPropertyPanePageTitle('conditionalStylePageTitle', strings.ConditionalStylePageTitle)
                            ]
                        },
                        {
                            groupName: strings.ConditionalStyleSummaryGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneLabel('conditionalStyleRuleSummaryLabel', {
                                    text: conditionalStyleRuleSummaryText
                                })
                            ]
                        },
                        {
                            groupName: strings.ConditionalStyleDesignerGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('conditionalStyleRuleName', {
                                    label: strings.PropConditionalStyleRuleNameLabel,
                                    placeholder: strings.PropConditionalStyleRuleNamePlaceholder,
                                    value: this.properties.conditionalStyleRuleName || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleScope', {
                                    label: strings.PropConditionalStyleScopeLabel,
                                    options: [
                                        { key: 'row', text: strings.PropConditionalStyleScopeRow },
                                        { key: 'column', text: strings.PropConditionalStyleScopeColumn }
                                    ],
                                    selectedKey: this.properties.conditionalStyleScope || 'row'
                                })
                            ].concat((this.properties.conditionalStyleScope === 'column' ? [
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleApplyField', {
                                    label: strings.PropConditionalStyleApplyFieldLabel,
                                    options: conditionalStyleFieldOptions,
                                    selectedKey: this.properties.conditionalStyleApplyField || ''
                                })
                            ] : []), [
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleConditionField', {
                                    label: strings.PropConditionalStyleConditionFieldLabel,
                                    options: conditionalStyleFieldOptions,
                                    selectedKey: this.properties.conditionalStyleConditionField || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleConditionOperator', {
                                    label: strings.PropConditionalStyleConditionOperatorLabel,
                                    options: [
                                        { key: 'eq', text: strings.RuntimeFilterOperatorEquals },
                                        { key: 'ne', text: strings.RuntimeFilterOperatorNotEquals },
                                        { key: 'contains', text: strings.RuntimeFilterOperatorContains },
                                        { key: 'notcontains', text: strings.RuntimeFilterOperatorNotContains },
                                        { key: 'startswith', text: strings.RuntimeFilterOperatorStartsWith },
                                        { key: 'endswith', text: strings.RuntimeFilterOperatorEndsWith },
                                        { key: 'gt', text: strings.RuntimeFilterOperatorGreaterThan },
                                        { key: 'ge', text: strings.RuntimeFilterOperatorGreaterThanOrEqual },
                                        { key: 'lt', text: strings.RuntimeFilterOperatorLessThan },
                                        { key: 'le', text: strings.RuntimeFilterOperatorLessThanOrEqual }
                                    ],
                                    selectedKey: this.properties.conditionalStyleConditionOperator || 'eq'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleConditionLogical', {
                                    label: strings.PropConditionalStyleConditionLogicalLabel,
                                    options: [
                                        { key: 'and', text: strings.PropFilterDesignerLogicalAnd },
                                        { key: 'or', text: strings.PropFilterDesignerLogicalOr }
                                    ],
                                    selectedKey: this.properties.conditionalStyleConditionLogical || 'and'
                                }),
                                sp_webpart_base_1.PropertyPaneCheckbox('conditionalStyleEnabled', {
                                    text: strings.PropConditionalStyleEnabledLabel,
                                    checked: this.properties.conditionalStyleEnabled !== false
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('conditionalStylePriority', {
                                    label: strings.PropConditionalStylePriorityLabel,
                                    placeholder: strings.PropConditionalStylePriorityPlaceholder,
                                    value: this.properties.conditionalStylePriority || '100'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleConditionValueType', {
                                    label: strings.PropFilterDesignerValueTypeLabel,
                                    options: [
                                        { key: 'static', text: strings.PropFilterDesignerValueTypeStatic },
                                        { key: 'expression', text: strings.PropFilterDesignerValueTypeExpression }
                                    ],
                                    selectedKey: this.properties.conditionalStyleConditionValueType || 'static'
                                }),
                                sp_webpart_base_1.PropertyPaneTextField('conditionalStyleConditionValue', {
                                    label: this.properties.conditionalStyleConditionValueType === 'expression' ? strings.PropFilterDesignerExpressionLabel : strings.PropConditionalStyleConditionValueLabel,
                                    placeholder: this.properties.conditionalStyleConditionValueType === 'expression' ? strings.PropFilterDesignerExpressionPlaceholder : strings.PropConditionalStyleConditionValuePlaceholder,
                                    value: this.properties.conditionalStyleConditionValue || ''
                                })
                            ], (this.properties.conditionalStyleConditionValueType === 'expression' ? [
                                this.createConditionalStyleExpressionHelpField()
                            ] : []), (this.properties.conditionalStyleConditionValueType !== 'expression' && this.isLookupTypeField(this.properties.conditionalStyleConditionField || '') ? [
                                sp_webpart_base_1.PropertyPaneLabel('conditionalStyleLookupHelperTitle', {
                                    text: 'Lookup item picker: choose an item from the target list to fill in its value.'
                                }),
                                sp_webpart_base_1.PropertyPaneButton('loadConditionalStyleLookupItems', {
                                    text: 'Refresh lookup items',
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleLoadConditionalStyleLookupItems.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleLookupPick', {
                                    label: 'Select item',
                                    options: this._conditionalStyleLookupItemOptions.length > 0 ? this._conditionalStyleLookupItemOptions : [{ key: '', text: 'No items loaded yet' }],
                                    selectedKey: this.properties.conditionalStyleLookupPick || ''
                                }),
                                sp_webpart_base_1.PropertyPaneButton('applyConditionalStyleLookupPick', {
                                    text: 'Use selected value',
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleApplyConditionalStyleLookupPick.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('conditionalStyleLookupHelperResult', {
                                    text: this._conditionalStyleLookupMessage || 'Lookup items load automatically when you select the column.'
                                }),
                            ] : []), [
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('conditionalStyleBackgroundColor', {
                                    label: strings.PropConditionalStyleBackgroundColorLabel,
                                    selectedColor: this.properties.conditionalStyleBackgroundColor || '',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'conditionalStyleBackgroundColorField-' + String(this._conditionalStyleDesignerRevision)
                                }),
                                PropertyFieldColorPicker_1.PropertyFieldColorPicker('conditionalStyleForegroundColor', {
                                    label: strings.PropConditionalStyleForegroundColorLabel,
                                    selectedColor: this.properties.conditionalStyleForegroundColor || '',
                                    onPropertyChange: this.handleColorPropertyChange.bind(this),
                                    properties: this.properties,
                                    style: PropertyFieldColorPicker_1.PropertyFieldColorPickerStyle.Inline,
                                    key: 'conditionalStyleForegroundColorField-' + String(this._conditionalStyleDesignerRevision)
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleFontFamily', {
                                    label: strings.PropConditionalStyleFontFamilyLabel,
                                    options: fontFamilyOptions,
                                    selectedKey: this.properties.conditionalStyleFontFamily || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleFontSize', {
                                    label: strings.PropConditionalStyleFontSizeLabel,
                                    options: [
                                        { key: '', text: 'Default' },
                                        { key: '12px', text: '12px' },
                                        { key: '14px', text: '14px' },
                                        { key: '16px', text: '16px' },
                                        { key: '18px', text: '18px' },
                                        { key: '20px', text: '20px' }
                                    ],
                                    selectedKey: this.properties.conditionalStyleFontSize || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleFontStyle', {
                                    label: strings.PropConditionalStyleFontStyleLabel,
                                    options: [
                                        { key: '', text: 'Default' },
                                        { key: 'normal', text: strings.PropFontStyleNormal },
                                        { key: 'italic', text: strings.PropFontStyleItalic },
                                        { key: 'oblique', text: strings.PropFontStyleOblique }
                                    ],
                                    selectedKey: this.properties.conditionalStyleFontStyle || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleFontWeight', {
                                    label: strings.PropConditionalStyleFontWeightLabel,
                                    options: [
                                        { key: '', text: 'Default' },
                                        { key: 'normal', text: 'Normal' },
                                        { key: 'bold', text: 'Bold' }
                                    ],
                                    selectedKey: this.properties.conditionalStyleFontWeight || ''
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleTextAlign', {
                                    label: strings.PropConditionalStyleTextAlignLabel,
                                    options: [
                                        { key: '', text: 'Default' },
                                        { key: 'left', text: strings.PropTextAlignLeft },
                                        { key: 'center', text: strings.PropTextAlignCenter },
                                        { key: 'right', text: strings.PropTextAlignRight }
                                    ],
                                    selectedKey: this.properties.conditionalStyleTextAlign || ''
                                }),
                                sp_webpart_base_1.PropertyPaneButton('addConditionalStyleCondition', {
                                    text: strings.PropConditionalStyleAddButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Primary,
                                    disabled: this._isEditingConditionalStyle,
                                    onClick: this.handleAddConditionalStyleCondition.bind(this)
                                })
                            ], (this._isEditingConditionalStyle ? [
                                sp_webpart_base_1.PropertyPaneButton('newConditionalStyle', {
                                    text: strings.PropConditionalStyleNewButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleNewConditionalStyle.bind(this)
                                })
                            ] : []), [
                                sp_webpart_base_1.PropertyPaneDropdown('conditionalStyleSelectedIndex', {
                                    label: strings.PropConditionalStyleExistingLabel,
                                    options: conditionalStyleConditionOptions,
                                    selectedKey: this.properties.conditionalStyleSelectedIndex || ''
                                }),
                                sp_webpart_base_1.PropertyPaneButton('loadConditionalStyleCondition', {
                                    text: strings.PropConditionalStyleLoadButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleLoadSelectedConditionalStyleCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('updateConditionalStyleCondition', {
                                    text: strings.PropConditionalStyleUpdateButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    disabled: !this._isEditingConditionalStyle,
                                    onClick: this.handleUpdateSelectedConditionalStyleCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('removeConditionalStyleCondition', {
                                    text: strings.PropConditionalStyleRemoveButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleRemoveSelectedConditionalStyleCondition.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('moveConditionalStyleConditionUp', {
                                    text: strings.PropConditionalStyleMoveUpButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleMoveConditionalStyleConditionUp.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('moveConditionalStyleConditionDown', {
                                    text: strings.PropConditionalStyleMoveDownButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleMoveConditionalStyleConditionDown.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneButton('resetConditionalStyleJson', {
                                    text: strings.PropConditionalStyleResetButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    disabled: this._isEditingConditionalStyle,
                                    onClick: this.handleResetConditionalStyleJson.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('conditionalStyleDesignerResult', {
                                    text: this._conditionalStyleDesignerMessage || strings.PropConditionalStyleStatusPlaceholder
                                })
                            ])
                        },
                        {
                            groupName: strings.ConditionalStyleJsonGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneTextField('conditionalStyleJson', {
                                    label: strings.PropConditionalStyleJsonLabel,
                                    placeholder: strings.PropConditionalStyleJsonPlaceholder,
                                    value: this.properties.conditionalStyleJson || '',
                                    multiline: true,
                                    resizable: true,
                                    rows: 12,
                                    description: strings.PropConditionalStyleJsonHelp
                                }),
                                sp_webpart_base_1.PropertyPaneButton('validateConditionalStyleJson', {
                                    text: strings.PropValidateConditionalStyleJsonButtonLabel,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Normal,
                                    onClick: this.handleValidateConditionalStyleJson.bind(this)
                                }),
                                sp_webpart_base_1.PropertyPaneLabel('validateConditionalStyleJsonResult', {
                                    text: this._conditionalStyleJsonValidationMessage || strings.ConditionalStyleJsonValidationResultPlaceholder
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    GridControlWebPart.prototype.getWebPartVersion = function () {
        var solutionVersion = packageSolutionConfig && packageSolutionConfig.solution
            ? String(packageSolutionConfig.solution.version || '')
            : '';
        if (solutionVersion) {
            return solutionVersion;
        }
        var manifestVersion = this.context && this.context.manifest ? String(this.context.manifest.version || '') : '';
        if (manifestVersion && manifestVersion !== '*') {
            return manifestVersion;
        }
        return 'Unknown';
    };
    GridControlWebPart.prototype.openGridDesigner = function () {
        if (!this.properties.listName) {
            return;
        }
        this._isGridDesignerOpen = true;
        this.render();
    };
    GridControlWebPart.prototype.closeGridDesigner = function () {
        this._isGridDesignerOpen = false;
        this.render();
    };
    GridControlWebPart.prototype.saveGridDesign = function (schemaJson) {
        this.properties.gridSchemaJson = schemaJson;
        this.properties.viewColumns = [];
        this._isGridDesignerOpen = false;
        this.context.propertyPane.refresh();
        this.render();
    };
    GridControlWebPart.prototype.getGridDesignerStatus = function () {
        var json = String(this.properties.gridSchemaJson || '').trim();
        if (!json) {
            return strings.PropGridDesignerNotConfigured;
        }
        try {
            var schema = JSON.parse(json);
            var steps = schema && Array.isArray(schema.steps) ? schema.steps : [];
            var count = 0;
            for (var i = 0; i < steps.length; i += 1) {
                count += steps[i] && Array.isArray(steps[i].fields) ? steps[i].fields.length : 0;
            }
            return strings.PropGridDesignerConfigured.replace('{0}', String(count));
        }
        catch (_error) {
            return strings.PropGridDesignerInvalid;
        }
    };
    GridControlWebPart.prototype.getFontFamilyOptions = function () {
        return [
            { key: '', text: 'Default' },
            { key: 'Segoe UI', text: 'Segoe UI' },
            { key: 'Calibri', text: 'Calibri' },
            { key: 'Arial', text: 'Arial' },
            { key: 'Tahoma', text: 'Tahoma' },
            { key: 'Verdana', text: 'Verdana' },
            { key: 'Georgia', text: 'Georgia' },
            { key: '"Times New Roman"', text: 'Times New Roman' },
            { key: '"Courier New"', text: 'Courier New' }
        ];
    };
    GridControlWebPart.prototype.getFontStyleOptions = function () {
        return [
            { key: 'normal', text: strings.PropFontStyleNormal },
            { key: 'italic', text: strings.PropFontStyleItalic },
            { key: 'oblique', text: strings.PropFontStyleOblique },
            { key: 'initial', text: strings.PropFontStyleInitial },
            { key: 'inherit', text: strings.PropFontStyleInherit },
            { key: 'unset', text: strings.PropFontStyleUnset }
        ];
    };
    GridControlWebPart.prototype.getTextAlignOptions = function () {
        return [
            { key: 'left', text: strings.PropTextAlignLeft },
            { key: 'center', text: strings.PropTextAlignCenter },
            { key: 'right', text: strings.PropTextAlignRight }
        ];
    };
    GridControlWebPart.prototype.validateJsonText = function (raw, expectedShape) {
        var source = String(raw || '').trim();
        if (!source) {
            return {
                valid: true,
                message: strings.JsonValidationEmpty
            };
        }
        try {
            var parsed = JSON.parse(source);
            if (expectedShape === 'array' && !Array.isArray(parsed)) {
                return {
                    valid: false,
                    message: strings.JsonValidationExpectedArray
                };
            }
            if (expectedShape === 'object' && (!parsed || typeof parsed !== 'object' || Array.isArray(parsed))) {
                return {
                    valid: false,
                    message: strings.JsonValidationExpectedObject
                };
            }
            return {
                valid: true,
                message: strings.JsonValidationSuccess
            };
        }
        catch (error) {
            return {
                valid: false,
                message: strings.JsonValidationFailedPrefix + ': ' + (error && error.message ? error.message : '')
            };
        }
    };
    GridControlWebPart.prototype.parseFilterJsonArray = function (raw) {
        var source = String(raw || '').trim();
        if (!source) {
            return [];
        }
        var parsed = JSON.parse(source);
        if (!Array.isArray(parsed)) {
            throw new Error(strings.JsonValidationExpectedArray);
        }
        return parsed;
    };
    GridControlWebPart.prototype.handleValidateFilterJson = function () {
        var result = this.validateJsonText(this.properties.filterJson, 'array');
        this._filterJsonValidationMessage = result.message;
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleAddFilterDesignerCondition = function () {
        try {
            var field = String(this.properties.filterDesignerField || '').trim();
            if (!field) {
                this._filterDesignerMessage = strings.PropFilterDesignerFieldRequired;
                this.context.propertyPane.refresh();
                return;
            }
            var operator = String(this.properties.filterDesignerOperator || 'contains').trim().toLowerCase();
            var logical = String(this.properties.filterDesignerLogical || 'and').trim().toLowerCase();
            var valueType = this.properties.filterDesignerValueType === 'expression' ? 'expression' : (this.properties.filterDesignerValueType === 'fieldValue' ? 'fieldValue' : 'static');
            var value = String(this.properties.filterDesignerValue || '').trim();
            if (!value) {
                this._filterDesignerMessage = strings.PropFilterDesignerValueRequired;
                this.context.propertyPane.refresh();
                return;
            }
            if (valueType === 'expression' && !this.isSupportedFilterExpression(value)) {
                this._filterDesignerMessage = strings.PropFilterDesignerExpressionInvalid;
                this.context.propertyPane.refresh();
                return;
            }
            var persistedValue = valueType === 'fieldValue' && this.isMultiFilterField(field) ? this.getSelectedFilterFieldValues() : value;
            var existing = this.parseFilterJsonArray(this.properties.filterJson);
            existing.push({
                field: field,
                operator: operator,
                logical: logical,
                valueType: valueType,
                value: persistedValue
            });
            this.properties.filterJson = JSON.stringify(existing, null, 2);
            this._filterDesignerMessage = strings.PropFilterDesignerAddSuccess;
            this._filterJsonValidationMessage = strings.JsonValidationSuccess;
        }
        catch (error) {
            this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
        }
        this.context.propertyPane.refresh();
        this.render();
    };
    GridControlWebPart.prototype.handleLoadSelectedFilterCondition = function () {
        try {
            var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
            var existing = this.parseFilterJsonArray(this.properties.filterJson);
            if (isNaN(index) || index < 0 || index >= existing.length) {
                this._filterDesignerMessage = strings.PropFilterDesignerSelectValid;
                this.context.propertyPane.refresh();
                return;
            }
            var selected = existing[index] || {};
            this.properties.filterDesignerField = String(selected.field || '');
            this.properties.filterDesignerOperator = String(selected.operator || 'contains').toLowerCase();
            this.properties.filterDesignerLogical = String(selected.logical || 'and').toLowerCase();
            var storedValueType = String(selected.valueType || '').toLowerCase();
            this.properties.filterDesignerValueType = storedValueType === 'expression' ? 'expression' : (storedValueType === 'fieldvalue' ? 'fieldValue' : 'static');
            this.properties.filterDesignerValue = Array.isArray(selected.value) ? JSON.stringify(selected.value) : String(selected.value === undefined || selected.value === null ? '' : selected.value);
            this._filterDesignerMessage = strings.PropFilterDesignerLoadSuccess;
        }
        catch (error) {
            this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleUpdateSelectedFilterCondition = function () {
        try {
            var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
            var existing = this.parseFilterJsonArray(this.properties.filterJson);
            if (isNaN(index) || index < 0 || index >= existing.length) {
                this._filterDesignerMessage = strings.PropFilterDesignerSelectValid;
                this.context.propertyPane.refresh();
                return;
            }
            var field = String(this.properties.filterDesignerField || '').trim();
            if (!field) {
                this._filterDesignerMessage = strings.PropFilterDesignerFieldRequired;
                this.context.propertyPane.refresh();
                return;
            }
            var value = String(this.properties.filterDesignerValue || '').trim();
            if (!value) {
                this._filterDesignerMessage = strings.PropFilterDesignerValueRequired;
                this.context.propertyPane.refresh();
                return;
            }
            var valueType = this.properties.filterDesignerValueType === 'expression' ? 'expression' : (this.properties.filterDesignerValueType === 'fieldValue' ? 'fieldValue' : 'static');
            if (valueType === 'expression' && !this.isSupportedFilterExpression(value)) {
                this._filterDesignerMessage = strings.PropFilterDesignerExpressionInvalid;
                this.context.propertyPane.refresh();
                return;
            }
            var persistedValue = valueType === 'fieldValue' && this.isMultiFilterField(field) ? this.getSelectedFilterFieldValues() : value;
            existing[index] = {
                field: field,
                operator: String(this.properties.filterDesignerOperator || 'contains').trim().toLowerCase(),
                logical: String(this.properties.filterDesignerLogical || 'and').trim().toLowerCase(),
                valueType: valueType,
                value: persistedValue
            };
            this.properties.filterJson = JSON.stringify(existing, null, 2);
            this._filterDesignerMessage = strings.PropFilterDesignerUpdateSuccess;
            this._filterJsonValidationMessage = strings.JsonValidationSuccess;
        }
        catch (error) {
            this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
        }
        this.context.propertyPane.refresh();
        this.render();
    };
    GridControlWebPart.prototype.handleRemoveSelectedFilterCondition = function () {
        try {
            var index = parseInt(String(this.properties.filterDesignerSelectedIndex || ''), 10);
            var existing = this.parseFilterJsonArray(this.properties.filterJson);
            if (isNaN(index) || index < 0 || index >= existing.length) {
                this._filterDesignerMessage = strings.PropFilterDesignerSelectValid;
                this.context.propertyPane.refresh();
                return;
            }
            existing.splice(index, 1);
            this.properties.filterJson = JSON.stringify(existing, null, 2);
            this.properties.filterDesignerSelectedIndex = '';
            this._filterDesignerMessage = strings.PropFilterDesignerRemoveSuccess;
            this._filterJsonValidationMessage = existing.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
        }
        catch (error) {
            this._filterDesignerMessage = error && error.message ? error.message : strings.PropFilterDesignerActionFailed;
        }
        this.context.propertyPane.refresh();
        this.render();
    };
    GridControlWebPart.prototype.handleResetFilterJson = function () {
        this.properties.filterJson = '';
        this.properties.filterDesignerSelectedIndex = '';
        this._filterJsonValidationMessage = strings.JsonValidationEmpty;
        this._filterDesignerMessage = strings.PropFilterDesignerResetSuccess;
        this.context.propertyPane.refresh();
        this.render();
    };
    GridControlWebPart.prototype.isSupportedFilterExpression = function (value) {
        var normalized = String(value || '').trim().toLowerCase();
        return normalized === 'today'
            || normalized === 'now'
            || normalized === 'me'
            || normalized === 'me.email'
            || normalized === 'me.login'
            || normalized === 'me.id'
            || /^date\(\s*[+-]?\d+\s*\)$/.test(normalized);
    };
    GridControlWebPart.prototype.parseConditionalStyleJsonArray = function (raw) {
        var source = String(raw || '').trim();
        if (!source) {
            return [];
        }
        var parsed = JSON.parse(source);
        if (!Array.isArray(parsed)) {
            throw new Error(strings.JsonValidationExpectedArray);
        }
        return parsed;
    };
    GridControlWebPart.prototype.handleValidateConditionalStyleJson = function () {
        var result = this.validateJsonText(this.properties.conditionalStyleJson, 'array');
        this._conditionalStyleJsonValidationMessage = result.message;
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.parseConditionalStylePriority = function (raw) {
        var parsed = parseInt(String(raw || '').trim(), 10);
        if (isNaN(parsed)) {
            return 100;
        }
        return parsed;
    };
    GridControlWebPart.prototype.createConditionalStyleEntryFromDesigner = function () {
        var ruleName = String(this.properties.conditionalStyleRuleName || '').trim();
        if (!ruleName) {
            throw new Error(strings.PropConditionalStyleRuleNameRequired);
        }
        var scope = String(this.properties.conditionalStyleScope || 'row').trim().toLowerCase();
        if (scope !== 'row' && scope !== 'column') {
            scope = 'row';
        }
        var applyField = String(this.properties.conditionalStyleApplyField || '').trim();
        if (scope === 'column' && !applyField) {
            throw new Error(strings.PropConditionalStyleApplyFieldRequired);
        }
        var conditionField = String(this.properties.conditionalStyleConditionField || '').trim();
        if (!conditionField) {
            throw new Error(strings.PropConditionalStyleConditionFieldRequired);
        }
        var conditionValue = String(this.properties.conditionalStyleConditionValue || '').trim();
        if (!conditionValue) {
            throw new Error(strings.PropConditionalStyleConditionValueRequired);
        }
        var conditionValueType = this.properties.conditionalStyleConditionValueType === 'expression' ? 'expression' : 'static';
        if (conditionValueType === 'expression' && !this.isSupportedFilterExpression(conditionValue)) {
            throw new Error(strings.PropFilterDesignerExpressionInvalid);
        }
        return {
            rule: ruleName,
            scope: scope,
            applyField: applyField,
            enabled: this.properties.conditionalStyleEnabled !== false,
            priority: this.parseConditionalStylePriority(this.properties.conditionalStylePriority),
            conditionField: conditionField,
            operator: String(this.properties.conditionalStyleConditionOperator || 'eq').trim().toLowerCase(),
            logical: String(this.properties.conditionalStyleConditionLogical || 'and').trim().toLowerCase(),
            valueType: conditionValueType,
            value: conditionValue,
            style: {
                backgroundColor: String(this.properties.conditionalStyleBackgroundColor || '').trim(),
                color: String(this.properties.conditionalStyleForegroundColor || '').trim(),
                fontFamily: String(this.properties.conditionalStyleFontFamily || '').trim(),
                fontSize: String(this.properties.conditionalStyleFontSize || '').trim(),
                fontStyle: String(this.properties.conditionalStyleFontStyle || '').trim(),
                fontWeight: String(this.properties.conditionalStyleFontWeight || '').trim(),
                textAlign: String(this.properties.conditionalStyleTextAlign || '').trim()
            }
        };
    };
    GridControlWebPart.prototype.resetConditionalStyleDesigner = function () {
        this.properties.conditionalStyleRuleName = '';
        this.properties.conditionalStyleScope = 'row';
        this.properties.conditionalStyleApplyField = '';
        this.properties.conditionalStyleConditionField = '';
        this.properties.conditionalStyleConditionOperator = 'eq';
        this.properties.conditionalStyleConditionLogical = 'and';
        this.properties.conditionalStyleConditionValueType = 'static';
        this.properties.conditionalStyleEnabled = true;
        this.properties.conditionalStylePriority = '100';
        this.properties.conditionalStyleConditionValue = '';
        this.properties.conditionalStyleBackgroundColor = '';
        this.properties.conditionalStyleForegroundColor = '';
        this.properties.conditionalStyleFontFamily = '';
        this.properties.conditionalStyleFontSize = '';
        this.properties.conditionalStyleFontStyle = '';
        this.properties.conditionalStyleFontWeight = '';
        this.properties.conditionalStyleTextAlign = '';
        this.properties.conditionalStyleSelectedIndex = '';
        this._conditionalStyleDesignerRevision += 1;
    };
    GridControlWebPart.prototype.handleNewConditionalStyle = function () {
        this._isEditingConditionalStyle = false;
        this.resetConditionalStyleDesigner();
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleNewSuccess;
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleAddConditionalStyleCondition = function () {
        try {
            var nextEntry = this.createConditionalStyleEntryFromDesigner();
            var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            existing.push(nextEntry);
            this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
            this.resetConditionalStyleDesigner();
            this._conditionalStyleDesignerMessage = strings.PropConditionalStyleAddSuccess;
            this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
        }
        catch (error) {
            this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleLoadSelectedConditionalStyleCondition = function () {
        try {
            var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
            var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            if (isNaN(index) || index < 0 || index >= existing.length) {
                this._conditionalStyleDesignerMessage = strings.PropConditionalStyleSelectValid;
                this.context.propertyPane.refresh();
                return;
            }
            var selected = existing[index] || {};
            var selectedStyle = selected.style || {};
            this.properties.conditionalStyleRuleName = String(selected.rule || '');
            this.properties.conditionalStyleScope = String(selected.scope || 'row').toLowerCase();
            this.properties.conditionalStyleApplyField = String(selected.applyField || '');
            this.properties.conditionalStyleConditionField = String(selected.conditionField || '');
            this.properties.conditionalStyleConditionOperator = String(selected.operator || 'eq').toLowerCase();
            this.properties.conditionalStyleConditionLogical = String(selected.logical || 'and').toLowerCase();
            this.properties.conditionalStyleConditionValueType = String(selected.valueType || '').toLowerCase() === 'expression' ? 'expression' : 'static';
            this.properties.conditionalStyleEnabled = selected.enabled !== false;
            this.properties.conditionalStylePriority = String(selected.priority === undefined || selected.priority === null ? 100 : selected.priority);
            this.properties.conditionalStyleConditionValue = String(selected.value === undefined || selected.value === null ? '' : selected.value);
            this.properties.conditionalStyleBackgroundColor = String(selectedStyle.backgroundColor || '');
            this.properties.conditionalStyleForegroundColor = String(selectedStyle.color || '');
            this.properties.conditionalStyleFontFamily = String(selectedStyle.fontFamily || '');
            this.properties.conditionalStyleFontSize = String(selectedStyle.fontSize || '');
            this.properties.conditionalStyleFontStyle = String(selectedStyle.fontStyle || '');
            this.properties.conditionalStyleFontWeight = String(selectedStyle.fontWeight || '');
            this.properties.conditionalStyleTextAlign = String(selectedStyle.textAlign || '');
            this._isEditingConditionalStyle = true;
            this._conditionalStyleDesignerRevision += 1;
            this._conditionalStyleDesignerMessage = strings.PropConditionalStyleLoadSuccess;
        }
        catch (error) {
            this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleUpdateSelectedConditionalStyleCondition = function () {
        try {
            var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
            var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            if (isNaN(index) || index < 0 || index >= existing.length) {
                this._conditionalStyleDesignerMessage = strings.PropConditionalStyleSelectValid;
                this.context.propertyPane.refresh();
                return;
            }
            existing[index] = this.createConditionalStyleEntryFromDesigner();
            this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
            this._conditionalStyleDesignerMessage = strings.PropConditionalStyleUpdateSuccess;
            this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
        }
        catch (error) {
            this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleRemoveSelectedConditionalStyleCondition = function () {
        try {
            var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
            var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            if (isNaN(index) || index < 0 || index >= existing.length) {
                this._conditionalStyleDesignerMessage = strings.PropConditionalStyleSelectValid;
                this.context.propertyPane.refresh();
                return;
            }
            existing.splice(index, 1);
            this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
            this._isEditingConditionalStyle = false;
            this.resetConditionalStyleDesigner();
            this._conditionalStyleDesignerMessage = strings.PropConditionalStyleRemoveSuccess;
            this._conditionalStyleJsonValidationMessage = existing.length === 0 ? strings.JsonValidationEmpty : strings.JsonValidationSuccess;
        }
        catch (error) {
            this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleMoveConditionalStyleConditionUp = function () {
        try {
            var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
            var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            if (isNaN(index) || index <= 0 || index >= existing.length) {
                this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveBoundary;
                this.context.propertyPane.refresh();
                return;
            }
            var current = existing[index];
            existing[index] = existing[index - 1];
            existing[index - 1] = current;
            this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
            this.properties.conditionalStyleSelectedIndex = String(index - 1);
            this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveUpSuccess;
            this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
        }
        catch (error) {
            this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleMoveConditionalStyleConditionDown = function () {
        try {
            var index = parseInt(String(this.properties.conditionalStyleSelectedIndex || ''), 10);
            var existing = this.parseConditionalStyleJsonArray(this.properties.conditionalStyleJson);
            if (isNaN(index) || index < 0 || index >= existing.length - 1) {
                this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveBoundary;
                this.context.propertyPane.refresh();
                return;
            }
            var current = existing[index];
            existing[index] = existing[index + 1];
            existing[index + 1] = current;
            this.properties.conditionalStyleJson = JSON.stringify(existing, null, 2);
            this.properties.conditionalStyleSelectedIndex = String(index + 1);
            this._conditionalStyleDesignerMessage = strings.PropConditionalStyleMoveDownSuccess;
            this._conditionalStyleJsonValidationMessage = strings.JsonValidationSuccess;
        }
        catch (error) {
            this._conditionalStyleDesignerMessage = error && error.message ? error.message : strings.PropConditionalStyleActionFailed;
        }
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.handleResetConditionalStyleJson = function () {
        this.properties.conditionalStyleJson = '';
        this._isEditingConditionalStyle = false;
        this.resetConditionalStyleDesigner();
        this._conditionalStyleJsonValidationMessage = strings.JsonValidationEmpty;
        this._conditionalStyleDesignerMessage = strings.PropConditionalStyleResetSuccess;
        this.context.propertyPane.refresh();
    };
    GridControlWebPart.prototype.initializeDynamicDataSource = function () {
        this._dynamicDataSourceManager = this.context.dynamicDataSourceManager || this.context._dynamicDataSourceManager;
        this.logDiagnostic('Initializing dynamic data source manager. Available=' + String(!!this._dynamicDataSourceManager));
        if (this._dynamicDataSourceManager && this._dynamicDataSourceManager.initializeSource) {
            this._dynamicDataSourceManager.initializeSource(this);
        }
        else if (this._dynamicDataSourceManager && this._dynamicDataSourceManager.registerSource) {
            this._dynamicDataSourceManager.registerSource(this);
        }
    };
    GridControlWebPart.prototype.notifyDynamicSourceChanged = function () {
        if (!this._dynamicDataSourceManager) {
            return;
        }
        if (this._dynamicDataSourceManager.notifySourceChanged) {
            this._dynamicDataSourceManager.notifySourceChanged();
        }
        else if (this._dynamicDataSourceManager.notifyDataChanged) {
            this._dynamicDataSourceManager.notifyDataChanged();
        }
    };
    GridControlWebPart.prototype.notifyDynamicData = function (propertyId) {
        if (!this._dynamicDataSourceManager) {
            return;
        }
        var hasNotified = false;
        if (this._dynamicDataSourceManager.notifyPropertyChanged) {
            this._dynamicDataSourceManager.notifyPropertyChanged(propertyId);
            hasNotified = true;
        }
        if (this._dynamicDataSourceManager.notifySourceChanged) {
            this._dynamicDataSourceManager.notifySourceChanged();
            hasNotified = true;
        }
        if (!hasNotified && this._dynamicDataSourceManager.notifyDataChanged) {
            this._dynamicDataSourceManager.notifyDataChanged();
        }
    };
    GridControlWebPart.prototype.handleSelectionChange = function (itemId, mode) {
        this._selectedItemId = itemId > 0 ? itemId : 0;
        this._selectedMode = mode || 'view';
        this.logDiagnostic('Selection changed. itemId=' + String(this._selectedItemId) + ', mode=' + String(this._selectedMode));
        this.notifyDynamicData('selectedItemId');
        this.notifyDynamicData('selectedMode');
        this.render();
    };
    GridControlWebPart.prototype.getJsonWithAcceptFallback = function (url) {
        return __awaiter(this, void 0, void 0, function () {
            var response;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=verbose'
                                }
                            })];
                    case 2:
                        response = _a.sent();
                        _a.label = 3;
                    case 3:
                        if (!!response.ok) return [3 /*break*/, 5];
                        return [4 /*yield*/, this.context.spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata'
                                }
                            })];
                    case 4:
                        response = _a.sent();
                        _a.label = 5;
                    case 5:
                        if (!response.ok) {
                            throw new Error('Request failed. HTTP ' + String(response.status) + ' ' + response.statusText);
                        }
                        return [2 /*return*/, response.json()];
                }
            });
        });
    };
    GridControlWebPart.prototype.loadLists = function () {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, data, lists, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
                        this.logDiagnostic('Loading lists from web URL: ' + webUrl);
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(webUrl + "/_api/web/lists?$select=Title,Hidden,BaseTemplate&$filter=Hidden eq false")];
                    case 1:
                        data = _a.sent();
                        lists = data && data.value ? data.value : [];
                        if (!lists || lists.length === 0) {
                            lists = data && data.d && data.d.results ? data.d.results : [];
                        }
                        this._lists = lists.map(function (list) {
                            return {
                                key: list.Title,
                                text: list.Title
                            };
                        });
                        this.logDiagnostic('Loaded lists successfully. Count=' + String(this._lists.length));
                        this.context.propertyPane.refresh();
                        return [3 /*break*/, 3];
                    case 2:
                        error_3 = _a.sent();
                        this.logDiagnostic('Failed to load lists: ' + (error_3 && error_3.message ? error_3.message : String(error_3)));
                        this._lists = [];
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.getTargetPageOptions = function () {
        var options = this._sitePages.length > 0
            ? this._sitePages.slice()
            : [{ key: '__defaultForm__', text: strings.PropLinkTargetDefaultFormOption }];
        var configuredUrl = String(this.properties.linkTargetPageUrl || '').trim();
        if (configuredUrl && configuredUrl !== '__defaultForm__'
            && !options.some(function (option) { return String(option.key) === configuredUrl; })) {
            options.push({ key: configuredUrl, text: configuredUrl + ' (' + strings.PropLinkTargetSavedUrlLabel + ')' });
        }
        return options;
    };
    GridControlWebPart.prototype.loadSitePages = function () {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, libraryData, libraries, pageOptionsByUrl, libraryIndex, libraryId, data, pages, pageError_1, defaultPageOptions, error_4;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(webUrl + '/_api/web/lists?$select=Id,Title,BaseTemplate&$filter=(BaseTemplate eq 119 or BaseTemplate eq 850)')];
                    case 1:
                        libraryData = _a.sent();
                        libraries = libraryData && libraryData.value ? libraryData.value
                            : (libraryData && libraryData.d && libraryData.d.results ? libraryData.d.results : []);
                        pageOptionsByUrl = {};
                        libraryIndex = 0;
                        _a.label = 2;
                    case 2:
                        if (!(libraryIndex < libraries.length)) return [3 /*break*/, 7];
                        libraryId = String(libraries[libraryIndex].Id || '').replace(/[{}]/g, '');
                        if (!libraryId) {
                            return [3 /*break*/, 6];
                        }
                        _a.label = 3;
                    case 3:
                        _a.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(webUrl + "/_api/web/lists(guid'" + libraryId
                                + "')/items?$select=File/Name,File/ServerRelativeUrl&$expand=File&$top=5000")];
                    case 4:
                        data = _a.sent();
                        pages = data && data.value ? data.value
                            : (data && data.d && data.d.results ? data.d.results : []);
                        pages.filter(function (page) {
                            return page.File && page.File.ServerRelativeUrl
                                && /\.aspx(?:$|[?#])/i.test(String(page.File.ServerRelativeUrl));
                        }).forEach(function (page) {
                            var fileRef = String(page.File.ServerRelativeUrl);
                            var title = String(page.File.Name || fileRef);
                            pageOptionsByUrl[fileRef.toLowerCase()] = { key: fileRef, text: title + ' (' + fileRef + ')' };
                        });
                        return [3 /*break*/, 6];
                    case 5:
                        pageError_1 = _a.sent();
                        this.logDiagnostic('Failed to load pages from library "' + String(libraries[libraryIndex].Title || '') + '": '
                            + String(pageError_1 && pageError_1.message ? pageError_1.message : pageError_1));
                        return [3 /*break*/, 6];
                    case 6:
                        libraryIndex += 1;
                        return [3 /*break*/, 2];
                    case 7:
                        defaultPageOptions = [
                            { key: '__defaultForm__', text: strings.PropLinkTargetDefaultFormOption }
                        ];
                        this._sitePages = defaultPageOptions
                            .concat(Object.keys(pageOptionsByUrl).map(function (url) { return pageOptionsByUrl[url]; })
                            .sort(function (left, right) {
                            return String(left.text).localeCompare(String(right.text));
                        }));
                        this.context.propertyPane.refresh();
                        return [3 /*break*/, 9];
                    case 8:
                        error_4 = _a.sent();
                        this.logDiagnostic('Failed to load Site Pages: ' + (error_4 && error_4.message ? error_4.message : String(error_4)));
                        this._sitePages = [{ key: '__defaultForm__', text: strings.PropLinkTargetDefaultFormOption }];
                        this.context.propertyPane.refresh();
                        return [3 /*break*/, 9];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.loadViews = function (listName) {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, data, views, i, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
                        this.logDiagnostic('Loading views for list: ' + String(listName));
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(webUrl + "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')/views?$select=Id,Title,DefaultView")];
                    case 1:
                        data = _a.sent();
                        views = data && data.value ? data.value : [];
                        if (!views || views.length === 0) {
                            views = data && data.d && data.d.results ? data.d.results : [];
                        }
                        this._views = views.map(function (view) {
                            return {
                                key: String(view.Id),
                                text: view.Title,
                                isDefault: view.DefaultView === true
                            };
                        });
                        if (!this.properties.viewId && this._views.length > 0) {
                            for (i = 0; i < this._views.length; i += 1) {
                                if (this._views[i].isDefault) {
                                    this.properties.viewId = String(this._views[i].key);
                                    break;
                                }
                            }
                            if (!this.properties.viewId) {
                                this.properties.viewId = String(this._views[0].key);
                            }
                            this.logDiagnostic('Auto-selected default viewId=' + String(this.properties.viewId));
                        }
                        if (!(this.properties.viewId && (!this.properties.viewColumns || this.properties.viewColumns.length === 0))) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.loadViewColumns(listName, this.properties.viewId)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        this.logDiagnostic('Loaded views successfully. Count=' + String(this._views.length));
                        this.context.propertyPane.refresh();
                        this.render();
                        return [3 /*break*/, 5];
                    case 4:
                        error_5 = _a.sent();
                        this.logDiagnostic('Failed to load views for list ' + String(listName) + ': ' + (error_5 && error_5.message ? error_5.message : String(error_5)));
                        this._views = [];
                        this.context.propertyPane.refresh();
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.loadViewColumns = function (listName, viewId) {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, listPath, normalizedViewId, viewData, viewFields, fieldData, fields, titleByName, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!listName || !viewId) {
                            this.properties.viewColumns = [];
                            this.context.propertyPane.refresh();
                            this.render();
                            return [2 /*return*/];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
                        listPath = "/_api/web/lists/getByTitle('" + escapeODataText(listName) + "')";
                        normalizedViewId = String(viewId).replace(/[{}]/g, '');
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(webUrl + listPath + "/views/getById('" + encodeURIComponent(normalizedViewId) + "')/ViewFields")];
                    case 2:
                        viewData = _a.sent();
                        viewFields = normalizeStringCollection(viewData && viewData.value);
                        if (viewFields.length === 0) {
                            viewFields = normalizeStringCollection(viewData && viewData.Items);
                        }
                        if (viewFields.length === 0) {
                            viewFields = normalizeStringCollection(viewData && viewData.d && viewData.d.Items);
                        }
                        return [4 /*yield*/, this.getJsonWithAcceptFallback(webUrl + listPath + "/fields?$select=InternalName,Title")];
                    case 3:
                        fieldData = _a.sent();
                        fields = fieldData && fieldData.value ? fieldData.value : [];
                        if (!fields || fields.length === 0) {
                            fields = fieldData && fieldData.d && fieldData.d.results ? fieldData.d.results : [];
                        }
                        titleByName = {};
                        fields.forEach(function (field) {
                            var internalName = String(field.InternalName || '');
                            if (internalName) {
                                titleByName[internalName] = String(field.Title || internalName);
                            }
                        });
                        if (this.properties.listName !== listName || this.properties.viewId !== viewId) {
                            return [2 /*return*/];
                        }
                        this.properties.viewColumns = viewFields.map(function (fieldName) {
                            return {
                                fieldName: fieldName,
                                displayName: titleByName[fieldName] || fieldName,
                                width: ''
                            };
                        });
                        this.logDiagnostic('Initialized view column collection. Count=' + String(this.properties.viewColumns.length));
                        this.context.propertyPane.refresh();
                        this.render();
                        return [3 /*break*/, 5];
                    case 4:
                        error_6 = _a.sent();
                        this.properties.viewColumns = [];
                        this.context.propertyPane.refresh();
                        this.render();
                        this.logDiagnostic('Failed to initialize columns for view ' + viewId + ': ' + (error_6 && error_6.message ? error_6.message : String(error_6)));
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    GridControlWebPart.prototype.logDiagnostic = function (message) {
        if (this.properties.enableDiagnostics === false) {
            return;
        }
        console.log('[GridControlWebPart] ' + message);
    };
    return GridControlWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = GridControlWebPart;

//# sourceMappingURL=GridControlWebPart.js.map
