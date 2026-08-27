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
var GridValidation_1 = require("./GridValidation");
require("./GridDesigner.css");
var SYSTEM_FIELDS = {
    ID: true,
    Created: true,
    Modified: true,
    Author: true,
    Editor: true,
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
    return 'gridfield_' + new Date().getTime().toString(36) + '_' + Math.floor(Math.random() * 100000).toString(36);
}
function escapeODataText(value) {
    return value.replace(/'/g, "''");
}
function copyFields(fields) {
    return JSON.parse(JSON.stringify(fields || []));
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
function mapFieldType(type) {
    switch (type) {
        case 'Note': return 'multiline';
        case 'Number':
        case 'Currency':
        case 'Integer': return 'number';
        case 'DateTime': return 'datetime';
        case 'Choice': return 'dropdown';
        case 'MultiChoice': return 'multiselect';
        case 'Lookup':
        case 'LookupMulti': return 'lookup';
        case 'User':
        case 'UserMulti': return 'person';
        case 'Boolean': return 'boolean';
        case 'URL':
        case 'Hyperlink': return 'url';
        case 'Image':
        case 'Thumbnail': return 'image';
        case 'TaxonomyFieldType':
        case 'TaxonomyFieldTypeMulti': return 'taxonomy';
        case 'Attachments': return 'attachment';
        default: return 'text';
    }
}
function getCompatibleControlTypes(sharePointType) {
    switch (sharePointType) {
        case 'Text':
        case 'Note':
            return ['text', 'multiline', 'number'];
        case 'Number':
        case 'Currency':
        case 'Integer':
            return ['number'];
        case 'Choice': return ['dropdown'];
        case 'MultiChoice': return ['multiselect'];
        case 'DateTime': return ['datetime'];
        case 'Boolean': return ['boolean'];
        case 'URL':
        case 'Hyperlink':
            return ['url', 'text'];
        case 'Lookup':
        case 'LookupMulti':
            return ['lookup'];
        case 'User':
        case 'UserMulti':
            return ['person'];
        case 'TaxonomyFieldType':
        case 'TaxonomyFieldTypeMulti':
            return ['taxonomy'];
        case 'Image':
        case 'Thumbnail':
            return ['image'];
        case 'Attachments': return ['attachment'];
        default: return [mapFieldType(sharePointType)];
    }
}
function isCompatibleControlType(sharePointType, controlType) {
    return getCompatibleControlTypes(sharePointType).indexOf(controlType) >= 0;
}
function getTypeLabel(type) {
    var labels = {
        text: 'Text', multiline: 'Multiline text', number: 'Number', datetime: 'Date and time',
        dropdown: 'Choice', multiselect: 'Multiple choice', lookup: 'Lookup', person: 'Person',
        boolean: 'Yes/No', url: 'URL', image: 'Image', taxonomy: 'Managed metadata', attachment: 'Attachments'
    };
    return labels[type] || type;
}
function parseSchemaFields(schemaJson) {
    if (!String(schemaJson || '').trim()) {
        return [];
    }
    try {
        var schema = JSON.parse(schemaJson);
        var steps = schema && Array.isArray(schema.steps) ? schema.steps : [];
        var fields = [];
        for (var stepIndex = 0; stepIndex < steps.length; stepIndex += 1) {
            var stepFields = steps[stepIndex] && Array.isArray(steps[stepIndex].fields) ? steps[stepIndex].fields : [];
            for (var fieldIndex = 0; fieldIndex < stepFields.length; fieldIndex += 1) {
                if (stepFields[fieldIndex] && stepFields[fieldIndex].fieldName) {
                    var parsedField = stepFields[fieldIndex];
                    if (parsedField.disabled === true) {
                        parsedField.readOnly = true;
                        delete parsedField.disabled;
                    }
                    fields.push(parsedField);
                }
            }
        }
        return fields;
    }
    catch (_error) {
        return [];
    }
}
function parseAdvancedValidation(schemaJson) {
    if (!String(schemaJson || '').trim()) {
        return { enabled: false, rules: [] };
    }
    try {
        var schema = JSON.parse(schemaJson);
        var config = schema && schema.advancedValidation;
        return {
            enabled: !!(config && config.enabled === true),
            rules: config && Array.isArray(config.rules) ? config.rules : []
        };
    }
    catch (_error) {
        return { enabled: false, rules: [] };
    }
}
var GridDesigner = (function (_super) {
    __extends(GridDesigner, _super);
    function GridDesigner(props) {
        var _this = _super.call(this, props) || this;
        var initialFields = parseSchemaFields(props.schemaJson);
        var initialAdvancedValidation = parseAdvancedValidation(props.schemaJson);
        _this.state = {
            fields: initialFields,
            sharePointFields: [],
            selectedFieldId: initialFields.length > 0 ? initialFields[0].id : '',
            loading: false,
            error: '',
            advancedValidationEnabled: initialAdvancedValidation.enabled,
            advancedValidationRules: initialAdvancedValidation.rules,
            validationExpression: '',
            validationMessage: '',
            validationTargetField: '',
            selectedValidationRuleIndex: '',
            validationDesignerMessage: ''
        };
        return _this;
    }
    GridDesigner.prototype.componentDidMount = function () {
        this.loadFields();
    };
    GridDesigner.prototype.getWebUrl = function () {
        return this.props.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    };
    GridDesigner.prototype.getFieldsResponse = function (url) {
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
    GridDesigner.prototype.loadFields = function () {
        return __awaiter(this, void 0, void 0, function () {
            var fieldsEndpoint, url, response, data, sourceFields, fields, configuredFields, configuredIndex, configuredField, sourceIndex, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.props.listName) {
                            this.setState({ error: 'Select a SharePoint list before opening Grid Designer.' });
                            return [2 /*return*/];
                        }
                        this.setState({ loading: true, error: '' });
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 6, , 7]);
                        fieldsEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=";
                        url = fieldsEndpoint
                            + 'InternalName,Title,Description,TypeAsString,Required,ReadOnlyField,Hidden,FromBaseType,Choices,MaxLength,DisplayFormat';
                        return [4 /*yield*/, this.getFieldsResponse(url)];
                    case 2:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 4];
                        url = fieldsEndpoint + 'InternalName,Title,Description,TypeAsString,Required,ReadOnlyField,Hidden,Choices,DisplayFormat';
                        return [4 /*yield*/, this.getFieldsResponse(url)];
                    case 3:
                        response = _a.sent();
                        _a.label = 4;
                    case 4:
                        if (!response.ok) {
                            throw new Error('SharePoint fields could not be loaded. HTTP ' + String(response.status) + ' ' + String(response.statusText || ''));
                        }
                        return [4 /*yield*/, response.json()];
                    case 5:
                        data = _a.sent();
                        sourceFields = toArray(data && data.value).length > 0 ? toArray(data.value) : toArray(data && data.d && data.d.results);
                        fields = sourceFields.filter(function (field) {
                            var isAttachment = field.InternalName === 'Attachments';
                            return !field.Hidden && !SYSTEM_FIELDS[field.InternalName]
                                && (!field.FromBaseType || field.InternalName === 'Title' || isAttachment);
                        }).map(function (field) {
                            return {
                                internalName: String(field.InternalName || ''),
                                title: String(field.Title || field.InternalName || ''),
                                description: String(field.Description || ''),
                                type: String(field.TypeAsString || 'Text'),
                                required: field.Required === true,
                                readOnly: field.ReadOnlyField === true,
                                choices: toArray(field.Choices).map(function (choice) { return String(choice); }),
                                maxLength: field.MaxLength,
                                displayFormat: String(field.DisplayFormat) === '0' ? 'dateOnly' : 'dateTime'
                            };
                        });
                        configuredFields = copyFields(this.state.fields);
                        for (configuredIndex = 0; configuredIndex < configuredFields.length; configuredIndex += 1) {
                            configuredField = configuredFields[configuredIndex];
                            for (sourceIndex = 0; sourceIndex < fields.length; sourceIndex += 1) {
                                if (fields[sourceIndex].internalName.toLowerCase() === String(configuredField.fieldName || '').toLowerCase()) {
                                    configuredField.sharePointType = fields[sourceIndex].type;
                                    configuredField.sharePointRequired = fields[sourceIndex].required;
                                    configuredField.sharePointDescription = fields[sourceIndex].description;
                                    if (fields[sourceIndex].required) {
                                        configuredField.required = true;
                                    }
                                    if (!isCompatibleControlType(fields[sourceIndex].type, configuredField.type)) {
                                        configuredField.type = mapFieldType(fields[sourceIndex].type);
                                    }
                                    break;
                                }
                            }
                        }
                        this.setState({ sharePointFields: fields, fields: configuredFields, loading: false });
                        return [3 /*break*/, 7];
                    case 6:
                        error_1 = _a.sent();
                        this.setState({ loading: false, error: error_1 && error_1.message ? error_1.message : 'SharePoint fields could not be loaded.' });
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    GridDesigner.prototype.getAvailableFields = function () {
        var used = {};
        this.state.fields.forEach(function (field) { used[String(field.fieldName || '').toLowerCase()] = true; });
        return this.state.sharePointFields.filter(function (field) { return !used[field.internalName.toLowerCase()]; });
    };
    GridDesigner.prototype.addField = function (source) {
        var config = {};
        if (source.choices.length > 0) {
            config.choices = source.choices.slice();
        }
        if (source.maxLength) {
            config.maxLength = source.maxLength;
        }
        if (source.type === 'DateTime') {
            config.displayFormat = source.displayFormat || 'dateTime';
        }
        var field = {
            id: createId(),
            fieldName: source.internalName,
            sharePointType: source.type,
            sharePointRequired: source.required,
            sharePointDescription: source.description,
            type: mapFieldType(source.type),
            label: source.title,
            description: '',
            visible: true,
            required: source.required,
            readOnly: source.readOnly,
            config: config
        };
        var fields = copyFields(this.state.fields);
        fields.push(field);
        this.setState({ fields: fields, selectedFieldId: field.id });
    };
    GridDesigner.prototype.getSelectedField = function () {
        for (var i = 0; i < this.state.fields.length; i += 1) {
            if (this.state.fields[i].id === this.state.selectedFieldId) {
                return this.state.fields[i];
            }
        }
        return undefined;
    };
    GridDesigner.prototype.updateSelectedField = function (mutator) {
        var fields = copyFields(this.state.fields);
        for (var i = 0; i < fields.length; i += 1) {
            if (fields[i].id === this.state.selectedFieldId) {
                mutator(fields[i]);
                break;
            }
        }
        this.setState({ fields: fields });
    };
    GridDesigner.prototype.moveField = function (fieldId, direction) {
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
    GridDesigner.prototype.removeField = function (fieldId) {
        var fields = this.state.fields.filter(function (field) { return field.id !== fieldId; });
        this.setState({
            fields: fields,
            selectedFieldId: this.state.selectedFieldId === fieldId ? (fields.length > 0 ? fields[0].id : '') : this.state.selectedFieldId
        });
    };
    GridDesigner.prototype.updateConfig = function (name, value) {
        this.updateSelectedField(function (field) {
            field.config = field.config || {};
            field.config[name] = value;
        });
    };
    GridDesigner.prototype.changeSelectedControlType = function (controlType) {
        this.updateSelectedField(function (field) {
            field.type = controlType;
            var allowedValidationTypes = controlType === 'number' ? ['min', 'max']
                : (controlType === 'text' || controlType === 'multiline' || controlType === 'url')
                    ? ['minLength', 'maxLength', 'pattern'] : [];
            field.validation = (field.validation || []).filter(function (rule) {
                return allowedValidationTypes.indexOf(rule.type) >= 0;
            });
            field.config = field.config || {};
            if (controlType === 'number') {
                delete field.config.maxLength;
                delete field.config.choices;
                delete field.config.displayFormat;
            }
            else if (controlType === 'text' || controlType === 'multiline' || controlType === 'url') {
                delete field.config.min;
                delete field.config.max;
                delete field.config.decimals;
                delete field.config.choices;
                delete field.config.displayFormat;
            }
        });
    };
    GridDesigner.prototype.updateValidation = function (type, value, message) {
        this.updateSelectedField(function (field) {
            field.validation = (field.validation || []).filter(function (rule) { return rule.type !== type; });
            if (String(value || '').trim()) {
                field.validation.push({ type: type, value: value, message: message || 'The value is invalid.' });
            }
        });
    };
    GridDesigner.prototype.updateValidationMessages = function (types, message) {
        this.updateSelectedField(function (field) {
            var rules = field.validation || [];
            for (var ruleIndex = 0; ruleIndex < rules.length; ruleIndex += 1) {
                if (types.indexOf(rules[ruleIndex].type) >= 0) {
                    rules[ruleIndex].message = message || 'The value is invalid.';
                }
            }
            field.validation = rules;
        });
    };
    GridDesigner.prototype.getValidation = function (field, type) {
        var rules = field.validation || [];
        for (var i = 0; i < rules.length; i += 1) {
            if (rules[i].type === type) {
                return rules[i];
            }
        }
        return {};
    };
    GridDesigner.prototype.getPatternError = function (pattern) {
        if (!String(pattern || '').trim()) {
            return '';
        }
        try {
            new RegExp(pattern);
            return '';
        }
        catch (error) {
            return error && error.message ? error.message : 'The regular expression is invalid.';
        }
    };
    GridDesigner.prototype.getValidationFields = function () {
        return this.state.fields.map(function (field) {
            return { id: field.id, fieldName: field.fieldName, label: field.label };
        });
    };
    GridDesigner.prototype.validateAdvancedExpression = function (expression) {
        try {
            GridValidation_1.evaluateGridValidationExpression(expression, this.getValidationFields(), {});
            return '';
        }
        catch (error) {
            return error && error.message ? error.message : 'The validation expression is invalid.';
        }
    };
    GridDesigner.prototype.appendValidationExpression = function (fragment) {
        var current = String(this.state.validationExpression || '').trim();
        this.setState({ validationExpression: current ? current + ' ' + fragment : fragment, validationDesignerMessage: '' });
    };
    GridDesigner.prototype.loadValidationRule = function (indexValue) {
        var index = parseInt(indexValue, 10);
        var rule = !isNaN(index) ? this.state.advancedValidationRules[index] : undefined;
        this.setState({
            selectedValidationRuleIndex: indexValue,
            validationExpression: rule ? rule.expression : '',
            validationMessage: rule ? rule.message : '',
            validationTargetField: rule && rule.targetField ? rule.targetField : '',
            validationDesignerMessage: ''
        });
    };
    GridDesigner.prototype.saveValidationRule = function (updateExisting) {
        var expression = String(this.state.validationExpression || '').trim();
        var message = String(this.state.validationMessage || '').trim();
        if (!expression || !message) {
            this.setState({ validationDesignerMessage: 'Enter both an expression and validation message.' });
            return;
        }
        var expressionError = this.validateAdvancedExpression(expression);
        if (expressionError) {
            this.setState({ validationDesignerMessage: 'Expression error: ' + expressionError });
            return;
        }
        var rules = this.state.advancedValidationRules.slice(0);
        var selectedIndex = parseInt(this.state.selectedValidationRuleIndex, 10);
        var rule = {
            id: !isNaN(selectedIndex) && rules[selectedIndex] && rules[selectedIndex].id
                ? rules[selectedIndex].id : createId(),
            expression: expression,
            message: message,
            targetField: this.state.validationTargetField || undefined
        };
        if (updateExisting && !isNaN(selectedIndex) && rules[selectedIndex]) {
            rules[selectedIndex] = rule;
        }
        else {
            rules.push(rule);
            selectedIndex = rules.length - 1;
        }
        this.setState({
            advancedValidationRules: rules,
            advancedValidationEnabled: true,
            selectedValidationRuleIndex: String(selectedIndex),
            validationDesignerMessage: updateExisting ? 'Validation rule updated.' : 'Validation rule added.'
        });
    };
    GridDesigner.prototype.removeValidationRule = function () {
        var index = parseInt(this.state.selectedValidationRuleIndex, 10);
        if (isNaN(index) || !this.state.advancedValidationRules[index]) {
            return;
        }
        var rules = this.state.advancedValidationRules.slice(0);
        rules.splice(index, 1);
        this.setState({
            advancedValidationRules: rules,
            advancedValidationEnabled: rules.length > 0 && this.state.advancedValidationEnabled,
            selectedValidationRuleIndex: '',
            validationExpression: '',
            validationMessage: '',
            validationTargetField: '',
            validationDesignerMessage: 'Validation rule removed.'
        });
    };
    GridDesigner.prototype.save = function () {
        if (this.state.loading) {
            this.setState({ validationDesignerMessage: 'Wait for SharePoint field metadata to finish loading.' });
            return;
        }
        var fieldsToSave = copyFields(this.state.fields);
        for (var fieldIndex = 0; fieldIndex < fieldsToSave.length; fieldIndex += 1) {
            for (var sourceIndex = 0; sourceIndex < this.state.sharePointFields.length; sourceIndex += 1) {
                if (this.state.sharePointFields[sourceIndex].internalName.toLowerCase() === String(fieldsToSave[fieldIndex].fieldName || '').toLowerCase()
                    && this.state.sharePointFields[sourceIndex].required) {
                    fieldsToSave[fieldIndex].required = true;
                    fieldsToSave[fieldIndex].sharePointRequired = true;
                    break;
                }
            }
        }
        for (var fieldIndex = 0; fieldIndex < fieldsToSave.length; fieldIndex += 1) {
            var patternRule = this.getValidation(fieldsToSave[fieldIndex], 'pattern');
            var patternError = this.getPatternError(patternRule.value || '');
            if (patternError) {
                this.setState({ selectedFieldId: fieldsToSave[fieldIndex].id, validationDesignerMessage: 'Fix the invalid pattern before saving: ' + patternError });
                return;
            }
        }
        for (var ruleIndex = 0; ruleIndex < this.state.advancedValidationRules.length; ruleIndex += 1) {
            var expressionError = this.validateAdvancedExpression(this.state.advancedValidationRules[ruleIndex].expression);
            if (expressionError) {
                this.setState({ selectedValidationRuleIndex: String(ruleIndex), validationDesignerMessage: 'Fix validation rule ' + String(ruleIndex + 1) + ': ' + expressionError });
                return;
            }
        }
        var schema = {
            id: 'grid_' + String(this.props.listName || '').replace(/[^a-z0-9]/gi, '_').toLowerCase(),
            name: this.props.listName + ' Grid',
            mode: 'edit',
            listName: this.props.listName,
            advancedValidation: {
                enabled: this.state.advancedValidationEnabled && this.state.advancedValidationRules.length > 0,
                rules: this.state.advancedValidationRules
            },
            steps: [{ id: 'grid_columns', title: 'Grid columns', showTitle: false, visible: true, fields: fieldsToSave }]
        };
        this.props.onSave(JSON.stringify(schema));
    };
    GridDesigner.prototype.renderPalette = function () {
        var _this = this;
        var available = this.getAvailableFields();
        return (React.createElement("aside", { className: "gd-panel gd-palette" },
            React.createElement("h3", null, "SharePoint fields"),
            React.createElement("p", null, "Add fields as editable grid columns."),
            this.state.loading && React.createElement("div", { className: "gd-empty" }, "Loading fields..."),
            this.state.error && React.createElement("div", { className: "gd-error" }, this.state.error),
            !this.state.loading && !this.state.error && available.length === 0 && React.createElement("div", { className: "gd-empty" }, "All available fields are in the grid."),
            available.map(function (field) { return (React.createElement("button", { key: field.internalName, type: "button", className: "gd-palette-item", onClick: function () { return _this.addField(field); } },
                React.createElement("span", null, field.title),
                React.createElement("small", null, getTypeLabel(mapFieldType(field.type))))); })));
    };
    GridDesigner.prototype.renderCanvas = function () {
        var _this = this;
        return (React.createElement("main", { className: "gd-canvas" },
            React.createElement("div", { className: "gd-canvas-header" },
                React.createElement("div", null,
                    React.createElement("h3", null, "Grid columns"),
                    React.createElement("p", null, "Column order follows this list from top to bottom.")),
                React.createElement("span", null,
                    this.state.fields.length,
                    " columns")),
            this.state.fields.length === 0 && React.createElement("div", { className: "gd-empty gd-empty-canvas" }, "Add SharePoint fields from the left panel."),
            this.state.fields.map(function (field, index) { return (React.createElement("div", { key: field.id, className: 'gd-column ' + (field.id === _this.state.selectedFieldId ? 'gd-column-selected' : ''), onClick: function () { return _this.setState({ selectedFieldId: field.id }); } },
                React.createElement("div", { className: "gd-column-order" }, index + 1),
                React.createElement("div", { className: "gd-column-main" },
                    React.createElement("strong", null, field.label || field.fieldName),
                    React.createElement("span", null,
                        field.fieldName,
                        " \u00B7 ",
                        getTypeLabel(field.type)),
                    React.createElement("div", { className: "gd-chips" },
                        field.required === true && React.createElement("em", null, "Required"),
                        field.readOnly === true && React.createElement("em", null, "Read only"),
                        field.visible === false && React.createElement("em", null, "Hidden"))),
                React.createElement("div", { className: "gd-column-actions" },
                    React.createElement("button", { type: "button", title: "Move up", disabled: index === 0, onClick: function (ev) { ev.stopPropagation(); _this.moveField(field.id, -1); } }, "\u2191"),
                    React.createElement("button", { type: "button", title: "Move down", disabled: index === _this.state.fields.length - 1, onClick: function (ev) { ev.stopPropagation(); _this.moveField(field.id, 1); } }, "\u2193"),
                    React.createElement("button", { type: "button", title: "Remove column", onClick: function (ev) { ev.stopPropagation(); _this.removeField(field.id); } }, "\u00D7")))); })));
    };
    GridDesigner.prototype.renderFieldEditor = function () {
        var _this = this;
        var field = this.getSelectedField();
        if (!field) {
            return React.createElement("aside", { className: "gd-panel gd-properties" },
                React.createElement("div", { className: "gd-empty" }, "Select a grid column to configure it."));
        }
        var config = field.config || {};
        var compatibleControlTypes = getCompatibleControlTypes(field.sharePointType || 'Text');
        var minLengthRule = this.getValidation(field, 'minLength');
        var maxLengthRule = this.getValidation(field, 'maxLength');
        var minRule = this.getValidation(field, 'min');
        var maxRule = this.getValidation(field, 'max');
        var patternRule = this.getValidation(field, 'pattern');
        var patternError = this.getPatternError(patternRule.value || '');
        return (React.createElement("aside", { className: "gd-panel gd-properties" },
            React.createElement("h3", null, "Column settings"),
            React.createElement("label", null,
                "Column label",
                React.createElement("input", { type: "text", value: field.label || '', onChange: function (ev) { return _this.updateSelectedField(function (next) { next.label = ev.currentTarget.value; }); } })),
            field.sharePointDescription && React.createElement("label", null,
                "SharePoint description",
                React.createElement("div", { className: "gd-readonly-value" }, field.sharePointDescription)),
            React.createElement("label", null,
                field.sharePointDescription ? 'Designer description (SharePoint description takes precedence)' : 'Description',
                React.createElement("textarea", { value: field.description || '', onChange: function (ev) { return _this.updateSelectedField(function (next) { next.description = ev.currentTarget.value; }); } })),
            React.createElement("label", null,
                "SharePoint field type",
                React.createElement("div", { className: "gd-readonly-value" }, field.sharePointType || 'Unknown')),
            React.createElement("label", null,
                "Cell control",
                React.createElement("select", { value: field.type, onChange: function (ev) { return _this.changeSelectedControlType(ev.currentTarget.value); } }, compatibleControlTypes.map(function (controlType) {
                    return React.createElement("option", { key: controlType, value: controlType }, getTypeLabel(controlType));
                }))),
            React.createElement("div", { className: "gd-hint" }, "Only controls that can be safely saved to this SharePoint field are available."),
            React.createElement("label", null,
                "Column width",
                React.createElement("input", { type: "text", value: field.gridWidth || '', placeholder: "180px or 20%", onChange: function (ev) { return _this.updateSelectedField(function (next) { next.gridWidth = ev.currentTarget.value; }); } })),
            React.createElement("label", { className: "gd-check" },
                React.createElement("input", { type: "checkbox", checked: field.visible !== false, onChange: function (ev) { return _this.updateSelectedField(function (next) { next.visible = ev.currentTarget.checked; }); } }),
                " Visible"),
            React.createElement("label", { className: "gd-check" },
                React.createElement("input", { type: "checkbox", checked: field.required === true || field.sharePointRequired === true, disabled: field.sharePointRequired === true, onChange: function (ev) { return _this.updateSelectedField(function (next) { next.required = ev.currentTarget.checked; }); } }),
                " ",
                field.sharePointRequired === true ? 'Required by SharePoint' : 'Required'),
            React.createElement("label", { className: "gd-check" },
                React.createElement("input", { type: "checkbox", checked: field.readOnly === true, onChange: function (ev) { return _this.updateSelectedField(function (next) { next.readOnly = ev.currentTarget.checked; }); } }),
                " Read only"),
            field.required === true && React.createElement("label", null,
                "Required message",
                React.createElement("input", { type: "text", value: field.requiredMessage || '', onChange: function (ev) { return _this.updateSelectedField(function (next) { next.requiredMessage = ev.currentTarget.value; }); } })),
            React.createElement("label", null,
                "Placeholder",
                React.createElement("input", { type: "text", value: config.placeholder || '', onChange: function (ev) { return _this.updateConfig('placeholder', ev.currentTarget.value); } })),
            React.createElement("label", null,
                "Default value",
                React.createElement("input", { type: "text", value: field.defaultValue === undefined ? '' : String(field.defaultValue), onChange: function (ev) { return _this.updateSelectedField(function (next) { next.defaultValue = ev.currentTarget.value; }); } })),
            (field.type === 'text' || field.type === 'multiline') && React.createElement("label", null,
                "Maximum length",
                React.createElement("input", { type: "number", min: "0", value: config.maxLength || '', onChange: function (ev) { return _this.updateConfig('maxLength', parseInt(ev.currentTarget.value, 10) || undefined); } })),
            field.type === 'number' && React.createElement("div", { className: "gd-inline-fields" },
                React.createElement("label", null,
                    "Minimum",
                    React.createElement("input", { type: "number", value: config.min === undefined ? '' : config.min, onChange: function (ev) { return _this.updateConfig('min', ev.currentTarget.value === '' ? undefined : Number(ev.currentTarget.value)); } })),
                React.createElement("label", null,
                    "Maximum",
                    React.createElement("input", { type: "number", value: config.max === undefined ? '' : config.max, onChange: function (ev) { return _this.updateConfig('max', ev.currentTarget.value === '' ? undefined : Number(ev.currentTarget.value)); } })),
                React.createElement("label", null,
                    "Decimals",
                    React.createElement("input", { type: "number", min: "0", max: "10", value: config.decimals === undefined ? '' : config.decimals, onChange: function (ev) { return _this.updateConfig('decimals', ev.currentTarget.value === '' ? undefined : Number(ev.currentTarget.value)); } }))),
            (field.type === 'dropdown' || field.type === 'multiselect') && React.createElement("label", null,
                "Choices, one per line",
                React.createElement("textarea", { value: (config.choices || []).join('\n'), onChange: function (ev) { return _this.updateConfig('choices', ev.currentTarget.value.split(/\r?\n/).filter(function (choice) { return !!choice.trim(); })); } })),
            field.type === 'datetime' && React.createElement("label", null,
                "Date/time format",
                React.createElement("select", { value: config.displayFormat || 'dateTime', onChange: function (ev) { return _this.updateConfig('displayFormat', ev.currentTarget.value); } },
                    React.createElement("option", { value: "dateOnly" }, "Date only"),
                    React.createElement("option", { value: "dateTime" }, "Date and time"),
                    React.createElement("option", { value: "timeOnly" }, "Time only"))),
            React.createElement("div", { className: "gd-section-title" }, "Validation"),
            (field.type === 'text' || field.type === 'multiline' || field.type === 'url') && (React.createElement("div", null,
                React.createElement("div", { className: "gd-inline-fields gd-inline-fields-two" },
                    React.createElement("label", null,
                        "Minimum length",
                        React.createElement("input", { type: "number", min: "0", value: minLengthRule.value || '', onChange: function (ev) { return _this.updateValidation('minLength', ev.currentTarget.value, minLengthRule.message || 'The value is too short.'); } })),
                    React.createElement("label", null,
                        "Maximum length",
                        React.createElement("input", { type: "number", min: "0", value: maxLengthRule.value || '', onChange: function (ev) { return _this.updateValidation('maxLength', ev.currentTarget.value, maxLengthRule.message || 'The value is too long.'); } }))),
                React.createElement("label", null,
                    "Length validation message",
                    React.createElement("input", { type: "text", value: minLengthRule.message || maxLengthRule.message || '', onChange: function (ev) { return _this.updateValidationMessages(['minLength', 'maxLength'], ev.currentTarget.value); } })),
                React.createElement("label", null,
                    "Pattern (optional)",
                    React.createElement("input", { type: "text", value: patternRule.value || '', placeholder: "Regular expression", onChange: function (ev) { return _this.updateValidation('pattern', ev.currentTarget.value, patternRule.message || 'The value format is invalid.'); } })),
                patternError && React.createElement("div", { className: "gd-error gd-validation-error" }, patternError),
                React.createElement("label", null,
                    "Pattern message",
                    React.createElement("input", { type: "text", value: patternRule.message || '', onChange: function (ev) { return _this.updateValidation('pattern', patternRule.value || '', ev.currentTarget.value); } })))),
            field.type === 'number' && (React.createElement("div", null,
                React.createElement("div", { className: "gd-inline-fields gd-inline-fields-two" },
                    React.createElement("label", null,
                        "Minimum value",
                        React.createElement("input", { type: "number", value: minRule.value || '', onChange: function (ev) { return _this.updateValidation('min', ev.currentTarget.value, minRule.message || 'The value is below the minimum.'); } })),
                    React.createElement("label", null,
                        "Maximum value",
                        React.createElement("input", { type: "number", value: maxRule.value || '', onChange: function (ev) { return _this.updateValidation('max', ev.currentTarget.value, maxRule.message || 'The value exceeds the maximum.'); } }))),
                React.createElement("label", null,
                    "Range validation message",
                    React.createElement("input", { type: "text", value: minRule.message || maxRule.message || '', onChange: function (ev) { return _this.updateValidationMessages(['min', 'max'], ev.currentTarget.value); } })))),
            React.createElement("div", { className: "gd-section-title" }, "Advanced validation"),
            React.createElement("label", { className: "gd-check" },
                React.createElement("input", { type: "checkbox", checked: this.state.advancedValidationEnabled, onChange: function (ev) { return _this.setState({ advancedValidationEnabled: ev.currentTarget.checked }); } }),
                " Enable advanced validation rules"),
            React.createElement("div", { className: "gd-hint" }, "Uses the same expression grammar as Dynamic Form: ==, !=, >, >=, <, <=, &&, ||, !, and parentheses."),
            React.createElement("div", { className: "gd-validation-actions" },
                React.createElement("button", { type: "button", onClick: function () { return _this.appendValidationExpression('field("' + field.fieldName.replace(/"/g, '\\"') + '")'); } }, "Insert selected field"),
                React.createElement("button", { type: "button", onClick: function () { return _this.appendValidationExpression('hasValue(field("' + field.fieldName.replace(/"/g, '\\"') + '"))'); } }, "Has value"),
                React.createElement("button", { type: "button", onClick: function () { return _this.appendValidationExpression('isEmpty(field("' + field.fieldName.replace(/"/g, '\\"') + '"))'); } }, "Is empty"),
                React.createElement("button", { type: "button", onClick: function () { return _this.appendValidationExpression('between(field("' + field.fieldName.replace(/"/g, '\\"') + '"), 1, 10)'); } }, "Between"),
                React.createElement("button", { type: "button", onClick: function () { return _this.appendValidationExpression('today()'); } }, "Today"),
                React.createElement("button", { type: "button", onClick: function () { return _this.appendValidationExpression('daysFromToday(7)'); } }, "Days from today")),
            React.createElement("label", null,
                "Expression",
                React.createElement("textarea", { value: this.state.validationExpression, placeholder: 'field("' + field.fieldName + '") == "value"', onChange: function (ev) { return _this.setState({ validationExpression: ev.currentTarget.value, validationDesignerMessage: '' }); } })),
            React.createElement("label", null,
                "Message target",
                React.createElement("select", { value: this.state.validationTargetField, onChange: function (ev) { return _this.setState({ validationTargetField: ev.currentTarget.value }); } },
                    React.createElement("option", { value: "" }, "Grid-level message"),
                    this.state.fields.map(function (targetField) { return React.createElement("option", { key: targetField.id, value: targetField.fieldName }, targetField.label || targetField.fieldName); }))),
            React.createElement("label", null,
                "Validation message",
                React.createElement("input", { type: "text", value: this.state.validationMessage, onChange: function (ev) { return _this.setState({ validationMessage: ev.currentTarget.value, validationDesignerMessage: '' }); } })),
            React.createElement("label", null,
                "Existing rules",
                React.createElement("select", { value: this.state.selectedValidationRuleIndex, onChange: function (ev) { return _this.loadValidationRule(ev.currentTarget.value); } },
                    React.createElement("option", { value: "" }, "Select a rule"),
                    this.state.advancedValidationRules.map(function (rule, ruleIndex) { return React.createElement("option", { key: String(ruleIndex), value: String(ruleIndex) }, String(ruleIndex + 1) + '. ' + rule.message); }))),
            React.createElement("div", { className: "gd-validation-actions" },
                React.createElement("button", { type: "button", onClick: function () { return _this.saveValidationRule(false); } }, "Add rule"),
                React.createElement("button", { type: "button", disabled: this.state.selectedValidationRuleIndex === '', onClick: function () { return _this.saveValidationRule(true); } }, "Update selected"),
                React.createElement("button", { type: "button", disabled: this.state.selectedValidationRuleIndex === '', onClick: function () { return _this.removeValidationRule(); } }, "Remove selected"),
                React.createElement("button", { type: "button", onClick: function () { return _this.setState({ validationExpression: '', validationMessage: '', validationTargetField: '', selectedValidationRuleIndex: '', validationDesignerMessage: '' }); } }, "Clear editor")),
            this.state.validationDesignerMessage && React.createElement("div", { className: "gd-validation-message" }, this.state.validationDesignerMessage)));
    };
    GridDesigner.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: "gd-overlay" },
            React.createElement("header", { className: "gd-toolbar" },
                React.createElement("div", null,
                    React.createElement("h2", null, "Grid Designer"),
                    React.createElement("span", null, this.props.listName || 'No list selected')),
                React.createElement("div", { className: "gd-toolbar-actions" },
                    React.createElement("button", { type: "button", onClick: this.props.onCancel }, "Cancel"),
                    React.createElement("button", { type: "button", className: "gd-primary", onClick: function () { return _this.save(); }, disabled: !this.props.listName || this.state.loading }, "Save grid"))),
            React.createElement("div", { className: "gd-workspace" },
                this.renderPalette(),
                this.renderCanvas(),
                this.renderFieldEditor())));
    };
    return GridDesigner;
}(React.Component));
exports.GridDesigner = GridDesigner;

//# sourceMappingURL=GridDesigner.js.map
