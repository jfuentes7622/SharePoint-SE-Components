"use strict";
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
function validateLibraryFilterJson(sources) {
    sources.forEach(function (source) {
        if (!String(source || '').trim()) {
            return;
        }
        var parsed = JSON.parse(source);
        if (!Array.isArray(parsed) || parsed.some(function (condition) {
            return !condition || typeof condition.field !== 'string' || !condition.field.trim();
        })) {
            throw new Error('Library filters must be an array of conditions with field names.');
        }
    });
}
exports.validateLibraryFilterJson = validateLibraryFilterJson;
function resolveNowExpression(expression, time) {
    var normalized = String(expression || '').trim().toLowerCase();
    if (normalized === 'now') {
        return new Date(time).toISOString();
    }
    var match = /^now\(\s*([+-]?\d+)h\s*\)$/.exec(normalized);
    var value = match ? time + Number(match[1]) * 3600000 : NaN;
    if (!isFinite(value) || Math.abs(value) > 8640000000000000) {
        throw new Error('Invalid elapsed-hour expression: ' + expression);
    }
    return new Date(value).toISOString();
}
exports.resolveNowExpression = resolveNowExpression;
function documentUrl(webUrl, path) {
    var value = String(path || '').trim();
    if (!value || /^\/\//.test(value) || /^(?!https?:)[a-z][a-z0-9+.-]*:/i.test(value) || /[\\\r\n]/.test(value)) {
        return '';
    }
    var origin = /^https?:\/\/[^/]+/i.exec(webUrl);
    if (!origin) {
        return '';
    }
    var absolute = /^https?:\/\//i.test(value) ? value
        : value.charAt(0) === '/' ? origin[0] + value : webUrl.replace(/\/$/, '') + '/' + value;
    // FileRef can already contain escapes. Preserve those, but escape literal %, # and spaces.
    return encodeURI(absolute.replace(/%(?![0-9a-f]{2})/ig, '%25'))
        .replace(/%25([0-9a-f]{2})/ig, '%$1').replace(/#/g, '%23');
}
exports.documentUrl = documentUrl;
function cleanPath(path) {
    var value = String(path || '').replace(/\/+$/, '');
    if (!value || value.charAt(0) !== '/' || /[\\\r\n]/.test(value)
        || value.split('/').some(function (segment) {
            return segment === '.' || segment === '..' || /^(%2e){1,2}$/i.test(segment);
        })) {
        throw new Error('Invalid library path: ' + path);
    }
    return value;
}
function withinRoot(root, path) {
    return path.toLowerCase() === root.toLowerCase()
        || path.toLowerCase().indexOf(root.toLowerCase() + '/') === 0;
}
function folderDocuments(current, documents, recursive) {
    if (recursive === void 0) { recursive = false; }
    var path = cleanPath(current);
    return documents.filter(function (row) {
        var parent = cleanPath(row.FileDirRef);
        return recursive ? withinRoot(path, parent) : parent.toLowerCase() === path.toLowerCase();
    });
}
exports.folderDocuments = folderDocuments;
function folderDocumentCount(current, documents) {
    return folderDocuments(current, documents, true).length;
}
exports.folderDocumentCount = folderDocumentCount;
function ancestorPaths(rootPath, rows) {
    var root = cleanPath(rootPath);
    var paths = {};
    rows.forEach(function (row) {
        var parent = cleanPath(row.FileDirRef);
        if (!withinRoot(root, parent)) {
            throw new Error('Document is outside the library root: ' + parent);
        }
        while (parent.toLowerCase() !== root.toLowerCase()) {
            paths[parent.toLowerCase()] = parent;
            parent = parent.substring(0, parent.lastIndexOf('/'));
        }
    });
    return Object.keys(paths).map(function (key) { return paths[key]; });
}
exports.ancestorPaths = ancestorPaths;
function folderLevel(root, current, documents, folders, prune) {
    if (prune === void 0) { prune = true; }
    var required = ancestorPaths(root, documents).map(function (path) { return path.toLowerCase(); });
    var children = folders.filter(function (folder) {
        return (!prune || required.indexOf(folder.path.toLowerCase()) >= 0)
            && folder.path.substring(0, folder.path.lastIndexOf('/')).toLowerCase() === current.toLowerCase();
    }).sort(function (left, right) { return left.name.localeCompare(right.name); });
    return children.map(function (folder) { return { libraryFolder: folder }; })
        .concat(folderDocuments(current, documents));
}
exports.folderLevel = folderLevel;
function nearestFolder(root, current, paths) {
    var keys = paths.map(function (path) { return path.toLowerCase(); });
    var candidate = current || root;
    while (withinRoot(root, candidate) && candidate.toLowerCase() !== root.toLowerCase()) {
        if (keys.indexOf(candidate.toLowerCase()) >= 0) {
            return candidate;
        }
        candidate = candidate.substring(0, candidate.lastIndexOf('/'));
    }
    return root;
}
exports.nearestFolder = nearestFolder;
function recursiveView(viewXml, fields, conditions, time, batchSize, dateFields, recursive) {
    if (recursive === void 0) { recursive = true; }
    var doc = new DOMParser().parseFromString(viewXml || '<View/>', 'text/xml');
    if (doc.getElementsByTagName('parsererror').length) {
        throw new Error('Invalid source view XML.');
    }
    var view = doc.documentElement;
    if (recursive) {
        view.setAttribute('Scope', 'RecursiveAll');
    }
    ['GroupBy', 'Aggregations', 'Folder'].forEach(function (tag) {
        var nodes = view.getElementsByTagName(tag);
        while (nodes.length) {
            nodes[0].parentNode.removeChild(nodes[0]);
        }
    });
    var query = view.getElementsByTagName('Query')[0];
    if (!query) {
        query = doc.createElement('Query');
        view.appendChild(query);
    }
    var where = query.getElementsByTagName('Where')[0];
    if (!where) {
        where = doc.createElement('Where');
        query.appendChild(where);
    }
    function predicate(tag, field, type, value) {
        var node = doc.createElement(tag);
        var ref = doc.createElement('FieldRef');
        ref.setAttribute('Name', field);
        var val = doc.createElement('Value');
        val.setAttribute('Type', type);
        if (type === 'DateTime') {
            val.setAttribute('IncludeTimeValue', value.indexOf('T') >= 0 ? 'TRUE' : 'FALSE');
            if (value.indexOf('T') >= 0) {
                val.setAttribute('StorageTZ', 'TRUE');
            }
        }
        val.appendChild(doc.createTextNode(value));
        node.appendChild(ref);
        node.appendChild(val);
        return node;
    }
    function appendPredicate(node) {
        var existing = where.firstElementChild;
        if (existing) {
            var and = doc.createElement('And');
            where.removeChild(existing);
            and.appendChild(existing);
            and.appendChild(node);
            where.appendChild(and);
        }
        else {
            where.appendChild(node);
        }
    }
    appendPredicate(predicate('Eq', 'FSObjType', 'Integer', '0'));
    // Push down only safe conjuncts. Mixed OR expressions retain their complete client evaluation.
    if (!conditions.some(function (condition, index) {
        return index > 0 && String(condition.logical || '').toLowerCase() === 'or';
    })) {
        var operators = { eq: 'Eq', ge: 'Geq', le: 'Leq', gt: 'Gt', lt: 'Lt', ne: 'Neq' };
        conditions.forEach(function (condition) {
            if (dateFields && dateFields.indexOf(condition.field) < 0) {
                return;
            }
            var expression = String(condition.value || '').trim().toLowerCase();
            var value = '';
            if (String(condition.valueType).toLowerCase() === 'expression') {
                if (/^now(?:\(|$)/.test(expression)) {
                    value = resolveNowExpression(expression, time);
                }
                var days = /^date\(\s*([+-]?\d+)\s*\)$/.exec(expression);
                if (expression === 'today' || days) {
                    var local = new Date(time);
                    if (days) {
                        local.setDate(local.getDate() + Number(days[1]));
                    }
                    value = String(local.getFullYear()) + '-' + ('0' + String(local.getMonth() + 1)).slice(-2) + '-' + ('0' + String(local.getDate())).slice(-2);
                }
            }
            else if (typeof condition.value === 'string' && /^\d{4}-\d\d-\d\d(?:T.*)?$/.test(condition.value)) {
                value = condition.value;
            }
            var operator = operators[String(condition.operator || 'eq').toLowerCase()];
            if (value && operator) {
                appendPredicate(predicate(operator, condition.field, 'DateTime', value));
            }
        });
    }
    var viewFields = view.getElementsByTagName('ViewFields')[0];
    if (!viewFields) {
        viewFields = doc.createElement('ViewFields');
        view.appendChild(viewFields);
    }
    var existingNames = [];
    var refs = viewFields.getElementsByTagName('FieldRef');
    for (var index = 0; index < refs.length; index += 1) {
        existingNames.push(String(refs[index].getAttribute('Name') || '').toLowerCase());
    }
    fields.forEach(function (field) {
        if (existingNames.indexOf(field.toLowerCase()) >= 0) {
            return;
        }
        existingNames.push(field.toLowerCase());
        var ref = doc.createElement('FieldRef');
        ref.setAttribute('Name', field);
        viewFields.appendChild(ref);
    });
    var limit = view.getElementsByTagName('RowLimit')[0];
    if (!limit) {
        limit = doc.createElement('RowLimit');
        view.appendChild(limit);
    }
    limit.setAttribute('Paged', 'TRUE');
    limit.textContent = String(Math.max(1, batchSize || 200));
    return new XMLSerializer().serializeToString(doc);
}
exports.recursiveView = recursiveView;
function json(response, operation) {
    return __awaiter(this, void 0, void 0, function () {
        var body, _a, requestId, detail, guidance;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!!response.ok) return [3 /*break*/, 4];
                    if (!(typeof response.text === 'function')) return [3 /*break*/, 2];
                    return [4 /*yield*/, response.text()];
                case 1:
                    _a = _b.sent();
                    return [3 /*break*/, 3];
                case 2:
                    _a = '';
                    _b.label = 3;
                case 3:
                    body = _a;
                    requestId = response.headers && response.headers.get
                        ? response.headers.get('SPRequestGuid') || response.headers.get('request-id') : '';
                    detail = String(body || '').substring(0, 2000);
                    guidance = /SPQueryThrottledException|list view threshold|exceeds.*threshold/i.test(detail)
                        ? ' Use folder-scoped loading and a selective indexed view/filter. A folder can also exceed the list view threshold.' : '';
                    throw new Error(operation + ': HTTP ' + String(response.status) + ' ' + String(response.statusText || '')
                        + (detail ? ' - ' + detail : '') + (requestId ? ' [SPRequestGuid: ' + requestId + ']' : '') + guidance);
                case 4: return [2 /*return*/, response.json()];
            }
        });
    });
}
function values(data) {
    var rows = data && (data.value || (data.d && data.d.results));
    if (!Array.isArray(rows)) {
        throw new Error('Invalid SharePoint collection response.');
    }
    return rows;
}
function detectDocumentLibrary(webUrl, listName, get) {
    return __awaiter(this, void 0, void 0, function () {
        var data, _a, source;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _a = json;
                    return [4 /*yield*/, get(webUrl + "/_api/web/lists/getByTitle('"
                            + encodeURIComponent(listName.replace(/'/g, "''")) + "')?$select=BaseType")];
                case 1: return [4 /*yield*/, _a.apply(void 0, [_b.sent(), 'List type discovery'])];
                case 2:
                    data = _b.sent();
                    source = data && (data.d || data);
                    if (!source || !/^[0-5]$/.test(String(source.BaseType))) {
                        throw new Error('SharePoint did not return a valid list BaseType.');
                    }
                    return [2 /*return*/, Number(source.BaseType) === 1];
            }
        });
    });
}
exports.detectDocumentLibrary = detectDocumentLibrary;
function loadLibraryFolders(webUrl, listName, root, documents, get, current, batchSize) {
    return __awaiter(this, void 0, void 0, function () {
        var next, seenPages, byPath, data, _a, folders, required;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    next = webUrl + "/_api/web/lists/getByTitle('" + encodeURIComponent(listName.replace(/'/g, "''"))
                        + "')/items?$select=FileRef,FileDirRef,FileLeafRef&$filter=FSObjType%20eq%201&$top=" + String(Math.max(1, batchSize || 200));
                    seenPages = {};
                    byPath = {};
                    _b.label = 1;
                case 1:
                    if (!next) return [3 /*break*/, 4];
                    if (!current()) {
                        throw new Error('Superseded library folder request.');
                    }
                    if (seenPages[next]) {
                        throw new Error('Library folder paging did not advance.');
                    }
                    seenPages[next] = true;
                    _a = json;
                    return [4 /*yield*/, get(next)];
                case 2: return [4 /*yield*/, _a.apply(void 0, [_b.sent(), 'Library folders'])];
                case 3:
                    data = _b.sent();
                    values(data).forEach(function (item) {
                        var path = cleanPath(item.FileRef);
                        if (!withinRoot(root, path) || path.toLowerCase() === root.toLowerCase() || !item.FileLeafRef
                            || path.substring(0, path.lastIndexOf('/')).toLowerCase() !== cleanPath(item.FileDirRef).toLowerCase()) {
                            throw new Error('Invalid library folder metadata: ' + path);
                        }
                        byPath[path.toLowerCase()] = { path: path, name: String(item.FileLeafRef) };
                    });
                    next = data['@odata.nextLink'] || data['odata.nextLink'] || (data.d && data.d.__next) || '';
                    return [3 /*break*/, 1];
                case 4:
                    folders = Object.keys(byPath).map(function (key) { return byPath[key]; });
                    required = ancestorPaths(root, documents.concat(folders.map(function (folder) {
                        return { FileDirRef: folder.path.substring(0, folder.path.lastIndexOf('/')) };
                    })));
                    required.forEach(function (path) {
                        if (!byPath[path.toLowerCase()]) {
                            throw new Error('Library ancestor folder is unavailable: ' + path);
                        }
                    });
                    return [2 /*return*/, folders];
            }
        });
    });
}
exports.loadLibraryFolders = loadLibraryFolders;
function defaultLibraryViewId(webUrl, listName, get) {
    return __awaiter(this, void 0, void 0, function () {
        var data, _a, source;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _a = json;
                    return [4 /*yield*/, get(webUrl + "/_api/web/lists/getByTitle('"
                            + encodeURIComponent(listName.replace(/'/g, "''")) + "')/DefaultView?$select=Id")];
                case 1: return [4 /*yield*/, _a.apply(void 0, [_b.sent(), 'Library default view'])];
                case 2:
                    data = _b.sent();
                    source = data.d || data;
                    if (!source.Id) {
                        throw new Error('SharePoint did not return the default library view.');
                    }
                    return [2 /*return*/, String(source.Id)];
            }
        });
    });
}
exports.defaultLibraryViewId = defaultLibraryViewId;
function loadLibraryDocuments(options) {
    return __awaiter(this, void 0, void 0, function () {
        function internalName(name) {
            var match = metadata.filter(function (field) {
                return String(field.InternalName).toLowerCase() === name.toLowerCase() || String(field.Title).toLowerCase() === name.toLowerCase();
            })[0];
            return match ? String(match.InternalName) : name;
        }
        function hydrateFlatRows() {
            return __awaiter(this, void 0, void 0, function () {
                var row, id, itemData, _a, item, error_1;
                return __generator(this, function (_b) {
                    switch (_b.label) {
                        case 0:
                            if (!(cursor < rows.length && !hydrationError)) return [3 /*break*/, 6];
                            row = rows[cursor++];
                            id = Number(row.ID || row.Id);
                            _b.label = 1;
                        case 1:
                            _b.trys.push([1, 4, , 5]);
                            if (!options.current()) {
                                throw new Error('Superseded library request.');
                            }
                            if (!isFinite(id) || id <= 0 || Math.floor(id) !== id) {
                                throw new Error('Library returned an invalid document ID.');
                            }
                            _a = json;
                            return [4 /*yield*/, options.get(listUrl + '/items(' + String(id) + ')?$select=' + select.join(','))];
                        case 2: return [4 /*yield*/, _a.apply(void 0, [_b.sent(), 'Library document metadata'])];
                        case 3:
                            itemData = _b.sent();
                            item = itemData.d || itemData;
                            if (Number(item.ID || item.Id) !== id) {
                                throw new Error('Document metadata is no longer accessible: ' + String(id));
                            }
                            hydrated[String(id)] = item;
                            return [3 /*break*/, 5];
                        case 4:
                            error_1 = _b.sent();
                            hydrationError = error_1;
                            return [3 /*break*/, 5];
                        case 5: return [3 /*break*/, 0];
                        case 6: return [2 /*return*/];
                    }
                });
            });
        }
        var listUrl, sourceData, _a, source, root, folder, metadata, _b, _c, names, dateFields, queryConditions, viewXml, sourceScope, flat, rows, fields, paging, seen, parameters, data, _d, renderData, listData, renderError, batch, dateNames, select, hydrated, filesUrl, filePages, fileData, _e, cursor, hydrationError, offset, ids, next, itemPages, itemData, _f;
        return __generator(this, function (_g) {
            switch (_g.label) {
                case 0:
                    listUrl = options.webUrl + "/_api/web/lists/getByTitle('" + encodeURIComponent(options.listName.replace(/'/g, "''")) + "')";
                    _a = json;
                    return [4 /*yield*/, options.get(listUrl + '?$select=BaseType,RootFolder/ServerRelativeUrl&$expand=RootFolder')];
                case 1: return [4 /*yield*/, _a.apply(void 0, [_g.sent(), 'Library discovery'])];
                case 2:
                    sourceData = _g.sent();
                    source = sourceData.d || sourceData;
                    if (Number(source.BaseType) !== 1) {
                        throw new Error('Matching document folders requires a document library.');
                    }
                    root = cleanPath(source.RootFolder && source.RootFolder.ServerRelativeUrl);
                    folder = options.folder === undefined ? '' : cleanPath(options.folder || root);
                    if (folder && !withinRoot(root, folder)) {
                        throw new Error('Folder is outside the library root: ' + folder);
                    }
                    _b = values;
                    _c = json;
                    return [4 /*yield*/, options.get(listUrl
                            + '/fields?$select=InternalName,Title,TypeAsString,DisplayFormat,Choices,LookupList,LookupField,AllowMultipleValues')];
                case 3: return [4 /*yield*/, _c.apply(void 0, [_g.sent(), 'Library field metadata'])];
                case 4:
                    metadata = _b.apply(void 0, [_g.sent()]);
                    names = options.fields.concat(['ID', 'FSObjType', 'FileRef', 'FileDirRef', 'FileLeafRef', 'Modified', 'Created']);
                    options.conditions.forEach(function (condition) {
                        if (String(condition.valueType).toLowerCase() === 'expression') {
                            var expression = String(condition.value || '').trim().toLowerCase();
                            if (/^now(?:\(|$)/.test(expression)) {
                                resolveNowExpression(expression, options.time);
                            }
                            else if (['today', 'me', 'me.email', 'me.login', 'me.id'].indexOf(expression) < 0 && !/^date\(\s*[+-]?\d+\s*\)$/.test(expression)) {
                                throw new Error('Unsupported library filter expression: ' + expression);
                            }
                        }
                        if (names.indexOf(condition.field) < 0) {
                            names.push(condition.field);
                        }
                    });
                    names = names.map(internalName);
                    names = names.filter(function (name, index) { return !!name && names.indexOf(name) === index; });
                    dateFields = metadata.filter(function (field) { return field.TypeAsString === 'DateTime'; })
                        .map(function (field) { return field.InternalName; });
                    queryConditions = options.conditions.map(function (condition) {
                        return Object.assign({}, condition, { field: internalName(condition.field) });
                    });
                    viewXml = recursiveView(options.viewXml, names, queryConditions, options.time, options.batchSize, dateFields, !folder);
                    sourceScope = new DOMParser().parseFromString(options.viewXml || '<View/>', 'text/xml').documentElement.getAttribute('Scope') || '';
                    flat = /^(Recursive|RecursiveAll)$/i.test(sourceScope);
                    rows = [];
                    fields = [];
                    paging = '';
                    seen = {};
                    _g.label = 5;
                case 5:
                    if (!options.current()) {
                        throw new Error('Superseded library request.');
                    }
                    parameters = { RenderOptions: 7, ViewXml: viewXml, DatesInUtc: true };
                    if (folder) {
                        parameters.FolderServerRelativeUrl = folder;
                    }
                    if (paging) {
                        parameters.Paging = paging.replace(/^\?/, '');
                    }
                    _d = json;
                    return [4 /*yield*/, options.post(listUrl + '/RenderListDataAsStream', { parameters: parameters })];
                case 6: return [4 /*yield*/, _d.apply(void 0, [_g.sent(), 'Library documents'])];
                case 7:
                    data = _g.sent();
                    renderData = data && data.d && data.d.RenderListDataAsStream ? data.d.RenderListDataAsStream : data;
                    if (typeof renderData === 'string') {
                        renderData = JSON.parse(renderData);
                    }
                    listData = renderData.ListData || renderData;
                    if (typeof listData === 'string') {
                        listData = JSON.parse(listData);
                    }
                    renderError = renderData.ErrorInfo || renderData.ErrorMessage || listData.ErrorInfo || listData.ErrorMessage;
                    if (renderError) {
                        throw new Error('SharePoint could not render the library view: ' + JSON.stringify(renderError));
                    }
                    batch = options.extract(data);
                    rows = rows.concat(batch.rows);
                    if (!fields.length) {
                        fields = batch.fields;
                    }
                    paging = batch.nextHref || '';
                    if (paging && seen[paging]) {
                        throw new Error('Library paging did not advance.');
                    }
                    seen[paging] = true;
                    _g.label = 8;
                case 8:
                    if (paging) return [3 /*break*/, 5];
                    _g.label = 9;
                case 9:
                    dateNames = metadata.filter(function (field) {
                        return field.TypeAsString === 'DateTime' && names.indexOf(field.InternalName) >= 0;
                    }).map(function (field) { return field.InternalName; });
                    select = ['ID', 'FileRef', 'FileDirRef', 'FileLeafRef', 'FSObjType'].concat(dateNames);
                    hydrated = {};
                    if (!(folder && !flat && rows.length)) return [3 /*break*/, 14];
                    filesUrl = options.webUrl + "/_api/web/GetFolderByServerRelativeUrl('"
                        + encodeURIComponent(folder.replace(/'/g, "''")) + "')/Files?$expand=ListItemAllFields&$select="
                        + select.map(function (name) { return 'ListItemAllFields/' + name; }).join(',')
                        + '&$top=' + String(Math.max(1, options.batchSize || 200));
                    filePages = {};
                    _g.label = 10;
                case 10:
                    if (!filesUrl) return [3 /*break*/, 13];
                    if (!options.current()) {
                        throw new Error('Superseded library request.');
                    }
                    if (filePages[filesUrl]) {
                        throw new Error('Folder document metadata paging did not advance.');
                    }
                    filePages[filesUrl] = true;
                    _e = json;
                    return [4 /*yield*/, options.get(filesUrl)];
                case 11: return [4 /*yield*/, _e.apply(void 0, [_g.sent(), 'Folder document metadata'])];
                case 12:
                    fileData = _g.sent();
                    values(fileData).forEach(function (file) {
                        var item = file.ListItemAllFields;
                        var id = Number(item && item.ID);
                        if (!item || !isFinite(id) || id <= 0 || Math.floor(id) !== id || cleanPath(item.FileDirRef).toLowerCase() !== folder.toLowerCase()) {
                            throw new Error('Invalid folder document metadata.');
                        }
                        hydrated[String(item.ID)] = item;
                    });
                    filesUrl = fileData['@odata.nextLink'] || fileData['odata.nextLink'] || (fileData.d && fileData.d.__next) || '';
                    return [3 /*break*/, 10];
                case 13:
                    rows.forEach(function (row) {
                        if (!hydrated[String(row.ID || row.Id)]) {
                            throw new Error('Document metadata is no longer accessible: ' + String(row.ID || row.Id));
                        }
                    });
                    _g.label = 14;
                case 14:
                    cursor = 0;
                    if (!flat) return [3 /*break*/, 16];
                    // Indexed item reads avoid both library-wide OR filters and scanning unrelated folder files.
                    return [4 /*yield*/, Promise.all([hydrateFlatRows(), hydrateFlatRows(), hydrateFlatRows(), hydrateFlatRows()])];
                case 15:
                    // Indexed item reads avoid both library-wide OR filters and scanning unrelated folder files.
                    _g.sent();
                    if (hydrationError) {
                        throw hydrationError;
                    }
                    _g.label = 16;
                case 16:
                    offset = 0;
                    _g.label = 17;
                case 17:
                    if (!(!folder && !flat && offset < rows.length)) return [3 /*break*/, 23];
                    if (!options.current()) {
                        throw new Error('Superseded library request.');
                    }
                    ids = rows.slice(offset, offset + 50).map(function (row) { return Number(row.ID || row.Id); });
                    if (ids.some(function (id) { return !isFinite(id) || id <= 0 || Math.floor(id) !== id; })) {
                        throw new Error('Library returned an invalid document ID.');
                    }
                    next = listUrl + '/items?$top=50&$select=' + select.join(',') + '&$filter='
                        + encodeURIComponent(ids.map(function (id) { return 'ID eq ' + String(id); }).join(' or '));
                    itemPages = {};
                    _g.label = 18;
                case 18:
                    if (!next) return [3 /*break*/, 21];
                    if (!options.current()) {
                        throw new Error('Superseded library request.');
                    }
                    if (itemPages[next]) {
                        throw new Error('Library metadata paging did not advance.');
                    }
                    itemPages[next] = true;
                    _f = json;
                    return [4 /*yield*/, options.get(next)];
                case 19: return [4 /*yield*/, _f.apply(void 0, [_g.sent(), 'Library document metadata'])];
                case 20:
                    itemData = _g.sent();
                    values(itemData).forEach(function (item) { hydrated[String(item.ID || item.Id)] = item; });
                    next = itemData['@odata.nextLink'] || itemData['odata.nextLink'] || (itemData.d && itemData.d.__next) || '';
                    return [3 /*break*/, 18];
                case 21:
                    ids.forEach(function (id) {
                        if (!hydrated[String(id)]) {
                            throw new Error('Document metadata is no longer accessible: ' + String(id));
                        }
                    });
                    _g.label = 22;
                case 22:
                    offset += 50;
                    return [3 /*break*/, 17];
                case 23:
                    rows = rows.map(function (row) { return Object.assign({}, row, hydrated[String(row.ID || row.Id)]); })
                        .filter(function (row) { return Number(row.FSObjType || 0) === 0; });
                    return [2 /*return*/, { root: root, folder: folder || root, flat: flat, rows: rows, fields: fields, metadata: metadata }];
            }
        });
    });
}
exports.loadLibraryDocuments = loadLibraryDocuments;
function loadChildFolders(webUrl, root, folder, get, current, batchSize) {
    return __awaiter(this, void 0, void 0, function () {
        var path, next, pages, folders, data, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    path = cleanPath(folder);
                    if (!withinRoot(cleanPath(root), path)) {
                        throw new Error('Folder is outside the library root: ' + path);
                    }
                    next = webUrl + "/_api/web/GetFolderByServerRelativeUrl('" + encodeURIComponent(path.replace(/'/g, "''"))
                        + "')/Folders?$select=Name,ServerRelativeUrl&$top=" + String(Math.max(1, batchSize || 200));
                    pages = {};
                    folders = [];
                    _b.label = 1;
                case 1:
                    if (!next) return [3 /*break*/, 4];
                    if (!current()) {
                        throw new Error('Superseded library request.');
                    }
                    if (pages[next]) {
                        throw new Error('Child folder paging did not advance.');
                    }
                    pages[next] = true;
                    _a = json;
                    return [4 /*yield*/, get(next)];
                case 2: return [4 /*yield*/, _a.apply(void 0, [_b.sent(), 'Child folders'])];
                case 3:
                    data = _b.sent();
                    values(data).forEach(function (item) {
                        var child = cleanPath(item.ServerRelativeUrl);
                        if (child.substring(0, child.lastIndexOf('/')).toLowerCase() !== path.toLowerCase() || !item.Name) {
                            throw new Error('Invalid child folder metadata: ' + child);
                        }
                        if (path.toLowerCase() !== root.toLowerCase() || String(item.Name).toLowerCase() !== 'forms') {
                            folders.push({ path: child, name: String(item.Name) });
                        }
                    });
                    next = data['@odata.nextLink'] || data['odata.nextLink'] || (data.d && data.d.__next) || '';
                    return [3 /*break*/, 1];
                case 4: return [2 /*return*/, folders];
            }
        });
    });
}
exports.loadChildFolders = loadChildFolders;
function loadAncestorFolders(webUrl, root, rows, get) {
    return __awaiter(this, void 0, void 0, function () {
        function worker() {
            return __awaiter(this, void 0, void 0, function () {
                var path, endpoint, data, _a, source;
                return __generator(this, function (_b) {
                    switch (_b.label) {
                        case 0:
                            if (!(cursor < paths.length)) return [3 /*break*/, 3];
                            path = paths[cursor++];
                            endpoint = webUrl + "/_api/web/GetFolderByServerRelativeUrl('"
                                + encodeURIComponent(path.replace(/'/g, "''")) + "')?$select=Name,ServerRelativeUrl";
                            _a = json;
                            return [4 /*yield*/, get(endpoint)];
                        case 1: return [4 /*yield*/, _a.apply(void 0, [_b.sent(), 'Folder ' + path])];
                        case 2:
                            data = _b.sent();
                            source = data.d || data;
                            if (cleanPath(source.ServerRelativeUrl).toLowerCase() !== path.toLowerCase() || !source.Name) {
                                throw new Error('Invalid folder metadata: ' + path);
                            }
                            folders.push({ path: path, name: String(source.Name) });
                            return [3 /*break*/, 0];
                        case 3: return [2 /*return*/];
                    }
                });
            });
        }
        function boundedWorker() {
            return worker().catch(function (error) {
                failure = error;
                cursor = paths.length;
            });
        }
        var paths, folders, cursor, failure;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    paths = ancestorPaths(root, rows);
                    folders = [];
                    cursor = 0;
                    return [4 /*yield*/, Promise.all([boundedWorker(), boundedWorker(), boundedWorker(), boundedWorker()])];
                case 1:
                    _a.sent();
                    if (failure) {
                        throw failure;
                    }
                    return [2 /*return*/, folders];
            }
        });
    });
}
exports.loadAncestorFolders = loadAncestorFolders;

//# sourceMappingURL=LibraryBrowser.js.map
