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
var LIST_CONTROL_RUNTIME_CONFIG_EVENT = 'spse:listcontrol-runtime-config';
var LIST_CONTROL_RUNTIME_CONFIG_REQUEST_EVENT = 'spse:listcontrol-runtime-config-request';
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
function extractRenderRowsAndFields(data) {
    var directData = data && data.d && data.d.RenderListDataAsStream
        ? data.d.RenderListDataAsStream : data;
    var responseData = tryParseObject(directData) || {};
    var schema = tryParseObject(responseData.ListSchema) || tryParseObject(responseData.Schema) || {};
    var listData = tryParseObject(responseData.ListData) || responseData;
    var rows = toArray(responseData.Row);
    if (rows.length === 0) {
        rows = toArray(responseData.Rows);
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
        fields = toArray(responseData.Field);
    }
    return {
        rows: rows,
        fields: fields,
        nextHref: String(listData.NextHref || responseData.NextHref || '')
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
function padTwoDigits(value) {
    var text = String(value);
    return text.length < 2 ? '0' + text : text;
}
function applyDateFormatPattern(pattern, tokenValues) {
    return pattern.replace(/YYYY|YY|MMMM|MMM|MM|M|DD|D|dddd|ddd|HH|H|hh|h|mm|ss|tt/g, function (token) {
        return tokenValues[token] !== undefined ? tokenValues[token] : token;
    });
}
function applyTextCase(text, textCase) {
    if (textCase === 'upper') {
        return text.toUpperCase();
    }
    if (textCase === 'lower') {
        return text.toLowerCase();
    }
    return text;
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
        _this._fieldDisplayFormatMap = {};
        _this._loadRowsRequestId = 0;
        _this._stickyHeaderSourceHtml = '';
        _this._stickyHeaderVisible = false;
        _this._pagingEndpoint = '';
        _this._pagingRequestBody = undefined;
        _this._pagingRuntimeFilterFieldNames = [];
        _this.setDisplayFormFrameRef = function (frame) {
            if (!frame) {
                return;
            }
            var dialogFrame = frame;
            dialogFrame.cancelPopUp = function () { return _this.closeDefaultDisplayForm(); };
            dialogFrame.commitPopup = function () { return _this.closeDefaultDisplayForm(); };
            dialogFrame.commonModalDialogClose = function () { return _this.closeDefaultDisplayForm(); };
        };
        _this._setTableWrapRef = function (el) {
            _this._tableWrapEl = el;
            _this.refreshTableViewport();
        };
        _this._setTableHeadRef = function (el) {
            _this._tableHeadEl = el;
            _this.updateStickyHeaderPosition();
        };
        _this._setStickyHeaderViewportRef = function (el) {
            _this._stickyHeaderViewportEl = el;
            if (el) {
                el.onclick = _this.handleStickyHeaderClick;
            }
            _this.updateStickyHeaderPosition();
        };
        _this.handleStickyHeaderClick = function (event) {
            if (!_this._stickyHeaderViewportEl || !_this._tableHeadEl) {
                return;
            }
            var target = event.target;
            while (target && target !== _this._stickyHeaderViewportEl && !target.getAttribute('data-lc-sticky-index')) {
                target = target.parentElement;
            }
            if (!target || target === _this._stickyHeaderViewportEl) {
                return;
            }
            event.preventDefault();
            event.stopPropagation();
            var sourceControls = _this.getStickyHeaderControls(_this._tableHeadEl);
            var sourceIndex = parseInt(String(target.getAttribute('data-lc-sticky-index') || ''), 10);
            var sourceControl = sourceControls[sourceIndex];
            if (!sourceControl) {
                return;
            }
            var isFilterButton = (' ' + String(target.className || '') + ' ').indexOf(' lc-header-action ') >= 0;
            if (isFilterButton) {
                _this._filterAnchorEl = target;
            }
            sourceControl.click();
            if (isFilterButton && typeof window !== 'undefined') {
                _this._filterAnchorEl = target;
                window.setTimeout(function () { return _this.updateActiveFilterPosition(); }, 0);
            }
        };
        _this.state = {
            selectedViewId: props.defaultViewId || _this.getInitialViewId(props.views),
            fields: [],
            rows: [],
            loading: true,
            loadingMore: false,
            nextPageHref: '',
            error: null,
            selectedItemId: 0,
            selectedMode: 'view',
            deleting: false,
            sortFieldName: '',
            sortDirection: '',
            activeFilterFieldName: '',
            filterPopoverStyle: {},
            activeFilterIsDate: false,
            draftFilterOperator: 'contains',
            draftFilterValue: '',
            draftFilterEndValue: '',
            datePickerTarget: '',
            datePickerMonth: '',
            columnFilters: {},
            currentPage: 0,
            collapsedGroupKeys: {},
            displayFormUrl: '',
            displayFormLoading: false,
            displayFormError: '',
            embeddedByReportForms: false,
            runtimeFilterJson: '',
            runtimeConfigOwner: '',
            attachmentCountsByItemId: {},
            showScrollArrows: false,
            scrollArrowTop: 0,
            scrollArrowLeft: 0,
            scrollArrowRight: 0,
        };
        _this._refreshEventHandler = _this.handleExternalRefresh.bind(_this);
        _this._runtimeConfigEventHandler = _this.handleRuntimeConfig.bind(_this);
        _this._scrollArrowResizeHandler = _this.refreshTableViewport.bind(_this);
        _this._scrollArrowScrollHandler = _this.refreshTableViewport.bind(_this);
        return _this;
    }
    ListControl.prototype.componentDidMount = function () {
        this.logDiagnostic('Component mounted. listName=' + String(this.props.listName || '(none)') + ', defaultViewId=' + String(this.props.defaultViewId || '(none)'));
        if (typeof window !== 'undefined' && window.addEventListener) {
            window.addEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
            window.addEventListener(LIST_CONTROL_RUNTIME_CONFIG_EVENT, this._runtimeConfigEventHandler);
            window.addEventListener('resize', this._scrollArrowResizeHandler);
            window.addEventListener('scroll', this._scrollArrowScrollHandler, true);
            var requestEvent;
            var requestDetail = { instanceId: String(this.props.context && this.props.context.instanceId || '').toLowerCase() };
            if (typeof window.CustomEvent === 'function') {
                requestEvent = new window.CustomEvent(LIST_CONTROL_RUNTIME_CONFIG_REQUEST_EVENT, { detail: requestDetail });
            }
            else {
                requestEvent = document.createEvent('CustomEvent');
                requestEvent.initCustomEvent(LIST_CONTROL_RUNTIME_CONFIG_REQUEST_EVENT, false, false, requestDetail);
            }
            window.dispatchEvent(requestEvent);
        }
        this.loadRows();
    };
    ListControl.prototype.componentWillUnmount = function () {
        this._loadRowsRequestId += 1;
        if (typeof window !== 'undefined' && window.removeEventListener) {
            window.removeEventListener(LIST_CONTROL_REFRESH_EVENT, this._refreshEventHandler);
            window.removeEventListener(LIST_CONTROL_RUNTIME_CONFIG_EVENT, this._runtimeConfigEventHandler);
            window.removeEventListener('resize', this._scrollArrowResizeHandler);
            window.removeEventListener('scroll', this._scrollArrowScrollHandler, true);
        }
    };
    ListControl.prototype.openItemLinkDialog = function (itemLinkUrl) {
        if (!itemLinkUrl) {
            return;
        }
        this.setState({
            displayFormUrl: appendQueryParam(itemLinkUrl, 'IsDlg', '1'),
            displayFormLoading: false,
            displayFormError: ''
        });
    };
    ListControl.prototype.openItemLinkWindow = function (itemLinkUrl) {
        var width = 1100;
        var height = 800;
        var left = typeof window !== 'undefined' ? Math.max(0, Math.round((window.screen.availWidth - width) / 2)) : 0;
        var top = typeof window !== 'undefined' ? Math.max(0, Math.round((window.screen.availHeight - height) / 2)) : 0;
        var openedWindow = window.open(itemLinkUrl, '_blank', 'noopener,noreferrer,resizable=yes,scrollbars=yes,width=' + width + ',height=' + height + ',left=' + left + ',top=' + top);
        if (openedWindow) {
            openedWindow.opener = null;
        }
    };
    ListControl.prototype.closeDefaultDisplayForm = function () {
        var _this = this;
        this.setState({ displayFormUrl: '', displayFormLoading: false, displayFormError: '' }, function () { return _this.loadRows(); });
    };
    ListControl.prototype.componentDidUpdate = function (prevProps, prevState) {
        var _this = this;
        if (prevProps.listName !== this.props.listName || prevProps.defaultViewId !== this.props.defaultViewId) {
            this.logDiagnostic('Props changed; resetting selection and reloading rows. listName=' + String(this.props.listName || '(none)') + ', viewId=' + String(this.props.defaultViewId || '(none)'));
            var nextSelectedViewId = this.props.defaultViewId || this.getInitialViewId(this.props.views);
            var selectedViewWillChange = nextSelectedViewId !== this.state.selectedViewId;
            this.setState({
                selectedViewId: nextSelectedViewId,
                selectedItemId: 0,
                selectedMode: 'view',
                sortFieldName: '',
                sortDirection: '',
                activeFilterFieldName: '',
                draftFilterOperator: 'contains',
                draftFilterValue: '',
                columnFilters: {},
                currentPage: 0,
                collapsedGroupKeys: {},
            }, function () {
                _this.props.onSelectionChange(0, 'view');
                if (!selectedViewWillChange) {
                    _this.loadRows();
                }
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
        if (prevProps.fetchBatchSize !== this.props.fetchBatchSize) {
            this.loadRows();
            return;
        }
        if (prevState.runtimeFilterJson !== this.state.runtimeFilterJson) {
            this.loadRows();
            return;
        }
        if ((prevState.sortFieldName !== this.state.sortFieldName
            || prevState.sortDirection !== this.state.sortDirection
            || prevState.columnFilters !== this.state.columnFilters
            || prevState.runtimeFilterJson !== this.state.runtimeFilterJson
            || prevProps.filterJson !== this.props.filterJson) && this.state.nextPageHref) {
            this.loadAllRemainingRows();
        }
        if (typeof window !== 'undefined') {
            this._stickyHeaderSourceHtml = '';
            window.setTimeout(function () { return _this.refreshTableViewport(); }, 0);
        }
    };
    ListControl.prototype.getStickyHeaderControls = function (root) {
        var controls = root.querySelectorAll('button,input');
        var result = [];
        for (var index = 0; index < controls.length; index += 1) {
            var control = controls[index];
            var ancestor = control.parentElement;
            var insideFilter = false;
            while (ancestor && ancestor !== root) {
                if ((' ' + String(ancestor.className || '') + ' ').indexOf(' lc-filter-popover ') >= 0) {
                    insideFilter = true;
                    break;
                }
                ancestor = ancestor.parentElement;
            }
            if (!insideFilter) {
                result.push(control);
            }
        }
        return result;
    };
    ListControl.prototype.refreshTableViewport = function () {
        this.refreshScrollArrows();
        this.updateStickyHeaderPosition();
        this.updateActiveFilterPosition();
    };
    ListControl.prototype.refreshScrollArrows = function () {
        this.updateScrollArrowVisibility();
        this.updateScrollArrowPosition();
    };
    ListControl.prototype.updateScrollArrowVisibility = function () {
        if (!this._tableWrapEl) {
            return;
        }
        var overflowing = this._tableWrapEl.scrollWidth > this._tableWrapEl.clientWidth + 1;
        if (overflowing !== this.state.showScrollArrows) {
            this.setState({ showScrollArrows: overflowing });
        }
    };
    // Positions the arrows via fixed coordinates centered on the currently visible slice of the
    // table (intersection of its bounding rect with the viewport), so they stay reachable in the
    // middle of the view regardless of vertical scroll position instead of drifting to the bottom.
    ListControl.prototype.updateScrollArrowPosition = function () {
        if (!this._tableWrapEl || typeof window === 'undefined') {
            return;
        }
        var viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
        var viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
        var rect = this._tableWrapEl.getBoundingClientRect();
        var visibleTop = Math.max(rect.top, 0);
        var visibleBottom = Math.min(rect.bottom, viewportHeight);
        var centerY = Math.round((visibleTop + visibleBottom) / 2);
        var leftPx = Math.round(rect.left);
        var rightPx = Math.round(viewportWidth - rect.right);
        if (centerY !== this.state.scrollArrowTop || leftPx !== this.state.scrollArrowLeft || rightPx !== this.state.scrollArrowRight) {
            this.setState({ scrollArrowTop: centerY, scrollArrowLeft: leftPx, scrollArrowRight: rightPx });
        }
    };
    ListControl.prototype.updateStickyHeaderPosition = function () {
        if (!this._tableWrapEl || !this._tableHeadEl || !this._stickyHeaderViewportEl) {
            return;
        }
        var wrapRect = this._tableWrapEl.getBoundingClientRect();
        var headerHeight = this._tableHeadEl.getBoundingClientRect().height;
        var stickyTop = this.getStickyViewportTop(wrapRect);
        var shouldStick = wrapRect.top < stickyTop && wrapRect.bottom > stickyTop + headerHeight;
        if (shouldStick !== this._stickyHeaderVisible) {
            this._stickyHeaderVisible = shouldStick;
            this.logDiagnostic('Sticky header ' + (shouldStick ? 'shown' : 'hidden')
                + '. tableTop=' + String(Math.round(wrapRect.top))
                + ', tableBottom=' + String(Math.round(wrapRect.bottom))
                + ', stickyTop=' + String(Math.round(stickyTop))
                + ', headerHeight=' + String(Math.round(headerHeight)));
        }
        if (!shouldStick) {
            this._stickyHeaderViewportEl.style.display = 'none';
            return;
        }
        var sourceHtml = this._tableHeadEl.innerHTML;
        if (sourceHtml !== this._stickyHeaderSourceHtml || this._stickyHeaderViewportEl.children.length === 0) {
            var activeStickyIndex = this._filterAnchorEl && this._stickyHeaderViewportEl.contains(this._filterAnchorEl)
                ? parseInt(String(this._filterAnchorEl.getAttribute('data-lc-sticky-index') || ''), 10) : -1;
            this._stickyHeaderViewportEl.innerHTML = '';
            var stickyTable = document.createElement('table');
            stickyTable.className = 'lc-table lc-sticky-header-table';
            var stickyHead = this._tableHeadEl.cloneNode(true);
            var filterPopovers = stickyHead.querySelectorAll('.lc-filter-popover');
            for (var popoverIndex = 0; popoverIndex < filterPopovers.length; popoverIndex += 1) {
                var popover = filterPopovers[popoverIndex];
                if (popover.parentElement) {
                    popover.parentElement.removeChild(popover);
                }
            }
            var stickyControls = this.getStickyHeaderControls(stickyHead);
            for (var controlIndex = 0; controlIndex < stickyControls.length; controlIndex += 1) {
                stickyControls[controlIndex].setAttribute('data-lc-sticky-index', String(controlIndex));
            }
            if (this.state.activeFilterFieldName && activeStickyIndex >= 0 && activeStickyIndex < stickyControls.length) {
                this._filterAnchorEl = stickyControls[activeStickyIndex];
            }
            var stickyColGroup = document.createElement('colgroup');
            var stickyHeaderCells = stickyHead.children.length > 0 ? stickyHead.children[0].children : [];
            for (var colIndex = 0; colIndex < stickyHeaderCells.length; colIndex += 1) {
                stickyColGroup.appendChild(document.createElement('col'));
            }
            stickyTable.appendChild(stickyColGroup);
            stickyTable.appendChild(stickyHead);
            this._stickyHeaderViewportEl.appendChild(stickyTable);
            this._stickyHeaderSourceHtml = sourceHtml;
        }
        var tableElement = this._tableHeadEl.parentElement;
        var clonedTable = this._stickyHeaderViewportEl.children[0];
        var sourceCells = this._tableHeadEl.children.length > 0 ? this._tableHeadEl.children[0].children : [];
        var clonedColGroup = clonedTable.children[0];
        var clonedHead = clonedTable.children[1];
        var clonedCells = clonedHead && clonedHead.children.length > 0 ? clonedHead.children[0].children : [];
        var clonedColumns = clonedColGroup ? clonedColGroup.children : [];
        var totalTableWidth = 0;
        for (var cellIndex = 0; cellIndex < sourceCells.length && cellIndex < clonedCells.length; cellIndex += 1) {
            var measuredCellWidth = sourceCells[cellIndex].getBoundingClientRect().width;
            var cellWidth = measuredCellWidth.toFixed(2) + 'px';
            var clonedCell = clonedCells[cellIndex];
            clonedCell.style.width = '';
            clonedCell.style.minWidth = '';
            clonedCell.style.maxWidth = '';
            if (cellIndex < clonedColumns.length) {
                clonedColumns[cellIndex].style.width = cellWidth;
            }
            totalTableWidth += measuredCellWidth;
        }
        clonedTable.style.width = totalTableWidth.toFixed(2) + 'px';
        clonedTable.style.left = Math.round(-this._tableWrapEl.scrollLeft) + 'px';
        this._stickyHeaderViewportEl.style.display = 'block';
        this._stickyHeaderViewportEl.style.top = Math.round(stickyTop) + 'px';
        this._stickyHeaderViewportEl.style.left = Math.round(wrapRect.left + this._tableWrapEl.clientLeft) + 'px';
        this._stickyHeaderViewportEl.style.width = Math.round(this._tableWrapEl.clientWidth) + 'px';
        this._stickyHeaderViewportEl.style.height = Math.round(headerHeight) + 'px';
    };
    ListControl.prototype.getStickyViewportTop = function (wrapRect) {
        if (typeof document === 'undefined' || typeof window === 'undefined' || !document.elementFromPoint) {
            return 0;
        }
        var viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
        var probeX = Math.max(1, Math.min((window.innerWidth || document.documentElement.clientWidth || 0) - 1, Math.round((wrapRect.left + wrapRect.right) / 2)));
        var stickyTop = this.getScrollViewportTop(viewportHeight);
        for (var probeCount = 0; probeCount < 6 && stickyTop < Math.min(240, viewportHeight); probeCount += 1) {
            var probeY = Math.max(1, stickyTop + 1);
            var element = document.elementFromPoint(probeX, probeY);
            var chromeElement = undefined;
            while (element && element !== document.body) {
                var position = window.getComputedStyle(element).position;
                var elementRect = element.getBoundingClientRect();
                if (element !== this._stickyHeaderViewportEl && (position === 'fixed' || position === 'sticky') && elementRect.top <= probeY && elementRect.bottom > probeY) {
                    chromeElement = element;
                }
                element = element.parentElement;
            }
            if (!chromeElement) {
                break;
            }
            var chromeBottom = Math.round(chromeElement.getBoundingClientRect().bottom);
            if (chromeBottom <= stickyTop) {
                break;
            }
            stickyTop = chromeBottom;
        }
        return stickyTop;
    };
    ListControl.prototype.getScrollViewportTop = function (viewportHeight) {
        if (!this._tableWrapEl || typeof window === 'undefined') {
            return 0;
        }
        var stickyTop = 0;
        var ancestor = this._tableWrapEl.parentElement;
        while (ancestor && ancestor !== document.body && ancestor !== document.documentElement) {
            var computedStyle = window.getComputedStyle(ancestor);
            var overflowY = String(computedStyle.overflowY || '').toLowerCase();
            var canScrollVertically = (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay')
                && ancestor.scrollHeight > ancestor.clientHeight + 1;
            if (canScrollVertically) {
                var ancestorRect = ancestor.getBoundingClientRect();
                if (ancestorRect.bottom > 0 && ancestorRect.top < viewportHeight) {
                    stickyTop = Math.max(stickyTop, Math.max(0, Math.round(ancestorRect.top)));
                }
            }
            ancestor = ancestor.parentElement;
        }
        return stickyTop;
    };
    ListControl.prototype.scrollTableHorizontally = function (direction) {
        if (!this._tableWrapEl) {
            return;
        }
        var amount = Math.max(120, Math.round(this._tableWrapEl.clientWidth * 0.6));
        var nextLeft = this._tableWrapEl.scrollLeft + (amount * direction);
        if (this._tableWrapEl.scrollTo) {
            this._tableWrapEl.scrollTo({ left: nextLeft, behavior: 'smooth' });
        }
        else {
            this._tableWrapEl.scrollLeft = nextLeft;
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
    ListControl.prototype.getSelectedViewLabel = function () {
        var views = this.props.views || [];
        for (var i = 0; i < views.length; i += 1) {
            if (String(views[i].key) === String(this.state.selectedViewId)) {
                return views[i].text;
            }
        }
        return '';
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
    ListControl.prototype.handleRuntimeConfig = function (event) {
        var detail = event && event.detail ? event.detail : {};
        var targetInstanceId = String(detail.instanceId || '').toLowerCase();
        var currentInstanceId = String(this.props.context && this.props.context.instanceId || '').toLowerCase();
        if (!targetInstanceId || targetInstanceId !== currentInstanceId) {
            return;
        }
        var owner = String(detail.owner || '');
        if (detail.active === false) {
            if (this.state.runtimeConfigOwner && owner && this.state.runtimeConfigOwner !== owner) {
                return;
            }
            this.setState({ embeddedByReportForms: false, runtimeFilterJson: '', runtimeConfigOwner: '' });
            return;
        }
        this.setState({
            embeddedByReportForms: true,
            runtimeFilterJson: String(detail.filterJson || ''),
            runtimeConfigOwner: owner
        });
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
    ListControl.prototype.buildMinimalViewXml = function (viewQuery, viewFieldNames, rowLimit, scope) {
        var queryText = String(viewQuery || '').trim();
        var queryDocument = new DOMParser().parseFromString(/^<Query(?:\s|>)/i.test(queryText) ? queryText : '<Query>' + queryText + '</Query>', 'text/xml');
        if (queryDocument.getElementsByTagName('parsererror').length > 0 || queryDocument.getElementsByTagName('Query').length === 0) {
            throw new Error('The selected SharePoint view query is invalid.');
        }
        var xmlDocument = new DOMParser().parseFromString('<View><Query/><ViewFields/></View>', 'text/xml');
        var viewElement = xmlDocument.getElementsByTagName('View')[0];
        var existingQuery = xmlDocument.getElementsByTagName('Query')[0];
        var queryElement = xmlDocument.importNode(queryDocument.getElementsByTagName('Query')[0], true);
        viewElement.replaceChild(queryElement, existingQuery);
        var viewFieldsElement = xmlDocument.getElementsByTagName('ViewFields')[0];
        for (var fieldIndex = 0; fieldIndex < viewFieldNames.length; fieldIndex += 1) {
            var fieldName = String(viewFieldNames[fieldIndex] || '');
            if (fieldName) {
                var fieldRef = xmlDocument.createElement('FieldRef');
                fieldRef.setAttribute('Name', fieldName);
                viewFieldsElement.appendChild(fieldRef);
            }
        }
        var normalizedScope = String(scope === undefined || scope === null ? '' : scope).toLowerCase();
        var scopeNames = {
            '1': 'Recursive',
            '2': 'RecursiveAll',
            '3': 'FilesOnly',
            'recursive': 'Recursive',
            'recursiveall': 'RecursiveAll',
            'filesonly': 'FilesOnly'
        };
        if (scopeNames[normalizedScope]) {
            viewElement.setAttribute('Scope', scopeNames[normalizedScope]);
        }
        if (rowLimit > 0) {
            var rowLimitElement = xmlDocument.createElement('RowLimit');
            rowLimitElement.setAttribute('Paged', 'TRUE');
            rowLimitElement.appendChild(xmlDocument.createTextNode(String(rowLimit)));
            viewElement.appendChild(rowLimitElement);
        }
        return new XMLSerializer().serializeToString(xmlDocument);
    };
    ListControl.prototype.addFieldsToViewXml = function (viewXml, fieldNames) {
        if (!viewXml || fieldNames.length === 0) {
            return viewXml;
        }
        var xmlDocument = new DOMParser().parseFromString(viewXml, 'text/xml');
        var viewElement = xmlDocument.getElementsByTagName('View')[0];
        if (!viewElement || xmlDocument.getElementsByTagName('parsererror').length > 0) {
            return viewXml;
        }
        var viewFieldsElements = viewElement.getElementsByTagName('ViewFields');
        var viewFieldsElement = viewFieldsElements.length > 0 ? viewFieldsElements[0] : xmlDocument.createElement('ViewFields');
        if (viewFieldsElements.length === 0) {
            viewElement.appendChild(viewFieldsElement);
        }
        var existingFields = {};
        var fieldRefs = viewFieldsElement.getElementsByTagName('FieldRef');
        for (var i = 0; i < fieldRefs.length; i += 1) {
            existingFields[String(fieldRefs[i].getAttribute('Name') || '').toLowerCase()] = true;
        }
        for (var j = 0; j < fieldNames.length; j += 1) {
            var fieldName = String(fieldNames[j] || '').trim();
            if (!fieldName || existingFields[fieldName.toLowerCase()]) {
                continue;
            }
            var fieldRef = xmlDocument.createElement('FieldRef');
            fieldRef.setAttribute('Name', fieldName);
            viewFieldsElement.appendChild(fieldRef);
            existingFields[fieldName.toLowerCase()] = true;
        }
        return new XMLSerializer().serializeToString(xmlDocument);
    };
    ListControl.prototype.loadSelectedViewXml = function (selectedViewId, viewFieldNames) {
        return __awaiter(this, void 0, void 0, function () {
            var webUrl, listPath, viewIds, urls, i, encoded, normalized, j, response, data, viewData, rowLimit, viewXml, viewXmlError_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!selectedViewId) {
                            return [2 /*return*/, ''];
                        }
                        webUrl = this.getWebUrl();
                        listPath = "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')";
                        viewIds = this.buildViewIdCandidates(selectedViewId);
                        urls = [];
                        for (i = 0; i < viewIds.length; i += 1) {
                            encoded = encodeURIComponent(viewIds[i]);
                            normalized = encodeURIComponent(trimGuidBraces(viewIds[i]));
                            urls.push(webUrl + listPath + "/views/getById('" + encoded + "')?$select=ViewQuery,RowLimit,Scope");
                            urls.push(webUrl + listPath + "/views(guid'" + normalized + "')?$select=ViewQuery,RowLimit,Scope");
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
                        viewData = data && data.d ? data.d : data;
                        if (viewData && viewData.ViewQuery !== undefined && viewData.ViewQuery !== null) {
                            rowLimit = parseInt(String(viewData.RowLimit || ''), 10);
                            viewXml = this.buildMinimalViewXml(String(viewData.ViewQuery), viewFieldNames, isNaN(rowLimit) ? 0 : rowLimit, viewData.Scope);
                            this.logDiagnostic('Loaded minimal selected view CAML. HasFilter=' + String(/<Where(?:\s|>)/i.test(viewXml)) + ', hasSort=' + String(/<OrderBy(?:\s|>)/i.test(viewXml)) + ', fields=' + String(viewFieldNames.length) + '.');
                            return [2 /*return*/, viewXml];
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        viewXmlError_1 = _a.sent();
                        this.logDiagnostic('loadSelectedViewXml: Attempt failed for url=' + urls[j] + ': ' + (viewXmlError_1 && viewXmlError_1.message ? viewXmlError_1.message : String(viewXmlError_1)));
                        return [3 /*break*/, 6];
                    case 6:
                        j += 1;
                        return [3 /*break*/, 1];
                    case 7: throw new Error('Failed to load the selected SharePoint view definition.');
                }
            });
        });
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
        var visibleStateFields = this.state.fields.filter(function (field) {
            return field.RuntimeFilterOnly !== true;
        });
        if (this.state.selectedViewId !== this.props.defaultViewId || !this.props.viewColumns || this.props.viewColumns.length === 0) {
            return visibleStateFields;
        }
        var byName = {};
        for (var i = 0; i < visibleStateFields.length; i += 1) {
            var field = visibleStateFields[i];
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
        var addedFieldNames = {};
        for (var j = 0; j < this.props.viewColumns.length; j += 1) {
            var column = this.props.viewColumns[j];
            var columnKey = String(column.fieldName || '').toLowerCase();
            // Guard against a duplicate saved column (e.g. a SharePoint view listing the same field
            // twice) rendering the same value twice with no visible separator between cells.
            if (!columnKey || addedFieldNames[columnKey]) {
                continue;
            }
            var match = byName[columnKey];
            if (match) {
                addedFieldNames[columnKey] = true;
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
            var endpoint, response, data, fields, requested, i, map, j, field, internalName, displayFormat, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!viewFieldNames || viewFieldNames.length === 0) {
                            return [2 /*return*/, {}];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/fields?$select=InternalName,TypeAsString,DisplayFormat";
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
                            map[internalName.toLowerCase()] = String(field.TypeAsString || '');
                            if (String(field.TypeAsString || '').toLowerCase() === 'datetime') {
                                displayFormat = parseInt(String(field.DisplayFormat), 10);
                                this._fieldDisplayFormatMap[internalName.toLowerCase()] = isNaN(displayFormat) ? 1 : displayFormat;
                            }
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
    ListControl.prototype.applyFieldTypes = function (fields, typeMap) {
        return fields.map(function (field) {
            var internalName = String(field.RealFieldName || field.Name || '');
            return __assign({}, field, { TypeAsString: typeMap[internalName] || typeMap[internalName.toLowerCase()] || field.TypeAsString || '', DisplayFormat: this._fieldDisplayFormatMap[internalName.toLowerCase()] });
        }.bind(this));
    };
    ListControl.prototype.getRowFieldValue = function (row, field) {
        var fieldName = field.Name || field.RealFieldName || '';
        var value = row[fieldName];
        if ((value === undefined || value === null || value === '') && field.RealFieldName) {
            value = row[field.RealFieldName];
        }
        return value;
    };
    ListControl.prototype.getUrlCellValue = function (row, field) {
        var fieldType = String(field.TypeAsString || '').toLowerCase();
        if (fieldType !== 'url' && fieldType !== 'hyperlink') {
            return null;
        }
        var rawValue = this.getRowFieldValue(row, field);
        if (rawValue === undefined || rawValue === null || rawValue === '') {
            return null;
        }
        var href = '';
        var text = '';
        if (typeof rawValue === 'object') {
            href = String(rawValue.Url || rawValue.url || '').trim();
            text = String(rawValue.Description || rawValue.description || '').trim();
        }
        else {
            var rawText = String(rawValue).trim();
            if (rawText.indexOf('<') >= 0 && rawText.indexOf('>') >= 0) {
                var container = document.createElement('div');
                container.innerHTML = rawText;
                var anchor = container.querySelector('a');
                if (anchor) {
                    href = String(anchor.getAttribute('href') || '').trim();
                    text = String(anchor.textContent || '').trim();
                }
            }
            if (!href) {
                var descriptionSeparator = rawText.indexOf(', ');
                href = descriptionSeparator > 0 ? rawText.substring(0, descriptionSeparator).trim() : rawText;
                text = descriptionSeparator > 0 ? rawText.substring(descriptionSeparator + 2).trim() : '';
            }
        }
        if (!href || /^\s*(?:javascript|data|vbscript):/i.test(href)) {
            return null;
        }
        return { href: href, text: text || href };
    };
    ListControl.prototype.stringifyCellValue = function (value) {
        if (value === undefined || value === null || value === '') {
            return '';
        }
        if (typeof value === 'boolean') {
            return value ? strings.RuntimeBooleanYes : strings.RuntimeBooleanNo;
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
        if (typeof value === 'string') {
            // RenderListDataAsStream can return Person/Lookup fields as JSON-encoded strings.
            var trimmedValue = value.trim();
            if (trimmedValue.length > 0 && (trimmedValue.charAt(0) === '{' || trimmedValue.charAt(0) === '[')) {
                var parsedJsonValue = tryParseObject(trimmedValue);
                if (parsedJsonValue && typeof parsedJsonValue === 'object') {
                    return this.stringifyCellValue(parsedJsonValue);
                }
            }
        }
        if (typeof value === 'object') {
            if (Array.isArray(value.results)) {
                return this.stringifyCellValue(value.results);
            }
            if (value.lookupValue !== undefined && value.lookupValue !== null) {
                return String(value.lookupValue);
            }
            if (value.DisplayName) {
                return String(value.DisplayName);
            }
            if (value.displayName) {
                return String(value.displayName);
            }
            if (value.Title) {
                return String(value.Title);
            }
            if (value.LookupValue) {
                return String(value.LookupValue);
            }
            if (value.title) {
                return String(value.title);
            }
            if (value.Name) {
                return String(value.Name);
            }
            if (value.name) {
                return String(value.name);
            }
            if (value.Email) {
                return String(value.Email);
            }
            if (value.email) {
                return String(value.email);
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
    ListControl.prototype.formatDateCellValue = function (value, field, row) {
        if (String(field.TypeAsString || '') === 'Attachments') {
            return String(this.getAttachmentCountForRow(row, value));
        }
        var rawValue = this.stringifyCellValue(value);
        if (String(field.TypeAsString || '').toLowerCase() !== 'datetime') {
            return rawValue;
        }
        var dateOnlyMatch = field.DisplayFormat === 0 ? /^(\d{4})-(\d{2})-(\d{2})/.exec(rawValue) : null;
        var dateValue = new Date(rawValue);
        if (isNaN(dateValue.getTime())) {
            return rawValue;
        }
        var month = dateOnlyMatch ? dateOnlyMatch[2] : String(dateValue.getMonth() + 1);
        var day = dateOnlyMatch ? dateOnlyMatch[3] : String(dateValue.getDate());
        var year = dateOnlyMatch ? dateOnlyMatch[1] : String(dateValue.getFullYear());
        var monthPadded = padTwoDigits(Number(month));
        var dayPadded = padTwoDigits(Number(day));
        // Date-only fields build named tokens from the corrected local date to avoid a UTC day shift.
        var namedTokenDate = dateOnlyMatch ? new Date(Number(year), Number(month) - 1, Number(day)) : dateValue;
        var hours = dateValue.getHours();
        var twelveHour = hours % 12 || 12;
        var tokenValues = {
            YYYY: year,
            YY: year.length > 2 ? year.substring(year.length - 2) : year,
            MMMM: namedTokenDate.toLocaleString(undefined, { month: 'long' }),
            MMM: namedTokenDate.toLocaleString(undefined, { month: 'short' }),
            MM: monthPadded,
            M: String(Number(month)),
            DD: dayPadded,
            D: String(Number(day)),
            dddd: namedTokenDate.toLocaleString(undefined, { weekday: 'long' }),
            ddd: namedTokenDate.toLocaleString(undefined, { weekday: 'short' }),
            HH: padTwoDigits(hours),
            H: String(hours),
            hh: padTwoDigits(twelveHour),
            h: String(twelveHour),
            mm: padTwoDigits(dateValue.getMinutes()),
            ss: padTwoDigits(dateValue.getSeconds()),
            tt: hours >= 12 ? 'PM' : 'AM'
        };
        var dateText;
        if (this.props.dateDisplayFormat === 'custom') {
            dateText = applyTextCase(applyDateFormatPattern(String(this.props.dateCustomFormat || '').trim() || 'MM/DD/YYYY', tokenValues), String(this.props.dateCustomFormatCase || ''));
        }
        else if (this.props.dateDisplayFormat === 'dmy') {
            dateText = dayPadded + '/' + monthPadded + '/' + year;
        }
        else if (this.props.dateDisplayFormat === 'ymd') {
            dateText = year + '-' + monthPadded + '-' + dayPadded;
        }
        else {
            dateText = monthPadded + '/' + dayPadded + '/' + year;
        }
        if (field.DisplayFormat === 0) {
            return dateText;
        }
        var minutes = padTwoDigits(dateValue.getMinutes());
        if (this.props.timeDisplayFormat === 'custom') {
            // Custom time format fully controls the combined output so date/time tokens can be freely interleaved.
            return applyTextCase(applyDateFormatPattern(String(this.props.timeCustomFormat || '').trim() || 'MM/DD/YYYY HH:mm', tokenValues), String(this.props.timeCustomFormatCase || ''));
        }
        if (this.props.timeDisplayFormat === '12hour') {
            var period = hours >= 12 ? 'PM' : 'AM';
            return dateText + ' ' + (twelveHour < 10 ? '0' : '') + String(twelveHour) + ':' + minutes + ' ' + period;
        }
        return dateText + ' ' + (hours < 10 ? '0' : '') + String(hours) + ':' + minutes;
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
            if (this.getRowItemId(row) > 0) {
                filtered.push(row);
                continue;
            }
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
            // Rows without a standard item ID still require visible content.
            if (hasVisibleValue) {
                filtered.push(row);
            }
        }
        return filtered;
    };
    ListControl.prototype.loadRowsFromItemsEndpoint = function (viewFieldNames, itemIds) {
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
                        if (itemIds && itemIds.length > 0) {
                            endpoint += '&$filter=' + encodeURIComponent(itemIds.map(function (itemId) {
                                return 'ID eq ' + String(itemId);
                            }).join(' or '));
                        }
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
                        if (itemIds && itemIds.length > 0 && rows.length > 1) {
                            rows.sort(function (left, right) {
                                return itemIds.indexOf(toPositiveInt(left.ID || left.Id || left.id))
                                    - itemIds.indexOf(toPositiveInt(right.ID || right.Id || right.id));
                            });
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
    ListControl.prototype.getRuntimeFilterFieldNames = function () {
        var conditions = [];
        try {
            var parsed = JSON.parse(String(this.state.runtimeFilterJson || ''));
            conditions = Array.isArray(parsed) ? parsed : [];
        }
        catch (_parseError) {
            conditions = [];
        }
        var fieldNames = [];
        var seen = {};
        for (var i = 0; i < conditions.length; i += 1) {
            var fieldName = String(conditions[i] && conditions[i].field || '').trim();
            var normalizedName = fieldName.toLowerCase();
            if (fieldName && !seen[normalizedName]) {
                fieldNames.push(fieldName);
                seen[normalizedName] = true;
            }
        }
        return fieldNames;
    };
    ListControl.prototype.mergeHydratedRows = function (rows, hydratedRows) {
        var _this = this;
        var hydratedById = {};
        for (var i = 0; i < hydratedRows.length; i += 1) {
            var hydratedId = this.getRowItemId(hydratedRows[i]);
            if (hydratedId > 0) {
                hydratedById[hydratedId] = hydratedRows[i];
            }
        }
        return rows.map(function (row) {
            var itemId = _this.getRowItemId(row);
            return hydratedById[itemId] ? Object.assign({}, row, hydratedById[itemId]) : row;
        });
    };
    ListControl.prototype.loadAttachmentCounts = function (itemIds) {
        return __awaiter(this, void 0, void 0, function () {
            var counts, uniqueItemIds, endpoint, response, data, items, i, item, itemId, attachmentFiles, attachmentCountError_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        counts = {};
                        uniqueItemIds = itemIds.filter(function (itemId, index) {
                            return itemId > 0 && itemIds.indexOf(itemId) === index;
                        });
                        if (uniqueItemIds.length === 0) {
                            return [2 /*return*/, counts];
                        }
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 4, , 5]);
                        endpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/items"
                            + '?$top=200&$select=ID,AttachmentFiles&$expand=AttachmentFiles'
                            + '&$filter=' + encodeURIComponent(uniqueItemIds.map(function (itemId) {
                            return 'ID eq ' + String(itemId);
                        }).join(' or '));
                        return [4 /*yield*/, this.getJsonWithFallback(endpoint)];
                    case 2:
                        response = _a.sent();
                        if (!response.ok) {
                            return [2 /*return*/, counts];
                        }
                        return [4 /*yield*/, response.json()];
                    case 3:
                        data = _a.sent();
                        items = toArray(data.value);
                        if (items.length === 0) {
                            items = toArray(data && data.d && data.d.results);
                        }
                        for (i = 0; i < items.length; i += 1) {
                            item = items[i];
                            itemId = toPositiveInt(item.ID || item.Id || item.id);
                            attachmentFiles = item.AttachmentFiles && item.AttachmentFiles.results ? item.AttachmentFiles.results : item.AttachmentFiles;
                            counts[itemId] = Array.isArray(attachmentFiles) ? attachmentFiles.length : 0;
                        }
                        return [3 /*break*/, 5];
                    case 4:
                        attachmentCountError_1 = _a.sent();
                        this.logDiagnostic('loadAttachmentCounts failed: ' + (attachmentCountError_1 && attachmentCountError_1.message ? attachmentCountError_1.message : String(attachmentCountError_1)));
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/, counts];
                }
            });
        });
    };
    ListControl.prototype.loadRows = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var requestId, baseEndpoint, selectedViewId, viewFieldNames, selectedViewXml, runtimeFilterFieldNames, body, selectedRows, selectedFields, nextPageHref, lastError, hadSuccessfulResponse, requestUrls, requestIndex, requestUrl, response, errorText, _readError_1, data, extracted, rows, fields, itemsFallback, runtimeSupportFields, runtimeItemIds, runtimeItems, visibleFields, existingFieldNames, fieldTitleMap, visibleFieldNames, fieldTypeMap, renderableRows, hasAttachmentsField, attachmentCountsByItemId, _a, error_3, loadError;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (!this.props.listName) {
                            this.setState({ loading: false, error: null, fields: [], rows: [] });
                            return [2 /*return*/];
                        }
                        requestId = ++this._loadRowsRequestId;
                        this.logDiagnostic('Starting loadRows. requestId=' + String(requestId) + ', listName=' + String(this.props.listName) + ', selectedViewId=' + String(this.state.selectedViewId || '(none)'));
                        this.setState({ loading: true, loadingMore: false, nextPageHref: '', error: null });
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 23, , 24]);
                        baseEndpoint = this.getWebUrl() + "/_api/web/lists/getByTitle('" + escapeODataText(this.props.listName) + "')/RenderListDataAsStream";
                        selectedViewId = this.state.selectedViewId;
                        return [4 /*yield*/, this.loadSelectedViewFieldNames(selectedViewId)];
                    case 2:
                        viewFieldNames = _b.sent();
                        return [4 /*yield*/, this.loadSelectedViewXml(selectedViewId, viewFieldNames)];
                    case 3:
                        selectedViewXml = _b.sent();
                        runtimeFilterFieldNames = this.getRuntimeFilterFieldNames();
                        this._pagingRuntimeFilterFieldNames = runtimeFilterFieldNames.slice(0);
                        body = {
                            parameters: {
                                RenderOptions: 7
                            }
                        };
                        body.parameters.ViewXml = this.applyFetchBatchSize(selectedViewId
                            ? this.addFieldsToViewXml(selectedViewXml || '<View></View>', runtimeFilterFieldNames)
                            : (selectedViewXml || '<View></View>'));
                        this._pagingEndpoint = baseEndpoint;
                        this._pagingRequestBody = body;
                        selectedRows = [];
                        selectedFields = [];
                        nextPageHref = '';
                        lastError = '';
                        hadSuccessfulResponse = false;
                        requestUrls = [baseEndpoint];
                        requestIndex = 0;
                        _b.label = 4;
                    case 4:
                        if (!(requestIndex < requestUrls.length)) return [3 /*break*/, 13];
                        requestUrl = requestUrls[requestIndex];
                        return [4 /*yield*/, this.postJsonWithFallback(requestUrl, body)];
                    case 5:
                        response = _b.sent();
                        if (!!response.ok) return [3 /*break*/, 10];
                        errorText = '';
                        _b.label = 6;
                    case 6:
                        _b.trys.push([6, 8, , 9]);
                        return [4 /*yield*/, response.text()];
                    case 7:
                        errorText = _b.sent();
                        return [3 /*break*/, 9];
                    case 8:
                        _readError_1 = _b.sent();
                        errorText = '';
                        return [3 /*break*/, 9];
                    case 9:
                        lastError = 'Failed to load list view data. HTTP ' + String(response.status) + ' ' + response.statusText + (errorText ? (': ' + errorText) : '');
                        return [3 /*break*/, 12];
                    case 10:
                        hadSuccessfulResponse = true;
                        return [4 /*yield*/, response.json()];
                    case 11:
                        data = _b.sent();
                        extracted = extractRenderRowsAndFields(data);
                        nextPageHref = extracted.nextHref;
                        this.logDiagnostic('RenderListDataAsStream parsed. rows=' + String(extracted.rows.length) + ', fields=' + String(extracted.fields.length));
                        if (selectedFields.length === 0 && extracted.fields.length > 0) {
                            selectedFields = extracted.fields;
                        }
                        if (extracted.rows.length > 0 || requestIndex === requestUrls.length - 1) {
                            selectedRows = extracted.rows;
                            if (selectedFields.length === 0) {
                                selectedFields = extracted.fields;
                            }
                            return [3 /*break*/, 13];
                        }
                        _b.label = 12;
                    case 12:
                        requestIndex += 1;
                        return [3 /*break*/, 4];
                    case 13:
                        if (!hadSuccessfulResponse) {
                            throw new Error(lastError || 'Failed to load list view data.');
                        }
                        rows = selectedRows;
                        fields = selectedFields;
                        if (!(rows.length === 0 && !selectedViewId)) return [3 /*break*/, 15];
                        return [4 /*yield*/, this.loadRowsFromItemsEndpoint(viewFieldNames)];
                    case 14:
                        itemsFallback = _b.sent();
                        rows = itemsFallback.rows;
                        if (fields.length === 0) {
                            fields = itemsFallback.fields;
                        }
                        _b.label = 15;
                    case 15:
                        runtimeSupportFields = [];
                        if (!(runtimeFilterFieldNames.length > 0 && rows.length > 0)) return [3 /*break*/, 17];
                        runtimeItemIds = rows.map(function (row) { return _this.getRowItemId(row); }).filter(function (itemId) { return itemId > 0; });
                        return [4 /*yield*/, this.loadRowsFromItemsEndpoint(runtimeFilterFieldNames, runtimeItemIds)];
                    case 16:
                        runtimeItems = _b.sent();
                        rows = this.mergeHydratedRows(rows, runtimeItems.rows);
                        runtimeSupportFields = runtimeItems.fields;
                        _b.label = 17;
                    case 17:
                        visibleFields = this.getFieldsForConsumption(fields, viewFieldNames);
                        existingFieldNames = {};
                        visibleFields.forEach(function (field) {
                            existingFieldNames[String(field.RealFieldName || field.Name || '').toLowerCase()] = true;
                            existingFieldNames[String(field.Name || '').toLowerCase()] = true;
                        });
                        runtimeSupportFields.forEach(function (field) {
                            var supportFieldName = String(field.RealFieldName || field.Name || '').toLowerCase();
                            if (runtimeFilterFieldNames.map(function (name) { return name.toLowerCase(); }).indexOf(supportFieldName) >= 0 && !existingFieldNames[supportFieldName]) {
                                visibleFields.push(Object.assign({}, field, { RuntimeFilterOnly: true }));
                                existingFieldNames[supportFieldName] = true;
                            }
                        });
                        return [4 /*yield*/, this.loadListFieldTitleMap()];
                    case 18:
                        fieldTitleMap = _b.sent();
                        visibleFieldNames = visibleFields.map(function (field) { return _this.getFieldKey(field); });
                        return [4 /*yield*/, this.loadListFieldTypeMap(visibleFieldNames)];
                    case 19:
                        fieldTypeMap = _b.sent();
                        visibleFields = this.applyFieldDisplayNames(visibleFields, fieldTitleMap);
                        visibleFields = this.applyFieldTypes(visibleFields, fieldTypeMap);
                        renderableRows = this.filterRenderableRows(rows, visibleFields);
                        if (requestId !== this._loadRowsRequestId) {
                            this.logDiagnostic('Ignoring stale loadRows result. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
                            return [2 /*return*/];
                        }
                        hasAttachmentsField = visibleFields.some(function (field) {
                            return String(field.TypeAsString || '') === 'Attachments';
                        });
                        if (!hasAttachmentsField) return [3 /*break*/, 21];
                        return [4 /*yield*/, this.loadAttachmentCounts(renderableRows.map(function (row) { return _this.getRowItemId(row); }))];
                    case 20:
                        _a = _b.sent();
                        return [3 /*break*/, 22];
                    case 21:
                        _a = {};
                        _b.label = 22;
                    case 22:
                        attachmentCountsByItemId = _a;
                        if (requestId !== this._loadRowsRequestId) {
                            this.logDiagnostic('Ignoring stale loadRows result after attachment count fetch. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
                            return [2 /*return*/];
                        }
                        this.setState({
                            fields: visibleFields,
                            rows: renderableRows,
                            loadingMore: false,
                            nextPageHref: nextPageHref,
                            attachmentCountsByItemId: attachmentCountsByItemId,
                            loading: false,
                            error: null
                        }, function () {
                            if ((_this.parsePresetFilterConditions().length > 0 || _this.getGroupingConfig().enabled) && _this.state.nextPageHref) {
                                _this.loadAllRemainingRows();
                            }
                        });
                        this.logDiagnostic('loadRows completed. visibleFields=' + String(visibleFields.length) + ', renderableRows=' + String(renderableRows.length));
                        return [3 /*break*/, 24];
                    case 23:
                        error_3 = _b.sent();
                        loadError = error_3;
                        if (requestId !== this._loadRowsRequestId) {
                            this.logDiagnostic('Ignoring stale loadRows failure. requestId=' + String(requestId) + ', latestRequestId=' + String(this._loadRowsRequestId));
                            return [2 /*return*/];
                        }
                        this.setState({
                            loading: false,
                            loadingMore: false,
                            nextPageHref: '',
                            error: loadError && loadError.message ? loadError.message : 'Failed to load data.',
                            fields: [],
                            rows: []
                        });
                        this.logDiagnostic('loadRows failed: ' + (loadError && loadError.message ? loadError.message : String(loadError)));
                        return [3 /*break*/, 24];
                    case 24: return [2 /*return*/];
                }
            });
        });
    };
    ListControl.prototype.applyFetchBatchSize = function (viewXml) {
        var parser = new DOMParser();
        var xmlDocument = parser.parseFromString(String(viewXml || '<View></View>'), 'text/xml');
        var viewElement = xmlDocument.documentElement;
        var rowLimits = viewElement.getElementsByTagName('RowLimit');
        var rowLimit = rowLimits.length > 0 ? rowLimits[0] : xmlDocument.createElement('RowLimit');
        rowLimit.setAttribute('Paged', 'TRUE');
        while (rowLimit.firstChild) {
            rowLimit.removeChild(rowLimit.firstChild);
        }
        rowLimit.appendChild(xmlDocument.createTextNode(String(this.props.fetchBatchSize)));
        if (rowLimits.length === 0) {
            viewElement.appendChild(rowLimit);
        }
        return new XMLSerializer().serializeToString(xmlDocument);
    };
    ListControl.prototype.loadNextBatch = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var requestId, parameters, sourceParameters, parameterName, response, extracted, _a, nextRows, runtimeItemIds, runtimeItems, existingIds, uniqueRows, nextAttachmentCounts, hasAttachments, nextPageHref, error_4;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        if (this.state.loadingMore || !this.state.nextPageHref || !this._pagingEndpoint || !this._pagingRequestBody) {
                            return [2 /*return*/, false];
                        }
                        requestId = this._loadRowsRequestId;
                        this.setState({ loadingMore: true });
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 9, , 10]);
                        parameters = {};
                        sourceParameters = this._pagingRequestBody.parameters || {};
                        for (parameterName in sourceParameters) {
                            if (Object.prototype.hasOwnProperty.call(sourceParameters, parameterName)) {
                                parameters[parameterName] = sourceParameters[parameterName];
                            }
                        }
                        parameters.Paging = String(this.state.nextPageHref).replace(/^\?/, '');
                        return [4 /*yield*/, this.postJsonWithFallback(this._pagingEndpoint, { parameters: parameters })];
                    case 2:
                        response = _b.sent();
                        if (!response.ok) {
                            this.setState({ loadingMore: false });
                            return [2 /*return*/, false];
                        }
                        _a = extractRenderRowsAndFields;
                        return [4 /*yield*/, response.json()];
                    case 3:
                        extracted = _a.apply(void 0, [_b.sent()]);
                        if (requestId !== this._loadRowsRequestId) {
                            return [2 /*return*/, false];
                        }
                        nextRows = extracted.rows;
                        if (!(this._pagingRuntimeFilterFieldNames.length > 0 && nextRows.length > 0)) return [3 /*break*/, 5];
                        runtimeItemIds = nextRows.map(function (row) { return _this.getRowItemId(row); }).filter(function (itemId) { return itemId > 0; });
                        return [4 /*yield*/, this.loadRowsFromItemsEndpoint(this._pagingRuntimeFilterFieldNames, runtimeItemIds)];
                    case 4:
                        runtimeItems = _b.sent();
                        nextRows = this.mergeHydratedRows(nextRows, runtimeItems.rows);
                        _b.label = 5;
                    case 5:
                        nextRows = this.filterRenderableRows(nextRows, this.state.fields);
                        existingIds = {};
                        this.state.rows.forEach(function (row) { existingIds[_this.getRowItemId(row)] = true; });
                        uniqueRows = nextRows.filter(function (row) { return !existingIds[_this.getRowItemId(row)]; });
                        nextAttachmentCounts = {};
                        hasAttachments = this.state.fields.some(function (field) { return String(field.TypeAsString || '') === 'Attachments'; });
                        if (!hasAttachments) return [3 /*break*/, 7];
                        return [4 /*yield*/, this.loadAttachmentCounts(uniqueRows.map(function (row) { return _this.getRowItemId(row); }))];
                    case 6:
                        nextAttachmentCounts = _b.sent();
                        _b.label = 7;
                    case 7:
                        nextPageHref = extracted.nextHref === this.state.nextPageHref && uniqueRows.length === 0 ? '' : extracted.nextHref;
                        return [4 /*yield*/, new Promise(function (resolve) { return _this.setState({
                                rows: _this.state.rows.concat(uniqueRows),
                                attachmentCountsByItemId: __assign({}, _this.state.attachmentCountsByItemId, nextAttachmentCounts),
                                nextPageHref: nextPageHref,
                                loadingMore: false
                            }, function () { return resolve(true); }); })];
                    case 8:
                        _b.sent();
                        return [2 /*return*/, uniqueRows.length > 0 || !!nextPageHref];
                    case 9:
                        error_4 = _b.sent();
                        this.logDiagnostic('Loading next batch failed: ' + (error_4 && error_4.message ? error_4.message : String(error_4)));
                        this.setState({ loadingMore: false });
                        return [2 /*return*/, false];
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    ListControl.prototype.loadAllRemainingRows = function () {
        var _this = this;
        if (this.state.loadingMore || !this.state.nextPageHref) {
            return;
        }
        this.loadNextBatch().then(function (loaded) {
            if (loaded && _this.state.nextPageHref) {
                _this.loadAllRemainingRows();
            }
        });
    };
    ListControl.prototype.goToNextPage = function (currentPage, pageCount) {
        var _this = this;
        if (currentPage < pageCount - 1) {
            var nextPage = currentPage + 1;
            this.setState({ currentPage: nextPage }, function () {
                if (_this.state.nextPageHref && nextPage >= pageCount - 2) {
                    _this.loadNextBatch();
                }
            });
            return;
        }
        if (this.state.nextPageHref && !this.state.loadingMore) {
            this.loadNextBatch().then(function (loaded) {
                if (loaded) {
                    _this.setState({ currentPage: currentPage + 1 });
                }
            });
        }
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
            var endpoint, response, error_5, deleteError;
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
                        error_5 = _a.sent();
                        deleteError = error_5;
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
        var text = this.formatDateCellValue(value, field, row);
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
        if (String(field.TypeAsString || '') === 'Attachments') {
            return String(this.getAttachmentCountForRow(row, value));
        }
        return this.stringifyCellValue(value).replace(/<[^>]*>/g, '').trim();
    };
    ListControl.prototype.getAttachmentCountForRow = function (row, rawAttachmentsValue) {
        var itemId = row ? this.getRowItemId(row) : 0;
        var knownCount = this.state.attachmentCountsByItemId[itemId];
        if (knownCount !== undefined) {
            return knownCount;
        }
        // Fall back to the boolean-only presence indicator if the exact count hasn't loaded yet.
        return rawAttachmentsValue === true ? 1 : 0;
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
    ListControl.prototype.openFilter = function (field, anchorElement) {
        var fieldKey = this.getFieldKey(field);
        if (!fieldKey) {
            return;
        }
        this._filterAnchorEl = anchorElement;
        var popoverStyle = this.getFilterPopoverStyle(anchorElement);
        var existing = this.state.columnFilters[fieldKey];
        var isDateField = String(field.TypeAsString || '').toLowerCase() === 'datetime';
        this.setState({
            activeFilterFieldName: fieldKey,
            filterPopoverStyle: popoverStyle,
            activeFilterIsDate: isDateField,
            draftFilterOperator: isDateField ? 'eq' : (existing ? existing.operator : 'contains'),
            draftFilterValue: existing ? existing.value : '',
            draftFilterEndValue: existing ? String(existing.endValue || '') : '',
            datePickerTarget: '',
            datePickerMonth: ''
        });
    };
    ListControl.prototype.getFilterPopoverStyle = function (anchorElement) {
        var anchorRect = anchorElement.getBoundingClientRect();
        var viewportWidth = document.documentElement.clientWidth || window.innerWidth;
        var viewportHeight = document.documentElement.clientHeight || window.innerHeight;
        var popoverWidth = 220;
        var estimatedPopoverHeight = 190;
        var viewportMargin = 8;
        var popoverGap = 6;
        var left = anchorRect.right - popoverWidth;
        left = Math.max(viewportMargin, Math.min(left, viewportWidth - popoverWidth - viewportMargin));
        var availableBelow = viewportHeight - anchorRect.bottom - popoverGap - viewportMargin;
        var availableAbove = anchorRect.top - popoverGap - viewportMargin;
        var popoverStyle = {
            left: left,
            right: 'auto'
        };
        if (availableBelow < estimatedPopoverHeight && availableAbove > availableBelow) {
            popoverStyle.top = 'auto';
            popoverStyle.bottom = viewportHeight - anchorRect.top + popoverGap;
            popoverStyle.maxHeight = Math.max(120, availableAbove);
        }
        else {
            popoverStyle.top = anchorRect.bottom + popoverGap;
            popoverStyle.bottom = 'auto';
            popoverStyle.maxHeight = Math.max(120, availableBelow);
        }
        return popoverStyle;
    };
    ListControl.prototype.updateActiveFilterPosition = function () {
        if (!this.state.activeFilterFieldName || !this._filterAnchorEl || !document.body.contains(this._filterAnchorEl)) {
            return;
        }
        var nextStyle = this.getFilterPopoverStyle(this._filterAnchorEl);
        var currentStyle = this.state.filterPopoverStyle || {};
        if (nextStyle.left !== currentStyle.left || nextStyle.top !== currentStyle.top
            || nextStyle.bottom !== currentStyle.bottom || nextStyle.maxHeight !== currentStyle.maxHeight) {
            this.setState({ filterPopoverStyle: nextStyle });
        }
    };
    ListControl.prototype.closeFilter = function () {
        this._filterAnchorEl = undefined;
        this.setState({ activeFilterFieldName: '', datePickerTarget: '' });
    };
    ListControl.prototype.getIsoDate = function (date) {
        var month = String(date.getMonth() + 1);
        var day = String(date.getDate());
        return String(date.getFullYear()) + '-' + (month.length < 2 ? '0' + month : month) + '-' + (day.length < 2 ? '0' + day : day);
    };
    ListControl.prototype.toggleDatePicker = function (target) {
        if (this.state.datePickerTarget === target) {
            this.setState({ datePickerTarget: '' });
            return;
        }
        var value = target === 'start' ? this.state.draftFilterValue : this.state.draftFilterEndValue;
        var month = /^\d{4}-\d{2}-\d{2}$/.test(value) ? value.substr(0, 7) : this.getIsoDate(new Date()).substr(0, 7);
        this.setState({ datePickerTarget: target, datePickerMonth: month });
    };
    ListControl.prototype.changeDatePickerMonth = function (offset) {
        var parts = this.state.datePickerMonth.split('-');
        var monthDate = new Date(Number(parts[0]), Number(parts[1]) - 1 + offset, 1);
        this.setState({ datePickerMonth: this.getIsoDate(monthDate).substr(0, 7) });
    };
    ListControl.prototype.selectFilterDate = function (value) {
        if (this.state.datePickerTarget === 'end') {
            this.setState({ draftFilterEndValue: value, datePickerTarget: '' });
        }
        else {
            this.setState({ draftFilterValue: value, datePickerTarget: '' });
        }
    };
    ListControl.prototype.renderDatePicker = function () {
        var _this = this;
        if (!this.state.datePickerTarget || !/^\d{4}-\d{2}$/.test(this.state.datePickerMonth)) {
            return null;
        }
        var parts = this.state.datePickerMonth.split('-');
        var year = Number(parts[0]);
        var monthIndex = Number(parts[1]) - 1;
        var firstWeekday = new Date(year, monthIndex, 1).getDay();
        var dayCount = new Date(year, monthIndex + 1, 0).getDate();
        var selectedValue = this.state.datePickerTarget === 'start' ? this.state.draftFilterValue : this.state.draftFilterEndValue;
        var cells = [];
        var index;
        for (index = 0; index < firstWeekday; index += 1) {
            cells.push(React.createElement("span", { key: 'blank-' + index, className: "lc-date-picker-blank" }));
        }
        var _loop_1 = function () {
            var dateValue = this_1.getIsoDate(new Date(year, monthIndex, index));
            var disabled = this_1.state.datePickerTarget === 'end' && !!this_1.state.draftFilterValue && dateValue < this_1.state.draftFilterValue;
            cells.push(React.createElement("button", { key: dateValue, type: "button", className: dateValue === selectedValue ? 'lc-date-picker-day lc-date-picker-selected' : 'lc-date-picker-day', disabled: disabled, onClick: function () { return _this.selectFilterDate(dateValue); } }, index));
        };
        var this_1 = this;
        for (index = 1; index <= dayCount; index += 1) {
            _loop_1();
        }
        return (React.createElement("div", { className: "lc-date-picker", role: "dialog", "aria-label": "Choose date" },
            React.createElement("div", { className: "lc-date-picker-header" },
                React.createElement("button", { type: "button", className: "gc-compact-icon-button", title: "Previous month", "aria-label": "Previous month", onClick: function () { return _this.changeDatePickerMonth(-1); } },
                    React.createElement("i", { className: "ms-Icon ms-Icon--ChevronLeft", "aria-hidden": "true" })),
                React.createElement("strong", null, new Date(year, monthIndex, 1).toLocaleString(undefined, { month: 'long', year: 'numeric' })),
                React.createElement("button", { type: "button", className: "gc-compact-icon-button", title: "Next month", "aria-label": "Next month", onClick: function () { return _this.changeDatePickerMonth(1); } },
                    React.createElement("i", { className: "ms-Icon ms-Icon--ChevronRight", "aria-hidden": "true" }))),
            React.createElement("div", { className: "lc-date-picker-weekdays" },
                React.createElement("span", null, "Su"),
                React.createElement("span", null, "Mo"),
                React.createElement("span", null, "Tu"),
                React.createElement("span", null, "We"),
                React.createElement("span", null, "Th"),
                React.createElement("span", null, "Fr"),
                React.createElement("span", null, "Sa")),
            React.createElement("div", { className: "lc-date-picker-days" }, cells)));
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
        var endValue = String(this.state.draftFilterEndValue || '').trim();
        if (!value) {
            delete nextFilters[fieldKey];
        }
        else if (this.state.activeFilterIsDate) {
            if (endValue && endValue < value) {
                var originalValue = value;
                value = endValue;
                endValue = originalValue;
            }
            nextFilters[fieldKey] = {
                operator: 'eq',
                value: value,
                endValue: endValue,
                compareDateOnly: true
            };
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
            draftFilterEndValue: '',
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
        var _this = this;
        if (Array.isArray(filter.value)) {
            var values = filter.value.filter(function (value) { return value !== undefined && value !== null && String(value).trim() !== ''; });
            if (values.length === 0) {
                return true;
            }
            var isNegative = filter.operator === 'ne' || filter.operator === 'notcontains';
            var scalarOperator = filter.operator === 'ne' ? 'eq' : (filter.operator === 'notcontains' ? 'contains' : filter.operator);
            var scalarMatches = values.map(function (value) { return _this.rowMatchesFilter(row, field, { operator: scalarOperator, value: value, compareDateOnly: filter.compareDateOnly }); });
            return isNegative ? scalarMatches.every(function (matches) { return !matches; }) : scalarMatches.some(function (matches) { return matches; });
        }
        var valueText = this.getFilterCellText(row, field);
        var candidate = String(valueText || '').trim();
        var query = String(filter.value || '').trim();
        if (!query) {
            return true;
        }
        var normalizedCandidate = candidate.toLowerCase();
        var normalizedQuery = query.toLowerCase();
        var compareResult = this.compareComparableValues(candidate, query);
        var fieldType = String(field.TypeAsString || '').toLowerCase();
        var lookupCandidates = fieldType.indexOf('lookup') >= 0
            ? normalizedCandidate.split(';').map(function (value) { return value.trim(); }).filter(function (value) { return !!value; })
            : [];
        if (filter.compareDateOnly) {
            var candidateDate = new Date(candidate);
            if (isNaN(candidateDate.getTime())) {
                return false;
            }
            normalizedCandidate = formatLocalDate(candidateDate).toLowerCase();
            compareResult = this.compareComparableValues(normalizedCandidate, query);
            if (filter.endValue) {
                return normalizedCandidate >= normalizedQuery && normalizedCandidate <= String(filter.endValue).toLowerCase();
            }
        }
        switch (filter.operator) {
            case 'eq':
                return lookupCandidates.length > 0 ? lookupCandidates.indexOf(normalizedQuery) >= 0 : normalizedCandidate === normalizedQuery;
            case 'ne':
                return lookupCandidates.length > 0 ? lookupCandidates.indexOf(normalizedQuery) < 0 : normalizedCandidate !== normalizedQuery;
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
    ListControl.prototype.getFilterCellText = function (row, field) {
        var fieldType = String(field.TypeAsString || '').toLowerCase();
        if (fieldType.indexOf('lookup') < 0) {
            return this.getCellPlainText(row, field);
        }
        var rawValue = this.getRowFieldValue(row, field);
        var lookupIds = [];
        var collectLookupIds = function (value) {
            if (value === undefined || value === null || value === '') {
                return;
            }
            if (Array.isArray(value)) {
                for (var valueIndex = 0; valueIndex < value.length; valueIndex += 1) {
                    collectLookupIds(value[valueIndex]);
                }
                return;
            }
            if (typeof value === 'object') {
                var lookupId = value.LookupId !== undefined ? value.LookupId
                    : value.lookupId !== undefined ? value.lookupId
                        : value.Id !== undefined ? value.Id
                            : value.ID;
                if (lookupId !== undefined && lookupId !== null && String(lookupId).trim()) {
                    lookupIds.push(String(lookupId).trim());
                }
            }
        };
        collectLookupIds(rawValue);
        if (lookupIds.length === 0) {
            var lookupFieldNames = [String(field.RealFieldName || ''), String(field.Name || '')];
            for (var fieldIndex = 0; fieldIndex < lookupFieldNames.length; fieldIndex += 1) {
                var fieldName = lookupFieldNames[fieldIndex];
                if (!fieldName) {
                    continue;
                }
                var companionValues = [row[fieldName + 'Id'], row[fieldName + '.lookupId']];
                for (var companionIndex = 0; companionIndex < companionValues.length; companionIndex += 1) {
                    var companionValue = companionValues[companionIndex];
                    if (Array.isArray(companionValue)) {
                        for (var lookupIndex = 0; lookupIndex < companionValue.length; lookupIndex += 1) {
                            lookupIds.push(String(companionValue[lookupIndex]).trim());
                        }
                    }
                    else if (companionValue !== undefined && companionValue !== null && String(companionValue).trim()) {
                        lookupIds.push(String(companionValue).trim());
                    }
                }
            }
        }
        return lookupIds.length > 0 ? lookupIds.join('; ') : this.getCellPlainText(row, field);
    };
    ListControl.prototype.parsePresetFilterConditions = function () {
        var sources = [String(this.props.filterJson || '').trim(), String(this.state.runtimeFilterJson || '').trim()];
        var conditions = [];
        for (var sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
            var source = sources[sourceIndex];
            if (!source) {
                continue;
            }
            try {
                var parsed = JSON.parse(source);
                if (!Array.isArray(parsed)) {
                    continue;
                }
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
                        valueType: String(item.valueType || '').toLowerCase() === 'expression' ? 'expression' : (String(item.valueType || '').toLowerCase() === 'fieldvalue' ? 'fieldValue' : 'static'),
                        value: item.value
                    });
                }
            }
            catch (_parseError) {
                this.logDiagnostic('Preset filter JSON is invalid; skipping that preset filter source.');
            }
        }
        return conditions;
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
        var rawValue = Array.isArray(condition.value) ? condition.value : (condition.value === undefined || condition.value === null ? '' : String(condition.value).trim());
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
    ListControl.prototype.getGroupingConfig = function () {
        var json = String(this.props.groupingJson || '').trim();
        if (!json) {
            return { enabled: false, field1: '', field2: '', collapsedByDefault: false, showCount: true };
        }
        try {
            var grouping = JSON.parse(json);
            var field1 = String(grouping && grouping.field1 || '');
            return {
                enabled: !!(grouping && grouping.enabled === true && field1),
                field1: field1,
                field2: grouping ? String(grouping.field2 || '') : '',
                collapsedByDefault: !!(grouping && grouping.collapsedByDefault === true),
                showCount: !grouping || grouping.showCount !== false
            };
        }
        catch (_error) {
            return { enabled: false, field1: '', field2: '', collapsedByDefault: false, showCount: true };
        }
    };
    ListControl.prototype.resolveFieldDefinitionByName = function (fieldName) {
        var normalized = String(fieldName || '').toLowerCase();
        for (var i = 0; i < this.state.fields.length; i += 1) {
            var field = this.state.fields[i];
            if (String(field.RealFieldName || field.Name || '').toLowerCase() === normalized
                || String(field.Name || '').toLowerCase() === normalized) {
                return field;
            }
        }
        return { Name: fieldName, DisplayName: fieldName };
    };
    ListControl.prototype.getGroupValueText = function (row, fieldName) {
        var field = this.resolveFieldDefinitionByName(fieldName);
        var text = this.getCellPlainText(row, field).trim();
        return text || strings.RuntimeGroupNoneValue;
    };
    ListControl.prototype.buildRowGroups = function (rows, grouping) {
        var groupsLevel1 = [];
        var indexByKey1 = {};
        for (var i = 0; i < rows.length; i += 1) {
            var row = rows[i];
            var value1 = this.getGroupValueText(row, grouping.field1);
            var key1 = value1.toLowerCase();
            var groupIndex1 = indexByKey1[key1];
            if (groupIndex1 === undefined) {
                groupIndex1 = groupsLevel1.length;
                indexByKey1[key1] = groupIndex1;
                groupsLevel1.push({ key: key1, label: value1, rows: [], subGroups: [] });
            }
            var group1 = groupsLevel1[groupIndex1];
            if (!grouping.field2) {
                group1.rows.push(row);
                continue;
            }
            var value2 = this.getGroupValueText(row, grouping.field2);
            var key2 = value2.toLowerCase();
            // Must explicitly reset to undefined each iteration: a bare "var subGroup;" without an
            // initializer is a no-op on repeat visits and would keep the previous row's match/creation.
            var subGroup = undefined;
            for (var s = 0; s < group1.subGroups.length; s += 1) {
                if (group1.subGroups[s].key === key2) {
                    subGroup = group1.subGroups[s];
                    break;
                }
            }
            if (!subGroup) {
                subGroup = { key: key2, label: value2, rows: [], subGroups: [] };
                group1.subGroups.push(subGroup);
            }
            subGroup.rows.push(row);
        }
        return groupsLevel1;
    };
    ListControl.prototype.countGroupRows = function (group) {
        if (group.subGroups.length === 0) {
            return group.rows.length;
        }
        var total = 0;
        for (var i = 0; i < group.subGroups.length; i += 1) {
            total += group.subGroups[i].rows.length;
        }
        return total;
    };
    ListControl.prototype.isGroupCollapsed = function (groupKey, collapsedByDefault) {
        var override = this.state.collapsedGroupKeys[groupKey];
        return override !== undefined ? override : collapsedByDefault;
    };
    ListControl.prototype.toggleGroupCollapsed = function (groupKey, collapsedByDefault) {
        var next = Object.assign({}, this.state.collapsedGroupKeys);
        next[groupKey] = !this.isGroupCollapsed(groupKey, collapsedByDefault);
        this.setState({ collapsedGroupKeys: next });
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
                    borderColor: toCssString(item.borderColor || (item.style && item.style.borderColor) || ''),
                    borderStyle: toCssString(item.borderStyle || (item.style && item.style.borderStyle) || ''),
                    borderWidth: toCssString(item.borderWidth !== undefined ? item.borderWidth : (item.style && item.style.borderWidth)),
                    borderRadius: toCssString(item.borderRadius !== undefined ? item.borderRadius : (item.style && item.style.borderRadius)),
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
        var borderColor = toCssString(styleDefinition.borderColor);
        var borderStyle = toCssString(styleDefinition.borderStyle);
        var borderWidth = toCssString(styleDefinition.borderWidth);
        var borderRadius = toCssString(styleDefinition.borderRadius);
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
        if (borderColor) {
            style.borderColor = borderColor;
        }
        if (borderStyle) {
            style.borderStyle = borderStyle;
        }
        if (borderWidth) {
            style.borderWidth = borderWidth + (isNaN(Number(borderWidth)) ? '' : 'px');
        }
        if (borderRadius) {
            style.borderRadius = borderRadius + (isNaN(Number(borderRadius)) ? '' : 'px');
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
    ListControl.prototype.getCommandButtonClass = function (className) {
        return joinClassNames([
            className || '',
            'gc-command-button',
            this.props.buttonDisplayMode === 'icon' ? 'gc-command-button-icon-only' : ''
        ]);
    };
    ListControl.prototype.renderCommandContent = function (iconName, label) {
        if (this.props.buttonDisplayMode !== 'icon' && this.props.buttonDisplayMode !== 'iconText') {
            return label;
        }
        return (React.createElement("span", { className: "gc-command-content" },
            React.createElement("i", { className: 'ms-Icon ms-Icon--' + iconName, "aria-hidden": "true" }),
            React.createElement("span", { className: this.props.buttonDisplayMode === 'icon' ? 'gc-visually-hidden' : '' }, label)));
    };
    ListControl.prototype.renderDataRow = function (row, keySuffix, displayFields, fieldsByKey, conditionalRules) {
        var _this = this;
        var rowItemId = this.getRowItemId(row);
        var isSelected = rowItemId > 0 && rowItemId === this.state.selectedItemId;
        var conditionalStyle = this.getConditionalStyleForRow(row, fieldsByKey, conditionalRules);
        return (React.createElement("tr", { key: rowItemId > 0 ? String(rowItemId) : keySuffix, onClick: function () { return _this.selectRow(row); }, className: joinClassNames(['lc-row', isSelected ? 'lc-row-selected' : '']) }, displayFields.map(function (field) {
            var markup = _this.getCellMarkup(row, field);
            var urlCell = _this.getUrlCellValue(row, field);
            var showItemLink = _this.props.showLinkToItem && _this.isTitleField(field);
            var itemLinkUrl = showItemLink ? _this.getItemLinkUrl(row) : '';
            var linkOpenBehavior = String(_this.props.linkOpenBehavior || 'self');
            var itemLinkText = showItemLink ? _this.getCellPlainText(row, field) : '';
            var cellFieldKey = _this.getFieldKey(field);
            var columnStyle = conditionalStyle.columnStylesByFieldKey[cellFieldKey] || {};
            var mergedCellStyle = mergeStyleObjects(mergeStyleObjects(conditionalStyle.rowStyle, columnStyle), _this.getConfiguredColumnStyle(field));
            return (React.createElement("td", { key: field.Name, style: mergedCellStyle }, urlCell ? (React.createElement("a", { className: "lc-item-link", href: urlCell.href, target: "_blank", rel: "noopener noreferrer", onClick: function (ev) { return ev.stopPropagation(); } }, urlCell.text)) : showItemLink && itemLinkUrl ? (React.createElement("a", { className: "lc-item-link", href: itemLinkUrl, target: linkOpenBehavior === 'newTab' ? '_blank' : undefined, rel: linkOpenBehavior === 'newTab' ? 'noopener noreferrer' : undefined, onClick: function (ev) {
                    ev.stopPropagation();
                    if (linkOpenBehavior === 'dialog') {
                        ev.preventDefault();
                        _this.openItemLinkDialog(itemLinkUrl);
                    }
                    else if (linkOpenBehavior === 'newWindow') {
                        ev.preventDefault();
                        _this.openItemLinkWindow(itemLinkUrl);
                    }
                } }, itemLinkText || strings.RuntimeView)) : (markup ? React.createElement("span", { dangerouslySetInnerHTML: markup }) : null)));
        })));
    };
    ListControl.prototype.renderGroupHeaderRow = function (group, keyPath, level, totalColumnCount, groupingConfig) {
        var _this = this;
        var collapsed = this.isGroupCollapsed(keyPath, groupingConfig.collapsedByDefault);
        var count = this.countGroupRows(group);
        var toggleLabel = collapsed ? strings.RuntimeGroupExpand : strings.RuntimeGroupCollapse;
        return (React.createElement("tr", { key: 'group-' + keyPath, className: joinClassNames(['gc-group-row', 'gc-group-row-level-' + String(level)]) },
            React.createElement("td", { colSpan: totalColumnCount, style: { paddingLeft: (10 + level * 20) + 'px' } },
                React.createElement("button", { type: "button", className: this.getCommandButtonClass('gc-group-toggle'), title: toggleLabel, "aria-label": toggleLabel, onClick: function () { return _this.toggleGroupCollapsed(keyPath, groupingConfig.collapsedByDefault); } }, this.renderCommandContent(collapsed ? 'ChevronRight' : 'ChevronDown', toggleLabel)),
                React.createElement("span", { className: "gc-group-label" }, group.label),
                groupingConfig.showCount && React.createElement("span", { className: "gc-group-count" }, formatString(strings.RuntimeGroupItemCount, count)))));
    };
    ListControl.prototype.renderRowGroupRows = function (group, keyPath, level, groupingConfig, totalColumnCount, displayFields, fieldsByKey, conditionalRules) {
        var rowsOutput = [this.renderGroupHeaderRow(group, keyPath, level, totalColumnCount, groupingConfig)];
        if (this.isGroupCollapsed(keyPath, groupingConfig.collapsedByDefault)) {
            return rowsOutput;
        }
        if (group.subGroups.length > 0) {
            for (var i = 0; i < group.subGroups.length; i += 1) {
                var subGroup = group.subGroups[i];
                var subKeyPath = keyPath + '/' + subGroup.key;
                rowsOutput = rowsOutput.concat(this.renderRowGroupRows(subGroup, subKeyPath, level + 1, groupingConfig, totalColumnCount, displayFields, fieldsByKey, conditionalRules));
            }
        }
        else {
            for (var r = 0; r < group.rows.length; r += 1) {
                rowsOutput.push(this.renderDataRow(group.rows[r], keyPath + '-' + String(r), displayFields, fieldsByKey, conditionalRules));
            }
        }
        return rowsOutput;
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
        var groupingConfig = this.getGroupingConfig();
        var rowGroups = groupingConfig.enabled ? this.buildRowGroups(visibleRows, groupingConfig) : [];
        var totalColumnCount = displayFields.length;
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
            '--lc-header-bg-color': this.props.headerBackgroundColor || '#ffffff',
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
            (this.state.displayFormUrl || this.state.displayFormLoading || this.state.displayFormError)
                && React.createElement("div", { className: "lc-form-dialog-backdrop", role: "presentation" },
                    React.createElement("section", { className: "lc-form-dialog", role: "dialog", "aria-modal": "true", "aria-label": "Item details", onClick: function (event) { return event.stopPropagation(); } },
                        React.createElement("button", { type: "button", className: "lc-form-dialog-close", "aria-label": "Close item details", title: "Close item details", onClick: function () { return _this.closeDefaultDisplayForm(); } }, "\u00D7"),
                        this.state.displayFormLoading && React.createElement("div", { className: "lc-form-dialog-message" }, "Loading item..."),
                        this.state.displayFormError && React.createElement("div", { className: "lc-form-dialog-error" }, this.state.displayFormError),
                        this.state.displayFormUrl && React.createElement("iframe", { ref: this.setDisplayFormFrameRef, className: "lc-form-dialog-frame", src: this.state.displayFormUrl, title: "Item details" }))),
            React.createElement("div", { className: "lc-toolbar" },
                this.props.showViewSelector && !this.state.embeddedByReportForms && (React.createElement("label", null,
                    React.createElement("span", { style: { marginRight: '6px' } }, strings.RuntimeViewLabel),
                    this.props.showViewAsDropdown !== false ? (React.createElement("select", { value: this.state.selectedViewId, onChange: function (ev) { return _this.setState({ selectedViewId: ev.currentTarget.value }); } }, this.props.views.map(function (view) {
                        return React.createElement("option", { key: view.key, value: view.key }, view.text);
                    }))) : (React.createElement("span", { className: "gc-view-static-label" }, this.getSelectedViewLabel())))),
                this.props.showRefresh && !this.state.embeddedByReportForms && React.createElement("button", { type: "button", className: this.getCommandButtonClass(), title: strings.RuntimeRefresh, "aria-label": strings.RuntimeRefresh, onClick: function () { return _this.loadRows(); } }, this.renderCommandContent('Refresh', strings.RuntimeRefresh)),
                this.props.showAdd && !this.state.embeddedByReportForms && (React.createElement("button", { type: "button", className: this.getCommandButtonClass(), title: strings.RuntimeNew, "aria-label": strings.RuntimeNew, onClick: function () {
                        _this.setState({ selectedItemId: 0, selectedMode: 'new' });
                        _this.props.onSelectionChange(0, 'new');
                    } }, this.renderCommandContent('Add', strings.RuntimeNew))),
                this.props.showEdit && !this.state.embeddedByReportForms && (React.createElement("button", { type: "button", className: this.getCommandButtonClass(), disabled: !canOperateOnSelection, title: strings.RuntimeEdit, "aria-label": strings.RuntimeEdit, onClick: function () {
                        _this.setState({ selectedMode: 'edit' });
                        _this.props.onSelectionChange(_this.state.selectedItemId, 'edit');
                    } }, this.renderCommandContent('Edit', strings.RuntimeEdit))),
                this.props.showView && !this.state.embeddedByReportForms && (React.createElement("button", { type: "button", className: this.getCommandButtonClass(), disabled: !canOperateOnSelection, title: strings.RuntimeView, "aria-label": strings.RuntimeView, onClick: function () {
                        _this.setState({ selectedMode: 'view' });
                        _this.props.onSelectionChange(_this.state.selectedItemId, 'view');
                    } }, this.renderCommandContent('View', strings.RuntimeView))),
                this.props.showDelete && !this.state.embeddedByReportForms && (React.createElement("button", { type: "button", className: this.getCommandButtonClass(), disabled: !canOperateOnSelection || this.state.deleting, title: strings.RuntimeDelete, "aria-label": strings.RuntimeDelete, onClick: function () { return _this.deleteSelected(); } }, this.renderCommandContent('Delete', strings.RuntimeDelete)))),
            this.props.isEditMode && (React.createElement("div", { className: "lc-status" },
                React.createElement("div", null, formatString(strings.RuntimeSelectedItem, this.state.selectedItemId || 0)),
                React.createElement("div", null, formatString(strings.RuntimeSelectedMode, this.state.selectedMode)),
                React.createElement("div", null, formatString(strings.RuntimeConditionalStyleMatches, selectedMatchedRuleNames.length > 0 ? selectedMatchedRuleNames.join(', ') : strings.RuntimeConditionalStyleMatchesNone)),
                !canOperateOnSelection && React.createElement("div", null, strings.RuntimeSelectRowPrompt))),
            this.state.loading && React.createElement("div", null, strings.RuntimeLoading),
            !!this.state.error && React.createElement("div", null, this.state.error),
            !this.state.loading && !this.state.error && processedRows.length === 0 && (React.createElement("div", null, hasActiveFilters ? strings.RuntimeNoItemsAfterFilter : strings.RuntimeNoItems)),
            !this.state.loading && !this.state.error && processedRows.length > 0 && (React.createElement("div", { className: "lc-table-container" },
                React.createElement("div", { className: "lc-sticky-header-viewport", ref: this._setStickyHeaderViewportRef }),
                React.createElement("div", { className: "lc-table-wrap", ref: this._setTableWrapRef },
                    this.state.showScrollArrows && (React.createElement("button", { type: "button", className: "lc-scroll-arrow lc-scroll-arrow-left", style: { top: this.state.scrollArrowTop + 'px', left: (this.state.scrollArrowLeft + 6) + 'px' }, title: strings.RuntimeScrollLeft, "aria-label": strings.RuntimeScrollLeft, onClick: function () { return _this.scrollTableHorizontally(-1); } }, "\u2039")),
                    this.state.showScrollArrows && (React.createElement("button", { type: "button", className: "lc-scroll-arrow lc-scroll-arrow-right", style: { top: this.state.scrollArrowTop + 'px', right: (this.state.scrollArrowRight + 6) + 'px' }, title: strings.RuntimeScrollRight, "aria-label": strings.RuntimeScrollRight, onClick: function () { return _this.scrollTableHorizontally(1); } }, "\u203A")),
                    React.createElement("table", { className: "lc-table" },
                        React.createElement("thead", { ref: this._setTableHeadRef },
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
                                            React.createElement("button", { type: "button", className: _this.getCommandButtonClass(joinClassNames(['lc-header-action', filterActive ? 'lc-header-action-active' : ''])), title: strings.RuntimeFilterTitle, "aria-label": strings.RuntimeFilterTitle, onClick: function (ev) {
                                                    ev.preventDefault();
                                                    ev.stopPropagation();
                                                    if (_this.state.activeFilterFieldName === fieldKey) {
                                                        _this.closeFilter();
                                                    }
                                                    else {
                                                        _this.openFilter(field, ev.currentTarget);
                                                    }
                                                } }, _this.renderCommandContent('Filter', strings.RuntimeFilterIcon))),
                                        _this.state.activeFilterFieldName === fieldKey && (React.createElement("div", { className: "lc-filter-popover", style: _this.state.filterPopoverStyle, onClick: function (ev) {
                                                ev.preventDefault();
                                                ev.stopPropagation();
                                            } },
                                            _this.state.activeFilterIsDate ? (React.createElement("div", null,
                                                React.createElement("label", { className: "lc-filter-label" }, strings.RuntimeFilterDateLabel),
                                                React.createElement("div", { className: "lc-date-input-row" },
                                                    React.createElement("input", { className: "lc-filter-input", type: "text", placeholder: "YYYY-MM-DD", value: _this.state.draftFilterValue, onChange: function (ev) { return _this.setState({ draftFilterValue: ev.currentTarget.value }); } }),
                                                    React.createElement("button", { type: "button", className: "lc-date-picker-button gc-compact-icon-button", title: "Choose from date", "aria-label": "Choose from date", onClick: function () { return _this.toggleDatePicker('start'); } },
                                                        React.createElement("i", { className: "ms-Icon ms-Icon--Calendar", "aria-hidden": "true" }))),
                                                React.createElement("label", { className: "lc-filter-label" }, strings.RuntimeFilterEndDateLabel),
                                                React.createElement("div", { className: "lc-date-input-row" },
                                                    React.createElement("input", { className: "lc-filter-input", type: "text", placeholder: "YYYY-MM-DD", value: _this.state.draftFilterEndValue, onChange: function (ev) { return _this.setState({ draftFilterEndValue: ev.currentTarget.value }); } }),
                                                    React.createElement("button", { type: "button", className: "lc-date-picker-button gc-compact-icon-button", title: "Choose to date", "aria-label": "Choose to date", onClick: function () { return _this.toggleDatePicker('end'); } },
                                                        React.createElement("i", { className: "ms-Icon ms-Icon--Calendar", "aria-hidden": "true" }))),
                                                _this.renderDatePicker())) : (React.createElement("div", null,
                                                React.createElement("label", { className: "lc-filter-label" }, strings.RuntimeFilterOperatorLabel),
                                                React.createElement("select", { className: "lc-filter-select", value: _this.state.draftFilterOperator, onChange: function (ev) { return _this.setState({ draftFilterOperator: ev.currentTarget.value }); } }, _this.getFilterOperatorOptions().map(function (option) {
                                                    return React.createElement("option", { key: option.key, value: option.key }, option.label);
                                                })),
                                                React.createElement("label", { className: "lc-filter-label" }, strings.RuntimeFilterValueLabel),
                                                React.createElement("input", { className: "lc-filter-input", type: "text", value: _this.state.draftFilterValue, placeholder: strings.RuntimeFilterValuePlaceholder, onChange: function (ev) { return _this.setState({ draftFilterValue: ev.currentTarget.value }); }, onKeyDown: function (ev) {
                                                        if (ev.key === 'Enter') {
                                                            _this.applyActiveFilter();
                                                        }
                                                    } }))),
                                            React.createElement("div", { className: "lc-filter-actions" },
                                                React.createElement("button", { type: "button", className: _this.getCommandButtonClass(), title: strings.RuntimeFilterApply, "aria-label": strings.RuntimeFilterApply, onClick: function () { return _this.applyActiveFilter(); } }, _this.renderCommandContent('CheckMark', strings.RuntimeFilterApply)),
                                                React.createElement("button", { type: "button", className: _this.getCommandButtonClass(), title: strings.RuntimeFilterClear, "aria-label": strings.RuntimeFilterClear, onClick: function () { return _this.clearActiveFilter(); } }, _this.renderCommandContent('ClearFilter', strings.RuntimeFilterClear)),
                                                React.createElement("button", { type: "button", className: _this.getCommandButtonClass(), title: strings.RuntimeFilterClose, "aria-label": strings.RuntimeFilterClose, onClick: function () { return _this.closeFilter(); } }, _this.renderCommandContent('Cancel', strings.RuntimeFilterClose))))))));
                            }))),
                        React.createElement("tbody", null, groupingConfig.enabled
                            ? rowGroups.reduce(function (accumulated, group) { return accumulated.concat(_this.renderRowGroupRows(group, group.key, 0, groupingConfig, totalColumnCount, displayFields, fieldsByKey, conditionalRules)); }, [])
                            : visibleRows.map(function (row, index) { return _this.renderDataRow(row, String(index), displayFields, fieldsByKey, conditionalRules); })))))),
            !this.state.loading && !this.state.error && (React.createElement("div", { className: "lc-footer" },
                React.createElement("div", { className: "lc-total-items" }, this.state.nextPageHref
                    ? strings.RuntimeLoadedItemsLabel.replace('{0}', String(processedRows.length))
                    : formatString(strings.RuntimeTotalItemsLabel, processedRows.length)),
                pageSize > 0 && (pageCount > 1 || !!this.state.nextPageHref) && (React.createElement("div", { className: "lc-pagination" },
                    React.createElement("button", { type: "button", className: this.getCommandButtonClass(), disabled: currentPage === 0, title: strings.RuntimePreviousPage, "aria-label": strings.RuntimePreviousPage, onClick: function () { return _this.setState({ currentPage: Math.max(0, currentPage - 1) }); } }, this.renderCommandContent('ChevronLeft', strings.RuntimePreviousPage)),
                    React.createElement("span", null, (this.state.nextPageHref ? strings.RuntimePageStatusMore : strings.RuntimePageStatus).replace('{0}', String(currentPage + 1)).replace('{1}', String(pageCount)).replace('{2}', String(pageSize))),
                    React.createElement("button", { type: "button", className: this.getCommandButtonClass(), disabled: this.state.loadingMore || (currentPage >= pageCount - 1 && !this.state.nextPageHref), title: strings.RuntimeNextPage, "aria-label": strings.RuntimeNextPage, onClick: function () { return _this.goToNextPage(currentPage, pageCount); } }, this.renderCommandContent('ChevronRight', strings.RuntimeNextPage)),
                    this.state.loadingMore && React.createElement("span", null, strings.RuntimeLoadingMore)))))));
    };
    return ListControl;
}(React.Component));
exports.ListControl = ListControl;

//# sourceMappingURL=ListControl.js.map
