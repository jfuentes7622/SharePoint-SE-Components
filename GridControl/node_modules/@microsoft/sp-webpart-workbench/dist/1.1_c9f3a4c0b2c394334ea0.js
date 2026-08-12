webpackJsonp8be81a5c_af38_4bb2_af97_afa3b64dfbed_1_5_1([1],{

/***/ 297:
/***/ (function(module, exports, __webpack_require__) {

var content = __webpack_require__(298);
var loader = __webpack_require__(3);

if(typeof content === "string") content = [[module.i, content]];

// add the styles to the DOM
for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1], true);

if(content.locals) module.exports = content.locals;

/***/ }),

/***/ 298:
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__(4)(false);
// imports


// module
exports.push([module.i, ".controlsInOneRow_b475b251{display:-ms-flexbox;display:flex}.controlsInOneRow_b475b251.spaceInBetween_b475b251{-ms-flex-pack:justify;justify-content:space-between}", ""]);

// exports


/***/ }),

/***/ 299:
/***/ (function(module, exports, __webpack_require__) {

var content = __webpack_require__(300);
var loader = __webpack_require__(3);

if(typeof content === "string") content = [[module.i, content]];

// add the styles to the DOM
for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1], true);

if(content.locals) module.exports = content.locals;

/***/ }),

/***/ 300:
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__(4)(false);
// imports


// module
exports.push([module.i, ".pickerCallout_4733da1f{width:166px}.pickerLabel_4733da1f{display:block;margin-left:6px}.pickerLabel_4733da1f.isRTL_4733da1f{margin-right:6px}[dir=ltr] .buttonLabel_4733da1f{padding-left:10px}[dir=rtl] .buttonLabel_4733da1f{padding-right:10px}.previewSvg_4733da1f{width:20px;height:20px}.defaultSvg_4733da1f,.previewSvg_4733da1f.border_4733da1f{border:1px solid #000}.defaultSvg_4733da1f{padding:2px;fill:#fff}.defaultSvg_4733da1f.fillThemeColor_4733da1f{fill:#333;border:none}.defaultSvg_4733da1f.selected_4733da1f{width:12px;height:12px;border:4px solid #c8c8c8}.colorPickerButton_4733da1f{background:0 0;padding-left:4px;padding-right:12px}.previewRectangle_4733da1f{width:100%;height:100%}.previewIcon_4733da1f{font-size:7px;padding-top:8px}[dir=ltr] .previewIcon_4733da1f{padding-left:7px}[dir=rtl] .previewIcon_4733da1f{padding-right:7px}", ""]);

// exports


/***/ }),

/***/ 301:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/SwatchColorPicker/SwatchColorPicker.js
module.exports = __webpack_require__(302);

/***/ }),

/***/ 302:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SwatchColorPicker", function() { return SwatchColorPicker; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__SwatchColorPicker_base__ = __webpack_require__(303);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__SwatchColorPicker_base___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1__SwatchColorPicker_base__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__SwatchColorPicker_styles__ = __webpack_require__(327);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__SwatchColorPicker_styles___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__SwatchColorPicker_styles__);



var SwatchColorPicker = Object(__WEBPACK_IMPORTED_MODULE_0__Utilities__["styled"])(__WEBPACK_IMPORTED_MODULE_1__SwatchColorPicker_base__["SwatchColorPickerBase"], __WEBPACK_IMPORTED_MODULE_2__SwatchColorPicker_styles__["getStyles"]);


/***/ }),

/***/ 303:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/SwatchColorPicker/SwatchColorPicker.base.js
module.exports = __webpack_require__(304);

/***/ }),

/***/ 304:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "SwatchColorPickerBase", function() { return SwatchColorPickerBase; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_tslib__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react__ = __webpack_require__(2);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_react__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__utilities_grid_Grid__ = __webpack_require__(305);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__utilities_grid_Grid___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3__utilities_grid_Grid__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__ColorPickerGridCell__ = __webpack_require__(311);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__ColorPickerGridCell___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_4__ColorPickerGridCell__);





var getClassNames = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["classNamesFunction"])();
var SwatchColorPickerBase = /** @class */ (function (_super) {
    __WEBPACK_IMPORTED_MODULE_0_tslib__["c" /* __extends */](SwatchColorPickerBase, _super);
    function SwatchColorPickerBase(props) {
        var _this = _super.call(this, props) || this;
        _this.navigationIdleDelay = 250 /* ms */;
        /**
         * When the whole swatchColorPicker is blurred,
         * make sure to clear the pending focused stated
         */
        _this._onSwatchColorPickerBlur = function () {
            if (_this.props.onCellFocused) {
                _this._cellFocused = false;
                _this.props.onCellFocused();
            }
        };
        /**
         * Render a color cell
         * @param item - The item to render
         * @returns {JSX.Element} - Element representing the item
         */
        _this._renderOption = function (item) {
            var id = _this._id;
            return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"](__WEBPACK_IMPORTED_MODULE_4__ColorPickerGridCell__["ColorPickerGridCell"], { item: item, id: id, color: item.color, getStyles: _this.props.getColorGridCellStyles, disabled: _this.props.disabled, onClick: _this._onCellClick, onHover: _this._onGridCellHovered, onFocus: _this._onGridCellFocused, selected: _this.state.selectedIndex !== undefined && (_this.state.selectedIndex === item.index), circle: _this.props.cellShape === 'circle', label: item.label, onMouseEnter: _this._onMouseEnter, onMouseMove: _this._onMouseMove, onMouseLeave: _this._onMouseLeave, onWheel: _this._onWheel, onKeyDown: _this._onKeyDown }));
        };
        /**
         * Callback passed to the GridCell that will manage triggering the onCellHovered callback for mouseEnter
         */
        _this._onMouseEnter = function (ev) {
            if (!_this.props.focusOnHover) {
                if (!_this.isNavigationIdle || _this.props.disabled) {
                    return true;
                }
                return false;
            }
            if (_this.isNavigationIdle && !_this.props.disabled) {
                ev.currentTarget.focus();
            }
            return true;
        };
        /**
         * Callback passed to the GridCell that will manage Hover/Focus updates
         */
        _this._onMouseMove = function (ev) {
            if (!_this.props.focusOnHover) {
                if (!_this.isNavigationIdle || _this.props.disabled) {
                    return true;
                }
                return false;
            }
            var targetElement = ev.currentTarget;
            // If navigation is idle and the targetElement is the focused element bail out
            // if (!this.isNavigationIdle || (document && targetElement === (document.activeElement as HTMLElement))) {
            if (_this.isNavigationIdle && !(document && targetElement === document.activeElement)) {
                targetElement.focus();
            }
            return true;
        };
        /**
         * Callback passed to the GridCell that will manage Hover/Focus updates
         */
        _this._onMouseLeave = function (ev) {
            var parentSelector = _this.props.mouseLeaveParentSelector;
            if (!_this.props.focusOnHover ||
                !parentSelector ||
                !_this.isNavigationIdle ||
                _this.props.disabled) {
                return;
            }
            // Get the the elements that math the given selector
            var elements = document.querySelectorAll(parentSelector);
            // iterate over the elements return to make sure it is a parent of the target and focus it
            for (var index = 0; index < elements.length; index += 1) {
                if (elements[index].contains(ev.currentTarget)) {
                    /**
                     * IE11 focus() method forces parents to scroll to top of element.
                     * Edge and IE expose a setActive() function for focusable divs that
                     * sets the page focus but does not scroll the parent element.
                     */
                    if (elements[index].setActive) {
                        elements[index].setActive();
                    }
                    else {
                        elements[index].focus();
                    }
                    break;
                }
            }
        };
        /**
         * Callback to make sure we don't update the hovered element during mouse wheel
         */
        _this._onWheel = function () {
            _this.setNavigationTimeout();
        };
        /**
         * Callback that
         */
        _this._onKeyDown = function (ev) {
            if (ev.which === 38 /* up */ ||
                ev.which === 40 /* down */ ||
                ev.which === 37 /* left */ ||
                ev.which === 39 /* right */) {
                _this.setNavigationTimeout();
            }
        };
        /**
         * Sets a timeout so we won't process any mouse "hover" events
         * while navigating (via mouseWheel or arrowKeys)
         */
        _this.setNavigationTimeout = function () {
            if (!_this.isNavigationIdle && _this.navigationIdleTimeoutId !== undefined) {
                _this.async.clearTimeout(_this.navigationIdleTimeoutId);
                _this.navigationIdleTimeoutId = undefined;
            }
            else {
                _this.isNavigationIdle = false;
            }
            _this.navigationIdleTimeoutId = _this.async.setTimeout(function () {
                _this.isNavigationIdle = true;
            }, _this.navigationIdleDelay);
        };
        /**
         * Callback passed to the GridCell class that will trigger the onCellHovered callback of the SwatchColorPicker
         * NOTE: This will not be triggered if shouldFocusOnHover === true
         */
        _this._onGridCellHovered = function (item) {
            var onCellHovered = _this.props.onCellHovered;
            if (onCellHovered) {
                return item ? onCellHovered(item.id, item.color) : onCellHovered();
            }
        };
        /**
         * Callback passed to the GridCell class that will trigger the onCellFocus callback of the SwatchColorPicker
         */
        _this._onGridCellFocused = function (item) {
            var onCellFocused = _this.props.onCellFocused;
            if (onCellFocused) {
                if (item) {
                    _this._cellFocused = true;
                    return onCellFocused(item.id, item.color);
                }
                else {
                    _this._cellFocused = false;
                    return onCellFocused();
                }
            }
        };
        /**
         * Handle the click on a cell
         * @param item - The cell that the click was fired against
         */
        _this._onCellClick = function (item) {
            if (_this.props.disabled) {
                return;
            }
            var index = item.index;
            // If we have a valid index and it is not already
            // selected, select it
            if (index >= 0 && index !== _this.state.selectedIndex) {
                if (_this.props.onCellFocused && _this._cellFocused) {
                    _this._cellFocused = false;
                    _this.props.onCellFocused();
                }
                if (_this.props.onColorChanged) {
                    _this.props.onColorChanged(item.id, item.color);
                }
                _this.setState({
                    selectedIndex: index
                });
            }
        };
        _this._id = props.id || Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["getId"])('swatchColorPicker');
        _this._warnMutuallyExclusive({
            'focusOnHover': 'onHover'
        });
        _this._warnConditionallyRequiredProps(['focusOnHover'], 'mouseLeaveParentSelector', !!_this.props.mouseLeaveParentSelector);
        _this.isNavigationIdle = true;
        _this.async = new __WEBPACK_IMPORTED_MODULE_2__Utilities__["Async"](_this);
        var selectedIndex;
        if (props.selectedId) {
            selectedIndex = _this._getSelectedIndex(props.colorCells, props.selectedId);
        }
        _this.state = {
            selectedIndex: selectedIndex
        };
        return _this;
    }
    SwatchColorPickerBase.prototype.componentWillReceiveProps = function (newProps) {
        var newSelectedIndex;
        if (newProps.selectedId) {
            newSelectedIndex = this._getSelectedIndex(newProps.colorCells, newProps.selectedId);
        }
        if (newSelectedIndex !== this.state.selectedIndex) {
            this.setState({
                selectedIndex: newSelectedIndex
            });
        }
    };
    SwatchColorPickerBase.prototype.componentWillUnmount = function () {
        if (this.props.onCellFocused && this._cellFocused) {
            this._cellFocused = false;
            this.props.onCellFocused();
        }
    };
    SwatchColorPickerBase.prototype.render = function () {
        var _a = this.props, colorCells = _a.colorCells, columnCount = _a.columnCount, positionInSet = _a.positionInSet, setSize = _a.setSize, shouldFocusCircularNavigate = _a.shouldFocusCircularNavigate, className = _a.className, doNotContainWithinFocusZone = _a.doNotContainWithinFocusZone, getStyles = _a.getStyles;
        var classNames = getClassNames(getStyles, {
            theme: this.props.theme,
            className: className,
        });
        if (colorCells.length < 1 || columnCount < 1) {
            return null;
        }
        return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"](__WEBPACK_IMPORTED_MODULE_3__utilities_grid_Grid__["Grid"], __WEBPACK_IMPORTED_MODULE_0_tslib__["a" /* __assign */]({}, this.props, { items: colorCells.map(function (item, index) { return __WEBPACK_IMPORTED_MODULE_0_tslib__["a" /* __assign */]({}, item, { index: index }); }), columnCount: columnCount, onRenderItem: this._renderOption, positionInSet: positionInSet && positionInSet, setSize: setSize && setSize, shouldFocusCircularNavigate: shouldFocusCircularNavigate, doNotContainWithinFocusZone: doNotContainWithinFocusZone, onBlur: this._onSwatchColorPickerBlur, theme: this.props.theme, 
            // tslint:disable-next-line:jsx-no-lambda
            getStyles: function (props) { return ({
                root: classNames.root,
                tableCell: classNames.tableCell,
                focusedContainer: classNames.focusedContainer
            }); } })));
    };
    /**
     * Get the selected item's index
     * @param items - The items to search
     * @param selectedId - The selected item's id to find
     * @returns {number} - The index of the selected item's id, -1 if there was no match
     */
    SwatchColorPickerBase.prototype._getSelectedIndex = function (items, selectedId) {
        var selectedIndex = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["findIndex"])(items, (function (item) { return (item.id === selectedId); }));
        return selectedIndex >= 0 ? selectedIndex : undefined;
    };
    SwatchColorPickerBase.defaultProps = {
        cellShape: 'circle',
        disabled: false,
        shouldFocusCircularNavigate: true
    };
    SwatchColorPickerBase = __WEBPACK_IMPORTED_MODULE_0_tslib__["b" /* __decorate */]([
        Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["customizable"])('SwatchColorPicker', ['theme'])
    ], SwatchColorPickerBase);
    return SwatchColorPickerBase;
}(__WEBPACK_IMPORTED_MODULE_2__Utilities__["BaseComponent"]));



/***/ }),

/***/ 305:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/utilities/grid/Grid.js
module.exports = __webpack_require__(306);

/***/ }),

/***/ 306:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "Grid", function() { return Grid; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Grid_base__ = __webpack_require__(307);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Grid_base___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1__Grid_base__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Grid_styles__ = __webpack_require__(309);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Grid_styles___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__Grid_styles__);



var Grid = Object(__WEBPACK_IMPORTED_MODULE_0__Utilities__["styled"])(__WEBPACK_IMPORTED_MODULE_1__Grid_base__["GridBase"], __WEBPACK_IMPORTED_MODULE_2__Grid_styles__["getStyles"]);


/***/ }),

/***/ 307:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/utilities/grid/Grid.base.js
module.exports = __webpack_require__(308);

/***/ }),

/***/ 308:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "GridBase", function() { return GridBase; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_tslib__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react__ = __webpack_require__(2);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_react__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__FocusZone__ = __webpack_require__(125);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__FocusZone___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3__FocusZone__);




var getClassNames = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["classNamesFunction"])();
var GridBase = /** @class */ (function (_super) {
    __WEBPACK_IMPORTED_MODULE_0_tslib__["c" /* __extends */](GridBase, _super);
    function GridBase(props) {
        var _this = _super.call(this, props) || this;
        _this._id = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["getId"])();
        return _this;
    }
    GridBase.prototype.render = function () {
        var _this = this;
        var _a = this.props, items = _a.items, columnCount = _a.columnCount, onRenderItem = _a.onRenderItem, positionInSet = _a.positionInSet, setSize = _a.setSize, getStyles = _a.getStyles;
        var htmlProps = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["getNativeProps"])(this.props, __WEBPACK_IMPORTED_MODULE_2__Utilities__["htmlElementProperties"], ['onBlur, aria-posinset, aria-setsize']);
        var classNames = getClassNames(getStyles, { theme: this.props.theme });
        // Array to store the cells in the correct row index
        var rowsOfItems = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["toMatrix"])(items, columnCount);
        var content = (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("table", __WEBPACK_IMPORTED_MODULE_0_tslib__["a" /* __assign */]({}, htmlProps, { "aria-posinset": positionInSet, "aria-setsize": setSize, id: this._id, role: 'grid', className: classNames.root }),
            __WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("tbody", null, rowsOfItems.map(function (rows, rowIndex) {
                return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("tr", { role: 'row', key: _this._id + '-' + rowIndex + '-row' }, rows.map(function (cell, cellIndex) {
                    return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("td", { role: 'presentation', key: _this._id + '-' + cellIndex + '-cell', className: classNames.tableCell }, onRenderItem(cell, cellIndex)));
                })));
            }))));
        // Create the table/grid
        return (this.props.doNotContainWithinFocusZone ? content : (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"](__WEBPACK_IMPORTED_MODULE_3__FocusZone__["FocusZone"], { isCircularNavigation: this.props.shouldFocusCircularNavigate, className: classNames.focusedContainer, onBlur: this.props.onBlur }, content)));
    };
    return GridBase;
}(__WEBPACK_IMPORTED_MODULE_2__Utilities__["BaseComponent"]));



/***/ }),

/***/ 309:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/utilities/grid/Grid.styles.js
module.exports = __webpack_require__(310);

/***/ }),

/***/ 310:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getStyles", function() { return getStyles; });
var getStyles = function (props) {
    return {
        root: {
            padding: 2,
            outline: 'none'
        },
        tableCell: {
            padding: 0
        }
    };
};


/***/ }),

/***/ 311:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/SwatchColorPicker/ColorPickerGridCell.js
module.exports = __webpack_require__(312);

/***/ }),

/***/ 312:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ColorPickerGridCell", function() { return ColorPickerGridCell; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__ColorPickerGridCell_base__ = __webpack_require__(313);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__ColorPickerGridCell_base___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1__ColorPickerGridCell_base__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__ColorPickerGridCell_styles__ = __webpack_require__(325);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__ColorPickerGridCell_styles___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__ColorPickerGridCell_styles__);



var ColorPickerGridCell = Object(__WEBPACK_IMPORTED_MODULE_0__Utilities__["styled"])(__WEBPACK_IMPORTED_MODULE_1__ColorPickerGridCell_base__["ColorPickerGridCellBase"], __WEBPACK_IMPORTED_MODULE_2__ColorPickerGridCell_styles__["getStyles"]);


/***/ }),

/***/ 313:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/SwatchColorPicker/ColorPickerGridCell.base.js
module.exports = __webpack_require__(314);

/***/ }),

/***/ 314:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "ColorPickerGridCellBase", function() { return ColorPickerGridCellBase; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_tslib__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react__ = __webpack_require__(2);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_react__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__utilities_color_colors__ = __webpack_require__(315);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__utilities_color_colors___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3__utilities_color_colors__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__utilities_grid_GridCell__ = __webpack_require__(319);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_4__utilities_grid_GridCell___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_4__utilities_grid_GridCell__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__Button_ActionButton_ActionButton_styles__ = __webpack_require__(321);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_5__Button_ActionButton_ActionButton_styles___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_5__Button_ActionButton_ActionButton_styles__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6__Styling__ = __webpack_require__(119);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_6__Styling___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_6__Styling__);








var getClassNames = Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["classNamesFunction"])();
var ColorCell = /** @class */ (function (_super) {
    __WEBPACK_IMPORTED_MODULE_0_tslib__["c" /* __extends */](ColorCell, _super);
    function ColorCell() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    return ColorCell;
}(__WEBPACK_IMPORTED_MODULE_4__utilities_grid_GridCell__["GridCell"]));
var ColorPickerGridCellBase = /** @class */ (function (_super) {
    __WEBPACK_IMPORTED_MODULE_0_tslib__["c" /* __extends */](ColorPickerGridCellBase, _super);
    function ColorPickerGridCellBase() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        /**
       * Render the core of a color cell
       * @returns {JSX.Element} - Element representing the core of the item
       */
        _this._onRenderColorOption = function (colorOption) {
            // Build an SVG for the cell with the given shape and color properties
            return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("svg", { className: _this._classNames.svg, viewBox: '0 0 20 20', fill: Object(__WEBPACK_IMPORTED_MODULE_3__utilities_color_colors__["getColorFromString"])(colorOption.color).str }, _this.props.circle ?
                __WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("circle", { cx: '50%', cy: '50%', r: '50%' }) :
                __WEBPACK_IMPORTED_MODULE_1_react__["createElement"]("rect", { width: '100%', height: '100%' })));
        };
        /**
         * Method to override the getClassNames func in a button.
         */
        _this._getClassNames = function (theme, className, variantClassName, iconClassName, menuIconClassName, disabled, checked, expanded, isSplit) {
            var styles = Object(__WEBPACK_IMPORTED_MODULE_5__Button_ActionButton_ActionButton_styles__["getStyles"])(theme);
            return Object(__WEBPACK_IMPORTED_MODULE_6__Styling__["mergeStyleSets"])(_this._classNames, {
                root: [
                    'ms-Button',
                    styles.root,
                    variantClassName,
                    className,
                    checked && [
                        'is-checked',
                        styles.rootChecked
                    ],
                    disabled && [
                        'is-disabled',
                        styles.rootDisabled
                    ],
                    !disabled && !checked && {
                        selectors: {
                            ':hover': styles.rootHovered,
                            ':focus': styles.rootFocused,
                            ':active': styles.rootPressed,
                        }
                    },
                    disabled && checked && [
                        styles.rootCheckedDisabled
                    ],
                    !disabled && checked && {
                        selectors: {
                            ':hover': styles.rootCheckedHovered,
                            ':active': styles.rootCheckedPressed
                        }
                    }
                ],
                flexContainer: [
                    'ms-Button-flexContainer',
                    styles.flexContainer
                ]
            });
        };
        return _this;
    }
    ColorPickerGridCellBase.prototype.render = function () {
        var _a = this.props, item = _a.item, id = _a.id, selected = _a.selected, disabled = _a.disabled, getStyles = _a.getStyles, theme = _a.theme, circle = _a.circle, color = _a.color, onClick = _a.onClick, onHover = _a.onHover, onFocus = _a.onFocus, onMouseEnter = _a.onMouseEnter, onMouseMove = _a.onMouseMove, onMouseLeave = _a.onMouseLeave, onWheel = _a.onWheel, onKeyDown = _a.onKeyDown;
        this._classNames = getClassNames(getStyles, {
            theme: theme,
            disabled: disabled,
            selected: selected,
            circle: circle,
            isWhite: this._isWhiteCell(color)
        });
        return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"](ColorCell, { item: item, id: id + "-" + item.id + "-" + item.index, key: item.id, disabled: disabled, role: 'gridcell', onRenderItem: this._onRenderColorOption, selected: selected, onClick: onClick, onHover: onHover, onFocus: onFocus, label: item.label, className: this._classNames.colorCell, getClassNames: this._getClassNames, index: item.index, onMouseEnter: onMouseEnter, onMouseMove: onMouseMove, onMouseLeave: onMouseLeave, onWheel: onWheel, onKeyDown: onKeyDown }));
    };
    /**
     * Validate if the cell's color is white or not to apply whiteCell style
     * @param inputColor - The color of the current cell
     * @returns - Whether the cell's color is white or not.
     */
    ColorPickerGridCellBase.prototype._isWhiteCell = function (inputColor) {
        return inputColor.toLocaleLowerCase() === '#ffffff';
    };
    ColorPickerGridCellBase.defaultProps = {
        circle: true,
        disabled: false,
        selected: false,
    };
    ColorPickerGridCellBase = __WEBPACK_IMPORTED_MODULE_0_tslib__["b" /* __decorate */]([
        Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["customizable"])('ColorPickerGridCell', ['theme'])
    ], ColorPickerGridCellBase);
    return ColorPickerGridCellBase;
}(__WEBPACK_IMPORTED_MODULE_1_react__["Component"]));



/***/ }),

/***/ 315:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/utilities/color/colors.js
module.exports = __webpack_require__(316);

/***/ }),

/***/ 316:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "MAX_COLOR_SATURATION", function() { return MAX_COLOR_SATURATION; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "MAX_COLOR_HUE", function() { return MAX_COLOR_HUE; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "MAX_COLOR_VALUE", function() { return MAX_COLOR_VALUE; });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "MAX_COLOR_RGBA", function() { return MAX_COLOR_RGBA; });
/* harmony export (immutable) */ __webpack_exports__["cssColor"] = cssColor;
/* harmony export (immutable) */ __webpack_exports__["rgb2hex"] = rgb2hex;
/* harmony export (immutable) */ __webpack_exports__["hsv2hex"] = hsv2hex;
/* harmony export (immutable) */ __webpack_exports__["rgb2hsv"] = rgb2hsv;
/* harmony export (immutable) */ __webpack_exports__["hsl2hsv"] = hsl2hsv;
/* harmony export (immutable) */ __webpack_exports__["hsv2hsl"] = hsv2hsl;
/* harmony export (immutable) */ __webpack_exports__["hsl2rgb"] = hsl2rgb;
/* harmony export (immutable) */ __webpack_exports__["hsv2rgb"] = hsv2rgb;
/* harmony export (immutable) */ __webpack_exports__["getColorFromString"] = getColorFromString;
/* harmony export (immutable) */ __webpack_exports__["getColorFromRGBA"] = getColorFromRGBA;
/* harmony export (immutable) */ __webpack_exports__["getFullColorString"] = getFullColorString;
/* harmony export (immutable) */ __webpack_exports__["updateSV"] = updateSV;
/* harmony export (immutable) */ __webpack_exports__["updateH"] = updateH;
/* harmony export (immutable) */ __webpack_exports__["updateA"] = updateA;
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__colorValues__ = __webpack_require__(317);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__colorValues___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1__colorValues__);


var MAX_COLOR_SATURATION = 100;
var MAX_COLOR_HUE = 359;
var MAX_COLOR_VALUE = 100;
var MAX_COLOR_RGBA = 255;
function cssColor(color) {
    return (_named(color)
        || _hex3(color)
        || _hex6(color)
        || _rgb(color)
        || _rgba(color)
        || _hsl(color)
        || _hsla(color));
}
function rgb2hex(r, g, b) {
    return [
        _numberToPaddedHex(r),
        _numberToPaddedHex(g),
        _numberToPaddedHex(b)
    ].join('');
}
function hsv2hex(h, s, v) {
    var _a = hsv2rgb(h, s, v), r = _a.r, g = _a.g, b = _a.b;
    return rgb2hex(r, g, b);
}
function rgb2hsv(r, g, b) {
    var h = NaN;
    var s;
    var v;
    var max = Math.max(r, g, b);
    var min = Math.min(r, g, b);
    var delta = max - min;
    // hue
    if (delta === 0) {
        h = 0;
    }
    else if (r === max) {
        h = ((g - b) / delta) % 6;
    }
    else if (g === max) {
        h = (b - r) / delta + 2;
    }
    else if (b === max) {
        h = (r - g) / delta + 4;
    }
    h = Math.round(h * 60);
    if (h < 0) {
        h += 360;
    }
    // saturation
    s = Math.round((max === 0 ? 0 : (delta / max)) * 100);
    // value
    v = Math.round(max / 255 * 100);
    return { h: h, s: s, v: v };
}
function hsl2hsv(h, s, l) {
    s *= ((l < 50) ? l : (100 - l)) / 100;
    return {
        h: h,
        s: 2 * s / (l + s) * 100,
        v: l + s
    };
}
function hsv2hsl(h, s, v) {
    s /= MAX_COLOR_SATURATION;
    v /= MAX_COLOR_VALUE;
    var l = (2 - s) * v;
    var sl = s * v;
    sl /= (l <= 1) ? l : 2 - l;
    sl = sl || 0;
    l /= 2;
    return { h: h, s: sl * 100, l: l * 100 };
}
function hsl2rgb(h, s, l) {
    var hsv = hsl2hsv(h, s, l);
    return hsv2rgb(hsv.h, hsv.s, hsv.v);
}
function hsv2rgb(h, s, v) {
    s = s / 100;
    v = v / 100;
    var rgb = [];
    var c = v * s;
    var hh = h / 60;
    var x = c * (1 - Math.abs(hh % 2 - 1));
    var m = v - c;
    switch (Math.floor(hh)) {
        case 0:
            rgb = [c, x, 0];
            break;
        case 1:
            rgb = [x, c, 0];
            break;
        case 2:
            rgb = [0, c, x];
            break;
        case 3:
            rgb = [0, x, c];
            break;
        case 4:
            rgb = [x, 0, c];
            break;
        case 5:
            rgb = [c, 0, x];
            break;
    }
    return {
        r: Math.round(MAX_COLOR_RGBA * (rgb[0] + m)),
        g: Math.round(MAX_COLOR_RGBA * (rgb[1] + m)),
        b: Math.round(MAX_COLOR_RGBA * (rgb[2] + m))
    };
}
function getColorFromString(inputColor) {
    var color = cssColor(inputColor);
    if (!color) {
        return;
    }
    var a = color.a, b = color.b, g = color.g, r = color.r;
    var _a = rgb2hsv(r, g, b), h = _a.h, s = _a.s, v = _a.v;
    return {
        a: a,
        b: b,
        g: g,
        h: h,
        hex: rgb2hex(r, g, b),
        r: r,
        s: s,
        str: inputColor,
        v: v
    };
}
function getColorFromRGBA(rgba) {
    var a = rgba.a, b = rgba.b, g = rgba.g, r = rgba.r;
    var _a = rgb2hsv(r, g, b), h = _a.h, s = _a.s, v = _a.v;
    var hex = rgb2hex(r, g, b);
    return {
        a: a,
        b: b,
        g: g,
        h: h,
        hex: hex,
        r: r,
        s: s,
        str: (a === 100) ? "#" + hex : "rgba(" + r + ", " + g + ", " + b + ", " + a / 100 + ")",
        v: v
    };
}
function getFullColorString(color) {
    return "#" + hsv2hex(color.h, MAX_COLOR_SATURATION, MAX_COLOR_VALUE);
}
function updateSV(color, s, v) {
    var _a = hsv2rgb(color.h, s, v), r = _a.r, g = _a.g, b = _a.b;
    var hex = rgb2hex(r, g, b);
    return {
        a: color.a,
        b: b,
        g: g,
        h: color.h,
        hex: hex,
        r: r,
        s: s,
        str: (color.a === 100) ? "#" + hex : "rgba(" + r + ", " + g + ", " + b + ", " + color.a / 100 + ")",
        v: v
    };
}
function updateH(color, h) {
    var _a = hsv2rgb(h, color.s, color.v), r = _a.r, g = _a.g, b = _a.b;
    var hex = rgb2hex(r, g, b);
    return {
        a: color.a,
        b: b,
        g: g,
        h: h,
        hex: hex,
        r: r,
        s: color.s,
        str: (color.a === 100) ? "#" + hex : "rgba(" + r + ", " + g + ", " + b + ", " + color.a / 100 + ")",
        v: color.v
    };
}
function updateA(color, a) {
    return Object(__WEBPACK_IMPORTED_MODULE_0__Utilities__["assign"])({}, color, {
        a: a,
        str: (a === 100) ? "#" + color.hex : "rgba(" + color.r + ", " + color.g + ", " + color.b + ", " + a / 100 + ")"
    });
}
function _numberToPaddedHex(num) {
    var hex = num.toString(16);
    return hex.length === 1 ? '0' + hex : hex;
}
function _named(str) {
    var c = __WEBPACK_IMPORTED_MODULE_1__colorValues__["COLOR_VALUES"][str.toLowerCase()];
    if (c) {
        return {
            r: c[0],
            g: c[1],
            b: c[2],
            a: 100
        };
    }
}
function _rgb(str) {
    if (0 === str.indexOf('rgb(')) {
        str = (str.match(/rgb\(([^)]+)\)/))[1];
        var parts = str.split(/ *, */).map(Number);
        return {
            r: parts[0],
            g: parts[1],
            b: parts[2],
            a: 100
        };
    }
}
function _rgba(str) {
    if (str.indexOf('rgba(') === 0) {
        str = (str.match(/rgba\(([^)]+)\)/))[1];
        var parts = str.split(/ *, */).map(Number);
        return {
            r: parts[0],
            g: parts[1],
            b: parts[2],
            a: parts[3] * 100
        };
    }
}
function _hex6(str) {
    if ('#' === str[0] && 7 === str.length) {
        return {
            r: parseInt(str.slice(1, 3), 16),
            g: parseInt(str.slice(3, 5), 16),
            b: parseInt(str.slice(5, 7), 16),
            a: 100
        };
    }
}
function _hex3(str) {
    if ('#' === str[0] && 4 === str.length) {
        return {
            r: parseInt(str[1] + str[1], 16),
            g: parseInt(str[2] + str[2], 16),
            b: parseInt(str[3] + str[3], 16),
            a: 100
        };
    }
}
function _hsl(str) {
    if (str.indexOf('hsl(') === 0) {
        str = (str.match(/hsl\(([^)]+)\)/))[1];
        var parts = str.split(/ *, */);
        var h = parseInt(parts[0], 10);
        var s = parseInt(parts[1], 10);
        var l = parseInt(parts[2], 10);
        var rgba = hsl2rgb(h, s, l);
        rgba.a = 100;
        return rgba;
    }
}
function _hsla(str) {
    if (str.indexOf('hsla(') === 0) {
        str = (str.match(/hsla\(([^)]+)\)/))[1];
        var parts = str.split(/ *, */);
        var h = parseInt(parts[0], 10);
        var s = parseInt(parts[1], 10);
        var l = parseInt(parts[2], 10);
        var a = parseInt(parts[3], 10) * 100;
        var rgba = hsl2rgb(h, s, l);
        rgba.a = a;
        return rgba;
    }
}


/***/ }),

/***/ 317:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/utilities/color/colorValues.js
module.exports = __webpack_require__(318);

/***/ }),

/***/ 318:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "COLOR_VALUES", function() { return COLOR_VALUES; });
var COLOR_VALUES = {
    aliceblue: [240, 248, 255],
    antiquewhite: [250, 235, 215],
    aqua: [0, 255, 255],
    aquamarine: [127, 255, 212],
    azure: [240, 255, 255],
    beige: [245, 245, 220],
    bisque: [255, 228, 196],
    black: [0, 0, 0],
    blanchedalmond: [255, 235, 205],
    blue: [0, 0, 255],
    blueviolet: [138, 43, 226],
    brown: [165, 42, 42],
    burlywood: [222, 184, 135],
    cadetblue: [95, 158, 160],
    chartreuse: [127, 255, 0],
    chocolate: [210, 105, 30],
    coral: [255, 127, 80],
    cornflowerblue: [100, 149, 237],
    cornsilk: [255, 248, 220],
    crimson: [220, 20, 60],
    cyan: [0, 255, 255],
    darkblue: [0, 0, 139],
    darkcyan: [0, 139, 139],
    darkgoldenrod: [184, 132, 11],
    darkgray: [169, 169, 169],
    darkgreen: [0, 100, 0],
    darkgrey: [169, 169, 169],
    darkkhaki: [189, 183, 107],
    darkmagenta: [139, 0, 139],
    darkolivegreen: [85, 107, 47],
    darkorange: [255, 140, 0],
    darkorchid: [153, 50, 204],
    darkred: [139, 0, 0],
    darksalmon: [233, 150, 122],
    darkseagreen: [143, 188, 143],
    darkslateblue: [72, 61, 139],
    darkslategray: [47, 79, 79],
    darkslategrey: [47, 79, 79],
    darkturquoise: [0, 206, 209],
    darkviolet: [148, 0, 211],
    deeppink: [255, 20, 147],
    deepskyblue: [0, 191, 255],
    dimgray: [105, 105, 105],
    dimgrey: [105, 105, 105],
    dodgerblue: [30, 144, 255],
    firebrick: [178, 34, 34],
    floralwhite: [255, 255, 240],
    forestgreen: [34, 139, 34],
    fuchsia: [255, 0, 255],
    gainsboro: [220, 220, 220],
    ghostwhite: [248, 248, 255],
    gold: [255, 215, 0],
    goldenrod: [218, 165, 32],
    gray: [128, 128, 128],
    green: [0, 128, 0],
    greenyellow: [173, 255, 47],
    grey: [128, 128, 128],
    honeydew: [240, 255, 240],
    hotpink: [255, 105, 180],
    indianred: [205, 92, 92],
    indigo: [75, 0, 130],
    ivory: [255, 255, 240],
    khaki: [240, 230, 140],
    lavender: [230, 230, 250],
    lavenderblush: [255, 240, 245],
    lawngreen: [124, 252, 0],
    lemonchiffon: [255, 250, 205],
    lightblue: [173, 216, 230],
    lightcoral: [240, 128, 128],
    lightcyan: [224, 255, 255],
    lightgoldenrodyellow: [250, 250, 210],
    lightgray: [211, 211, 211],
    lightgreen: [144, 238, 144],
    lightgrey: [211, 211, 211],
    lightpink: [255, 182, 193],
    lightsalmon: [255, 160, 122],
    lightseagreen: [32, 178, 170],
    lightskyblue: [135, 206, 250],
    lightslategray: [119, 136, 153],
    lightslategrey: [119, 136, 153],
    lightsteelblue: [176, 196, 222],
    lightyellow: [255, 255, 224],
    lime: [0, 255, 0],
    limegreen: [50, 205, 50],
    linen: [250, 240, 230],
    magenta: [255, 0, 255],
    maroon: [128, 0, 0],
    mediumaquamarine: [102, 205, 170],
    mediumblue: [0, 0, 205],
    mediumorchid: [186, 85, 211],
    mediumpurple: [147, 112, 219],
    mediumseagreen: [60, 179, 113],
    mediumslateblue: [123, 104, 238],
    mediumspringgreen: [0, 250, 154],
    mediumturquoise: [72, 209, 204],
    mediumvioletred: [199, 21, 133],
    midnightblue: [25, 25, 112],
    mintcream: [245, 255, 250],
    mistyrose: [255, 228, 225],
    moccasin: [255, 228, 181],
    navajowhite: [255, 222, 173],
    navy: [0, 0, 128],
    oldlace: [253, 245, 230],
    olive: [128, 128, 0],
    olivedrab: [107, 142, 35],
    orange: [255, 165, 0],
    orangered: [255, 69, 0],
    orchid: [218, 112, 214],
    palegoldenrod: [238, 232, 170],
    palegreen: [152, 251, 152],
    paleturquoise: [175, 238, 238],
    palevioletred: [219, 112, 147],
    papayawhip: [255, 239, 213],
    peachpuff: [255, 218, 185],
    peru: [205, 133, 63],
    pink: [255, 192, 203],
    plum: [221, 160, 203],
    powderblue: [176, 224, 230],
    purple: [128, 0, 128],
    rebeccapurple: [102, 51, 153],
    red: [255, 0, 0],
    rosybrown: [188, 143, 143],
    royalblue: [65, 105, 225],
    saddlebrown: [139, 69, 19],
    salmon: [250, 128, 114],
    sandybrown: [244, 164, 96],
    seagreen: [46, 139, 87],
    seashell: [255, 245, 238],
    sienna: [160, 82, 45],
    silver: [192, 192, 192],
    skyblue: [135, 206, 235],
    slateblue: [106, 90, 205],
    slategray: [119, 128, 144],
    slategrey: [119, 128, 144],
    snow: [255, 255, 250],
    springgreen: [0, 255, 127],
    steelblue: [70, 130, 180],
    tan: [210, 180, 140],
    teal: [0, 128, 128],
    thistle: [216, 191, 216],
    tomato: [255, 99, 71],
    turquoise: [64, 224, 208],
    violet: [238, 130, 238],
    wheat: [245, 222, 179],
    white: [255, 255, 255],
    whitesmoke: [245, 245, 245],
    yellow: [255, 255, 0],
    yellowgreen: [154, 205, 5]
};


/***/ }),

/***/ 319:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/utilities/grid/GridCell.js
module.exports = __webpack_require__(320);

/***/ }),

/***/ 320:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "GridCell", function() { return GridCell; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0_tslib__ = __webpack_require__(0);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react__ = __webpack_require__(2);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1_react___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1_react__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__Button__ = __webpack_require__(126);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_3__Button___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_3__Button__);




var GridCell = /** @class */ (function (_super) {
    __WEBPACK_IMPORTED_MODULE_0_tslib__["c" /* __extends */](GridCell, _super);
    function GridCell() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._onClick = function () {
            var _a = _this.props, onClick = _a.onClick, disabled = _a.disabled, item = _a.item;
            if (onClick && !disabled) {
                onClick(item);
            }
        };
        _this._onMouseEnter = function (ev) {
            var _a = _this.props, onHover = _a.onHover, disabled = _a.disabled, item = _a.item, onMouseEnter = _a.onMouseEnter;
            var didUpdateOnEnter = onMouseEnter && onMouseEnter(ev);
            if (!didUpdateOnEnter && onHover && !disabled) {
                onHover(item);
            }
        };
        _this._onMouseMove = function (ev) {
            var _a = _this.props, onHover = _a.onHover, disabled = _a.disabled, item = _a.item, onMouseMove = _a.onMouseMove;
            var didUpdateOnMove = onMouseMove && onMouseMove(ev);
            if (!didUpdateOnMove && onHover && !disabled) {
                onHover(item);
            }
        };
        _this._onMouseLeave = function (ev) {
            var _a = _this.props, onHover = _a.onHover, disabled = _a.disabled, onMouseLeave = _a.onMouseLeave;
            var didUpdateOnLeave = onMouseLeave && onMouseLeave(ev);
            if (!didUpdateOnLeave && onHover && !disabled) {
                onHover();
            }
        };
        _this._onFocus = function () {
            var _a = _this.props, onFocus = _a.onFocus, disabled = _a.disabled, item = _a.item;
            if (onFocus && !disabled) {
                onFocus(item);
            }
        };
        return _this;
    }
    GridCell.prototype.render = function () {
        var _a = this.props, item = _a.item, id = _a.id, className = _a.className, role = _a.role, selected = _a.selected, disabled = _a.disabled, onRenderItem = _a.onRenderItem, cellDisabledStyle = _a.cellDisabledStyle, cellIsSelectedStyle = _a.cellIsSelectedStyle, index = _a.index, label = _a.label, getClassNames = _a.getClassNames;
        return (__WEBPACK_IMPORTED_MODULE_1_react__["createElement"](__WEBPACK_IMPORTED_MODULE_3__Button__["CommandButton"], { id: id, "data-index": index, "data-is-focusable": true, disabled: disabled, className: Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["css"])(className, (_b = {},
                _b['' + cellIsSelectedStyle] = selected,
                _b['' + cellDisabledStyle] = disabled,
                _b)), onClick: this._onClick, onMouseEnter: this._onMouseEnter, onMouseMove: this._onMouseMove, onMouseLeave: this._onMouseLeave, onFocus: this._onFocus, role: role, "aria-selected": selected, ariaLabel: label, title: label, getClassNames: getClassNames }, onRenderItem(item)));
        var _b;
    };
    GridCell.defaultProps = {
        disabled: false,
        id: Object(__WEBPACK_IMPORTED_MODULE_2__Utilities__["getId"])('gridCell')
    };
    return GridCell;
}(__WEBPACK_IMPORTED_MODULE_1_react__["Component"]));



/***/ }),

/***/ 321:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/Button/ActionButton/ActionButton.styles.js
module.exports = __webpack_require__(322);

/***/ }),

/***/ 322:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getStyles", function() { return getStyles; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Styling__ = __webpack_require__(119);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Styling___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Styling__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__BaseButton_styles__ = __webpack_require__(323);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_2__BaseButton_styles___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_2__BaseButton_styles__);



var DEFAULT_BUTTON_HEIGHT = '40px';
var DEFAULT_PADDING = '0 4px';
var getStyles = Object(__WEBPACK_IMPORTED_MODULE_1__Utilities__["memoizeFunction"])(function (theme, customStyles) {
    var baseButtonStyles = Object(__WEBPACK_IMPORTED_MODULE_2__BaseButton_styles__["getStyles"])(theme);
    var actionButtonStyles = {
        root: {
            padding: DEFAULT_PADDING,
            height: DEFAULT_BUTTON_HEIGHT,
            color: theme.palette.neutralPrimary,
            backgroundColor: 'transparent'
        },
        rootHovered: {
            color: theme.palette.themePrimary,
        },
        iconHovered: {
            color: theme.palette.themePrimary
        },
        rootPressed: {
            color: theme.palette.black,
        },
        rootExpanded: {
            color: theme.palette.themePrimary
        },
        iconPressed: {
            color: theme.palette.themeDarker
        },
        rootDisabled: {
            color: theme.palette.neutralTertiary,
            backgroundColor: 'transparent'
        },
        rootChecked: {
            color: theme.palette.black,
        },
        iconChecked: {
            color: theme.palette.themeDarker
        },
        flexContainer: {
            justifyContent: 'flex-start'
        },
        icon: {
            color: theme.palette.themeDarkAlt
        },
        iconDisabled: {
            color: 'inherit'
        },
        menuIcon: {
            color: theme.palette.neutralSecondary
        },
        textContainer: {
            flexGrow: 0
        }
    };
    return Object(__WEBPACK_IMPORTED_MODULE_0__Styling__["concatStyleSets"])(baseButtonStyles, actionButtonStyles, customStyles);
});


/***/ }),

/***/ 323:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/Button/BaseButton.styles.js
module.exports = __webpack_require__(324);

/***/ }),

/***/ 324:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getStyles", function() { return getStyles; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities__ = __webpack_require__(36);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Utilities___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Utilities__);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Styling__ = __webpack_require__(119);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_1__Styling___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_1__Styling__);


var noOutline = {
    outline: 0
};
var iconStyle = {
    fontSize: __WEBPACK_IMPORTED_MODULE_1__Styling__["FontSizes"].icon,
    margin: '0 4px',
    height: '16px',
    lineHeight: '16px',
    textAlign: 'center',
    verticalAlign: 'middle',
    flexShrink: 0
};
/**
 * Gets the base button styles. Note: because it is a base class to be used with the `mergeRules`
 * helper, it should have values for all class names in the interface. This let `mergeRules` optimize
 * mixing class names together.
 */
var getStyles = Object(__WEBPACK_IMPORTED_MODULE_0__Utilities__["memoizeFunction"])(function (theme) {
    var semanticColors = theme.semanticColors;
    var border = semanticColors.buttonBorder;
    var disabledBackground = semanticColors.disabledBackground;
    var disabledText = semanticColors.disabledText;
    var buttonHighContrastFocus = {
        left: -2,
        top: -2,
        bottom: -2,
        right: -2,
        border: 'none',
        outlineColor: 'ButtonText'
    };
    return {
        root: [
            Object(__WEBPACK_IMPORTED_MODULE_1__Styling__["getFocusStyle"])(theme, -1, 'relative', buttonHighContrastFocus),
            theme.fonts.medium,
            {
                boxSizing: 'border-box',
                border: '1px solid ' + border,
                userSelect: 'none',
                display: 'inline-block',
                textDecoration: 'none',
                textAlign: 'center',
                cursor: 'pointer',
                verticalAlign: 'top',
                padding: '0 16px',
                borderRadius: 0
            }
        ],
        rootDisabled: {
            backgroundColor: disabledBackground,
            color: disabledText,
            cursor: 'default',
            pointerEvents: 'none',
            selectors: {
                ':hover': noOutline,
                ':focus': noOutline
            }
        },
        iconDisabled: {
            color: disabledText
        },
        menuIconDisabled: {
            color: disabledText
        },
        flexContainer: {
            display: 'flex',
            height: '100%',
            flexWrap: 'nowrap',
            justifyContent: 'center',
            alignItems: 'center'
        },
        textContainer: {
            flexGrow: 1
        },
        icon: iconStyle,
        menuIcon: [
            iconStyle,
            {
                fontSize: __WEBPACK_IMPORTED_MODULE_1__Styling__["FontSizes"].small
            }
        ],
        label: {
            margin: '0 4px',
            lineHeight: '100%'
        },
        screenReaderText: __WEBPACK_IMPORTED_MODULE_1__Styling__["hiddenContentStyle"]
    };
});


/***/ }),

/***/ 325:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/SwatchColorPicker/ColorPickerGridCell.styles.js
module.exports = __webpack_require__(326);

/***/ }),

/***/ 326:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getStyles", function() { return getStyles; });
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Styling__ = __webpack_require__(119);
/* harmony import */ var __WEBPACK_IMPORTED_MODULE_0__Styling___default = __webpack_require__.n(__WEBPACK_IMPORTED_MODULE_0__Styling__);

var ACTIVE_BORDER_COLOR = '#969696';
function getSvgSelectorStyles(borderColor, isHover) {
    return {
        width: 12,
        height: 12,
        border: '4px solid',
        borderColor: borderColor,
        boxShadow: isHover ? 'none' : '0 0 0 1px #969696',
        padding: 4,
        margin: 0
    };
}
var getStyles = function (props) {
    var theme = props.theme, disabled = props.disabled, selected = props.selected, circle = props.circle, isWhite = props.isWhite;
    var semanticColors = theme.semanticColors;
    return {
        colorCell: [
            {
                backgroundColor: 'transparent',
                padding: 0,
                overflow: 'visible',
                position: 'relative',
                boxSizing: 'border-box',
                display: 'inline-block',
                border: '1px solid transparent',
                background: 'transparent',
                cursor: 'pointer',
                textAlign: 'center',
                verticalAlign: 'top',
                userSelect: 'none',
                height: 40,
                selectors: (_a = {},
                    _a[__WEBPACK_IMPORTED_MODULE_0__Styling__["HighContrastSelector"]] = { border: 'none' },
                    _a['.ms-Fabric.is-focusVisible &:focus, .ms-Fabric.is-focusVisible &:focus::after'] = { border: 'none' },
                    _a['.ms-Fabric.is-focusVisible &:focus $svg'] = getSvgSelectorStyles(theme.palette.neutralQuaternaryAlt, false),
                    _a[':hover $svg'] = getSvgSelectorStyles(theme.palette.neutralQuaternaryAlt, true),
                    _a[':focus $svg'] = getSvgSelectorStyles(theme.palette.neutralQuaternaryAlt, false),
                    _a[':active $svg'] = getSvgSelectorStyles(ACTIVE_BORDER_COLOR, false),
                    _a)
            },
            isWhite && {
                selectors: {
                    '$svg': {
                        padding: 0,
                        border: '1px solid',
                        borderColor: theme.palette.neutralTertiary,
                        margin: 4,
                    },
                }
            },
            circle && 'is-circle' && {
                selectors: {
                    '$svg': { borderRadius: '100%' },
                }
            },
            selected && 'isSelected' && {
                selectors: {
                    '$svg': {
                        boxShadow: '0 0 0 1px #969696',
                        border: '4px solid',
                        borderColor: theme.palette.neutralTertiaryAlt,
                        width: 12,
                        height: 12,
                    },
                    ':hover $svg': { boxShadow: '0 0 0 1px #969696' },
                    ':focus $svg': {
                        boxShadow: '0 0 0 1px #969696',
                    },
                    ':active $svg': {
                        boxShadow: '0 0 0 1px #969696',
                        borderColor: ACTIVE_BORDER_COLOR,
                    },
                },
            },
            selected && isWhite && {
                selectors: {
                    '$svg': {
                        padding: 4,
                        margin: 0
                    }
                }
            },
            disabled && 'is-disabled' && {
                color: semanticColors.disabledBodyText,
                cursor: 'default',
                pointerEvents: 'none',
                opacity: .3
            }
        ],
        svg: [
            {
                width: 20,
                height: 20,
                padding: 4,
                boxSizing: 'content-box',
            }
        ]
    };
    var _a;
};


/***/ }),

/***/ 327:
/***/ (function(module, exports, __webpack_require__) {

// Loading office-ui-fabric-react/components/SwatchColorPicker/SwatchColorPicker.styles.js
module.exports = __webpack_require__(328);

/***/ }),

/***/ 328:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });
/* harmony export (binding) */ __webpack_require__.d(__webpack_exports__, "getStyles", function() { return getStyles; });
var getStyles = function (props) {
    var className = props.className;
    return {
        root: {
            padding: 2,
            outline: 'none'
        },
        tableCell: {
            padding: 0
        },
        focusedContainer: [
            'ms-swatchColorPickerBodyContainer',
            {
                clear: 'both',
                display: 'block',
                minWidth: '180px',
            },
            className
        ]
    };
};


/***/ }),

/***/ 329:
/***/ (function(module, exports, __webpack_require__) {

var content = __webpack_require__(330);
var loader = __webpack_require__(3);

if(typeof content === "string") content = [[module.i, content]];

// add the styles to the DOM
for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1], true);

if(content.locals) module.exports = content.locals;

/***/ }),

/***/ 330:
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__(4)(false);
// imports


// module
exports.push([module.i, ".sprtePropertyPaneButton_7b16c658{background:0 0;border:none}.sprtePropertyPaneButton_7b16c658.iconButton_7b16c658:hover{background-color:\"[theme:neutralLight, default: #eaeaea]\"}.sprtePropertyPaneButton_7b16c658.iconButton_7b16c658:active{background-color:#dadada}.sprtePropertyPaneButton_7b16c658.iconButton_7b16c658{width:40px;height:40px;padding:0}.sprtePropertyPaneButton_7b16c658.imageButton_7b16c658{width:56px;height:56px;min-width:56px;padding:0}.sprtePropertyPaneButton_7b16c658 i{font-size:20px}.sprtePropertyPaneButton_7b16c658.tableButtonInPane_7b16c658 i{font-size:24px}.sprtePropertyPaneButton_7b16c658.sprtePropertyPaneButtonActive_7b16c658{color:\"[theme:themePrimary, default:#0078d4]\";background-color:#dadada}.sprtePropertyPaneButton_7b16c658:hover{color:\"[theme:themePrimary, default:#0078d4]\"}.sprtePropertyPaneButton_7b16c658.textButton_7b16c658{background:\"[theme:themePrimary, default:#0078d4]\";color:\"[theme:white, default: #ffffff]\"}.sprtePropertyPaneButton_7b16c658.textButton_7b16c658.disabled_7b16c658{background:\"[theme:neutralTertiary, default: #a6a6a6]\"}.sprtePropertyPaneButton_7b16c658.textButton_7b16c658:hover{background:\"[theme:themePrimary, default:#0078d4]\"}.sprtePropertyPaneButtonImage_7b16c658{height:40px;background-color:\"[theme:white, default: #ffffff]\"}", ""]);

// exports


/***/ }),

/***/ 331:
/***/ (function(module, exports, __webpack_require__) {

var content = __webpack_require__(332);
var loader = __webpack_require__(3);

if(typeof content === "string") content = [[module.i, content]];

// add the styles to the DOM
for (var i = 0; i < content.length; i++) loader.loadStyles(content[i][1], true);

if(content.locals) module.exports = content.locals;

/***/ }),

/***/ 332:
/***/ (function(module, exports, __webpack_require__) {

exports = module.exports = __webpack_require__(4)(false);
// imports


// module
exports.push([module.i, ".themeableSvgTableIcon_24371c77{width:40px;height:40px}.cls1_24371c77,.cls6_24371c77{fill:none}.cls1_24371c77,.cls2_24371c77,.cls3_24371c77{stroke:\"[theme:neutralPrimary, default: #333333]\"}.cls2_24371c77,.cls4_24371c77{fill:\"[theme:neutralLighterAlt, default: #f8f8f8]\"}.cls3_24371c77{fill:\"[theme:neutralPrimary, default: #333333]\"}.cls4_24371c77{stroke:\"[theme:neutralLighterAlt, default: #f8f8f8]\";stroke-width:.6px}.cls5_24371c77{stroke:none}.deleteTable_24371c77.cls1_24371c77{fill:\"[theme:neutralTertiary, default: #a6a6a6]\";stroke:\"[theme:neutralPrimary, default: #333333]\"}.deleteTable_24371c77.cls2_24371c77{fill:\"[theme:neutralPrimary, default: #333333]\";stroke:none}.deleteTable_24371c77.cls3_24371c77,.deleteTable_24371c77.cls6_24371c77{fill:none}.deleteTable_24371c77.cls3_24371c77{stroke:\"[theme:neutralTertiaryAlt, default: #c8c8c8]\"}.deleteTable_24371c77.cls4_24371c77{fill:\"[theme:neutralLighterAlt, default: #f8f8f8]\";stroke:\"[theme:neutralLighterAlt, default: #f8f8f8]\";stroke-width:.6px}.deleteTable_24371c77.cls5_24371c77{stroke:none}.neutralPrimaryIcon_24371c77,.neutralRows_24371c77{fill:\"[theme:neutralPrimary, default: #333333]\"}.neutralRows_24371c77{opacity:.2}.neutralLeftColumn_24371c77{opacity:.5;fill:\"[theme:neutralPrimary, default: #333333]\"}.themePrimaryIcon_24371c77,.themeRows_24371c77{fill:\"[theme:themePrimary, default:#0078d4]\"}.themeRows_24371c77{opacity:.2}.themeLeftColumn_24371c77{opacity:.5;fill:\"[theme:themePrimary, default:#0078d4]\"}", ""]);

// exports


/***/ }),

/***/ 354:
/***/ (function(module, __webpack_exports__, __webpack_require__) {

"use strict";
Object.defineProperty(__webpack_exports__, "__esModule", { value: true });

// EXTERNAL MODULE: C:/agent/1/_work/60/s/common/temp/node_modules/.onedrive.pkgs.visualstudio.com/tslib/1.8.1/node_modules/tslib/tslib.es6.js
var tslib_es6 = __webpack_require__(0);

// EXTERNAL MODULE: external "react"
var external__react_ = __webpack_require__(2);
var external__react__default = /*#__PURE__*/__webpack_require__.n(external__react_);

// EXTERNAL MODULE: external "react-dom"
var external__react_dom_ = __webpack_require__(23);
var external__react_dom__default = /*#__PURE__*/__webpack_require__.n(external__react_dom_);

// EXTERNAL MODULE: external "@microsoft/sp-webpart-base"
var sp_webpart_base_ = __webpack_require__(30);
var sp_webpart_base__default = /*#__PURE__*/__webpack_require__.n(sp_webpart_base_);

// EXTERNAL MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/loc/RteStrings.resx.js
var RteStrings_resx = __webpack_require__(18);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/ControlsInOneRow.module.scss.js
__webpack_require__(297);
var styles = {
    controlsInOneRow: 'controlsInOneRow_b475b251',
    spaceInBetween: 'spaceInBetween_b475b251',
};
/* harmony default export */ var ControlsInOneRow_module_scss = (styles);

// EXTERNAL MODULE: external "@microsoft/office-ui-fabric-react-bundle"
var office_ui_fabric_react_bundle_ = __webpack_require__(1);
var office_ui_fabric_react_bundle__default = /*#__PURE__*/__webpack_require__.n(office_ui_fabric_react_bundle_);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/ControlsInOneRow.js




var ControlsInOneRow_ControlsInOneRow = (function (_super) {
    tslib_es6["c" /* __extends */](ControlsInOneRow, _super);
    function ControlsInOneRow() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    ControlsInOneRow.prototype.render = function () {
        var cssClass = Object(office_ui_fabric_react_bundle_["css"])(ControlsInOneRow_module_scss.controlsInOneRow, (_a = {}, _a[ControlsInOneRow_module_scss.spaceInBetween] = this.props.spaceInBetween, _a));
        return (external__react_["createElement"]("div", { className: cssClass }, this.props.children));
        var _a;
    };
    return ControlsInOneRow;
}(external__react_["Component"]));
/* harmony default export */ var PropertyPane_ControlsInOneRow = (ControlsInOneRow_ControlsInOneRow);

// EXTERNAL MODULE: external "@microsoft/sp-lodash-subset"
var sp_lodash_subset_ = __webpack_require__(17);
var sp_lodash_subset__default = /*#__PURE__*/__webpack_require__.n(sp_lodash_subset_);

// EXTERNAL MODULE: external "@microsoft/sp-core-library"
var sp_core_library_ = __webpack_require__(5);
var sp_core_library__default = /*#__PURE__*/__webpack_require__.n(sp_core_library_);

// EXTERNAL MODULE: C:/agent/1/_work/60/s/common/temp/node_modules/.onedrive.pkgs.visualstudio.com/office-ui-fabric-react/5.83.0/react-dom@15.6.2+react@15.6.2/node_modules/office-ui-fabric-react/lib/components/Tooltip/TooltipHost.js
var TooltipHost = __webpack_require__(62);
var TooltipHost_default = /*#__PURE__*/__webpack_require__.n(TooltipHost);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRteColorPicker.module.scss.js
__webpack_require__(299);
var SPRteColorPicker_module_scss_styles = {
    pickerCallout: 'pickerCallout_4733da1f',
    pickerLabel: 'pickerLabel_4733da1f',
    isRTL: 'isRTL_4733da1f',
    buttonLabel: 'buttonLabel_4733da1f',
    previewSvg: 'previewSvg_4733da1f',
    border: 'border_4733da1f',
    defaultSvg: 'defaultSvg_4733da1f',
    fillThemeColor: 'fillThemeColor_4733da1f',
    selected: 'selected_4733da1f',
    colorPickerButton: 'colorPickerButton_4733da1f',
    previewRectangle: 'previewRectangle_4733da1f',
    previewIcon: 'previewIcon_4733da1f',
};
/* harmony default export */ var SPRteColorPicker_module_scss = (SPRteColorPicker_module_scss_styles);

// EXTERNAL MODULE: C:/agent/1/_work/60/s/common/temp/node_modules/.onedrive.pkgs.visualstudio.com/office-ui-fabric-react/5.83.0/react-dom@15.6.2+react@15.6.2/node_modules/office-ui-fabric-react/lib/components/SwatchColorPicker/SwatchColorPicker.js
var SwatchColorPicker = __webpack_require__(301);
var SwatchColorPicker_default = /*#__PURE__*/__webpack_require__.n(SwatchColorPicker);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/ColorPickerGroup.js





var ColorPickerGroup_ColorPickerGroup = (function (_super) {
    tslib_es6["c" /* __extends */](ColorPickerGroup, _super);
    function ColorPickerGroup(props) {
        return _super.call(this, props) || this;
    }
    ColorPickerGroup.prototype.render = function () {
        var _a = this.props, groupText = _a.groupText, groupColors = _a.groupColors, onColorChanged = _a.onColorChanged, selectedColor = _a.selectedColor;
        var pickerLabelCss = Object(office_ui_fabric_react_bundle_["css"])(SPRteColorPicker_module_scss.pickerLabel, (_b = {}, _b[SPRteColorPicker_module_scss.isRTL] = Object(office_ui_fabric_react_bundle_["getRTL"])(), _b));
        return (external__react_["createElement"]("div", null,
            external__react_["createElement"]("label", { htmlFor: groupText, className: pickerLabelCss }, groupText),
            external__react_["createElement"](SwatchColorPicker["SwatchColorPicker"], { id: groupText, key: groupText, columnCount: 5, cellShape: 'square', onColorChanged: onColorChanged, colorCells: groupColors, selectedId: selectedColor })));
        var _b;
    };
    return ColorPickerGroup;
}(external__react_["Component"]));
/* harmony default export */ var PropertyPane_ColorPickerGroup = (ColorPickerGroup_ColorPickerGroup);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRteColorGroups.js

var SPRteColorGroups_SPRteColorGroups = (function () {
    function SPRteColorGroups() {
    }
    SPRteColorGroups.getSelectedColor = function (colors, selectedColor) {
        var filteredArray = colors.filter(function (colorProp) { return colorProp.color === selectedColor; });
        return filteredArray[0] && filteredArray[0].color;
    };
    SPRteColorGroups.buildFontColorMap = function (theme) {
        var palette = theme.palette;
        return new Map([
            [palette.redDark, 'RedDark'],
            [palette.red, 'Red'],
            [palette.yellow, 'Yellow'],
            [palette.yellowLight, 'YellowLight'],
            [palette.greenLight, 'GreenLight'],
            [palette.green, 'Green'],
            [palette.blueLight, 'BlueLight'],
            [palette.blue, 'Blue'],
            [palette.blueDark, 'BlueDark'],
            [palette.purple, 'Purple'],
            [palette.neutralPrimary, 'NeutralPrimary'],
            [palette.neutralSecondary, 'NeutralSecondary'],
            [palette.neutralTertiary, 'NeutralTertiary'],
            [palette.neutralDark, 'NeutralDark'],
            [palette.neutralPrimaryAlt, 'NeutralPrimaryAlt'],
            [palette.themeDark, 'ThemeDark'],
            [palette.themeDarker, 'ThemeDarker'],
            [palette.themeDarkAlt, 'ThemeDarkAlt'],
            [palette.themePrimary, 'ThemePrimary'],
            [palette.themeSecondary, 'ThemeSecondary']
        ] );
    };
    Object.defineProperty(SPRteColorGroups, "HighlightColorMap", {
        get: function () {
            return new Map([
                ["#FFFF00" , 'Yellow'],
                ["#00FF00" , 'Green'],
                ["#00FFFF" , 'Aqua'],
                ["#FF00FF" , 'Magenta'],
                ["#0000FF" , 'Blue'],
                ["#FF0000" , 'Red'],
                ["#000080" , 'DarkBlue'],
                ["#008080" , 'Teal'],
                ["#008000" , 'DarkGreen'],
                ["#800080" , 'Purple'],
                ["#800000" , 'Maroon'],
                ["#808000" , 'Gold'],
                ["#808080" , 'DarkGrey'],
                ["#C0C0C0" , 'Grey'],
                ["#000000" , 'Black']
            ] );
        },
        enumerable: true,
        configurable: true
    });
    SPRteColorGroups.standardColorGroup = function (theme) {
        var palette = theme.palette;
        var standardColorGroup = {
            groupText: RteStrings_resx["a" /* default */].StandardColorLabel,
            groupColors: [
                { id: palette.redDark, label: RteStrings_resx["a" /* default */].RedDarkLabel, color: palette.redDark },
                { id: palette.red, label: RteStrings_resx["a" /* default */].RedLabel, color: palette.red },
                { id: palette.yellow, label: RteStrings_resx["a" /* default */].OrangeLabel, color: palette.yellow },
                { id: palette.yellowLight, label: RteStrings_resx["a" /* default */].YellowLabel, color: palette.yellowLight },
                { id: palette.greenLight, label: RteStrings_resx["a" /* default */].GreenLightLabel, color: palette.greenLight },
                { id: palette.green, label: RteStrings_resx["a" /* default */].GreenLabel, color: palette.green },
                { id: palette.blueLight, label: RteStrings_resx["a" /* default */].BlueLightLabel, color: palette.blueLight },
                { id: palette.blue, label: RteStrings_resx["a" /* default */].BlueLabel, color: palette.blue },
                { id: palette.blueDark, label: RteStrings_resx["a" /* default */].BlueDarkLabel, color: palette.blueDark },
                { id: palette.purple, label: RteStrings_resx["a" /* default */].PurpleLabel, color: palette.purple }
            ]
        };
        return standardColorGroup;
    };
    SPRteColorGroups.themeColorGroup = function (theme) {
        var palette = theme.palette;
        var themeColorGroup = {
            groupText: RteStrings_resx["a" /* default */].ThemeColorGroupLabel,
            groupColors: [
                { id: palette.themeDarker, label: RteStrings_resx["a" /* default */].ThemeDarkerLabel, color: palette.themeDarker },
                { id: palette.themeDark, label: RteStrings_resx["a" /* default */].ThemeDarkLabel, color: palette.themeDark },
                { id: palette.themeDarkAlt, label: RteStrings_resx["a" /* default */].ThemeDarkAltLabel, color: palette.themeDarkAlt },
                { id: palette.themePrimary, label: RteStrings_resx["a" /* default */].ThemePrimaryLabel, color: palette.themePrimary },
                { id: palette.themeSecondary, label: RteStrings_resx["a" /* default */].ThemeSecondaryLabel, color: palette.themeSecondary },
                { id: palette.neutralTertiary, label: RteStrings_resx["a" /* default */].NeutralTertiaryLabel, color: palette.neutralTertiary },
                { id: palette.neutralSecondary, label: RteStrings_resx["a" /* default */].NeutralSecondaryLabel, color: palette.neutralSecondary },
                { id: palette.neutralPrimaryAlt, label: RteStrings_resx["a" /* default */].NeutralPrimaryAltLabel, color: palette.neutralPrimaryAlt },
                { id: palette.neutralPrimary, label: RteStrings_resx["a" /* default */].NeutralPrimaryLabel, color: palette.neutralPrimary },
                { id: palette.neutralDark, label: RteStrings_resx["a" /* default */].NeutralDarkLabel, color: palette.neutralDark }
            ]
        };
        return themeColorGroup;
    };
    SPRteColorGroups.highlightColorGroup = {
        groupText: RteStrings_resx["a" /* default */].HightlightColorsLabel,
        groupColors: [
            { id: "#FFFF00" , label: RteStrings_resx["a" /* default */].YellowLabel, color: "#FFFF00"  },
            { id: "#00FF00" , label: RteStrings_resx["a" /* default */].GreenLabel, color: "#00FF00"  },
            { id: "#00FFFF" , label: RteStrings_resx["a" /* default */].AquaLabel, color: "#00FFFF"  },
            { id: "#FF00FF" , label: RteStrings_resx["a" /* default */].MagentaLabel, color: "#FF00FF"  },
            { id: "#0000FF" , label: RteStrings_resx["a" /* default */].BlueLabel, color: "#0000FF"  },
            { id: "#FF0000" , label: RteStrings_resx["a" /* default */].RedLabel, color: "#FF0000"  },
            { id: "#000080" , label: RteStrings_resx["a" /* default */].BlueDarkLabel, color: "#000080"  },
            { id: "#008080" , label: RteStrings_resx["a" /* default */].TealLabel, color: "#008080"  },
            { id: "#008000" , label: RteStrings_resx["a" /* default */].GreenDarkLabel, color: "#008000"  },
            { id: "#800080" , label: RteStrings_resx["a" /* default */].PurpleLabel, color: "#800080"  },
            { id: "#800000" , label: RteStrings_resx["a" /* default */].MaroonLabel, color: "#800000"  },
            { id: "#808000" , label: RteStrings_resx["a" /* default */].GoldLabel, color: "#808000"  },
            { id: "#808080" , label: RteStrings_resx["a" /* default */].DarkGreyLabel, color: "#808080"  },
            { id: "#C0C0C0" , label: RteStrings_resx["a" /* default */].GreyLabel, color: "#C0C0C0"  },
            { id: "#000000" , label: RteStrings_resx["a" /* default */].BlackLabel, color: "#000000"  }
        ]
    };
    return SPRteColorGroups;
}());
/* harmony default export */ var PropertyPane_SPRteColorGroups = (SPRteColorGroups_SPRteColorGroups);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRtePropertyPaneUtilities.js
var SPRtePropertyPaneUtilities = (function () {
    function SPRtePropertyPaneUtilities() {
    }
    SPRtePropertyPaneUtilities.isColorTransparent = function (previewColor) {
        return previewColor === 'rgba(0, 0, 0, 0)' || previewColor === 'transparent';
    };
    return SPRtePropertyPaneUtilities;
}());
/* harmony default export */ var PropertyPane_SPRtePropertyPaneUtilities = (SPRtePropertyPaneUtilities);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRteColorPicker.js










var SPRteColorPicker_SPRteColorPicker = (function (_super) {
    tslib_es6["c" /* __extends */](SPRteColorPicker, _super);
    function SPRteColorPicker(props) {
        var _this = _super.call(this, props) || this;
        _this._colorPickerGroups = [];
        _this.state = {
            isCalloutVisible: false
        };
        _this._setColorPickerGroups();
        return _this;
    }
    SPRteColorPicker.prototype.shouldComponentUpdate = function (nextProps, nextState) {
        var selectedColor = this.props.selectedColor;
        return (selectedColor !== nextProps.selectedColor ||
            this.state.isCalloutVisible !== nextState.isCalloutVisible ||
            !Object(sp_lodash_subset_["isEqual"])(nextProps.colorPickerGroups, this.props.colorPickerGroups));
    };
    SPRteColorPicker.prototype.componentWillUpdate = function (prevProps, prevState) {
        this._setColorPickerGroups();
    };
    SPRteColorPicker.prototype.render = function () {
        var _a = this.props, buttonLabel = _a.buttonLabel, switchToDefaultColor = _a.switchToDefaultColor, defaultButtonLabel = _a.defaultButtonLabel, fillThemeColor = _a.fillThemeColor, automationId = _a.automationId;
        var isCalloutVisible = this.state.isCalloutVisible;
        var previewSvgCss = Object(office_ui_fabric_react_bundle_["css"])(SPRteColorPicker_module_scss.previewSvg, (_b = {}, _b[SPRteColorPicker_module_scss.border] = PropertyPane_SPRtePropertyPaneUtilities.isColorTransparent(this.props.previewColor), _b));
        var tooltipCalloutProps = { gapSpace: 0 };
        var defaultSvgClassName = Object(office_ui_fabric_react_bundle_["css"])(SPRteColorPicker_module_scss.previewSvg, SPRteColorPicker_module_scss.defaultSvg, (_c = {}, _c[SPRteColorPicker_module_scss.selected] = !this._selectedColor, _c), (_d = {}, _d[SPRteColorPicker_module_scss.fillThemeColor] = fillThemeColor, _d));
        var colorPickerButtonId = automationId + "-propertyPaneButtonToolTip";
        var defaultButtonId = automationId + "-defaultButtonTooltip";
        var pickerCalloutClassName = sp_core_library_["_BrowserDetection"].getBrowserInformation().browser === 5  ? undefined : SPRteColorPicker_module_scss.pickerCallout;
        return (external__react_["createElement"]("div", null,
            external__react_["createElement"]("div", { ref: this._resolveRef('_buttonContainerRef') },
                external__react_["createElement"](TooltipHost["TooltipHost"], { content: buttonLabel, id: colorPickerButtonId, calloutProps: tooltipCalloutProps },
                    external__react_["createElement"](office_ui_fabric_react_bundle_["Button"], { className: SPRteColorPicker_module_scss.colorPickerButton, onClick: this._handleShowMenuClicked, "data-automation-id": automationId + "-propertyPaneButton", "aria-describedby": colorPickerButtonId },
                        external__react_["createElement"]("svg", { className: previewSvgCss, fill: this.props.previewColor, viewBox: '0 0 20 20' }, this._previewSvgRectangle),
                        external__react_["createElement"]("div", { className: SPRteColorPicker_module_scss.buttonLabel }, buttonLabel),
                        external__react_["createElement"](office_ui_fabric_react_bundle_["Icon"], { className: SPRteColorPicker_module_scss.previewIcon, iconName: 'CaretDownSolid8', role: 'presentation' })))),
            isCalloutVisible &&
                external__react_["createElement"](office_ui_fabric_react_bundle_["Callout"], { className: pickerCalloutClassName, directionalHint: 4 , gapSpace: 0, isBeakVisible: false, onDismiss: this._handleCalloutDismiss, role: 'alertdialog', setInitialFocus: true, target: this._buttonContainerRef },
                    external__react_["createElement"](TooltipHost["TooltipHost"], { content: defaultButtonLabel, id: defaultButtonId, calloutProps: tooltipCalloutProps },
                        external__react_["createElement"](office_ui_fabric_react_bundle_["Button"], { className: SPRteColorPicker_module_scss.colorPickerButton, onClick: switchToDefaultColor, "data-automation-id": automationId + "-defaultButton", "aria-describedby": defaultButtonId },
                            external__react_["createElement"]("svg", { className: defaultSvgClassName, viewBox: '0 0 20 20' }, this._previewSvgRectangle),
                            external__react_["createElement"]("div", { className: SPRteColorPicker_module_scss.buttonLabel }, defaultButtonLabel))),
                    this._colorPickerGroups)));
        var _b, _c, _d;
    };
    Object.defineProperty(SPRteColorPicker.prototype, "_previewSvgRectangle", {
        get: function () {
            return external__react_["createElement"]("rect", { className: SPRteColorPicker_module_scss.previewRectangle, width: '100%', height: '100%' });
        },
        enumerable: true,
        configurable: true
    });
    SPRteColorPicker.prototype._setColorPickerGroups = function () {
        var _a = this.props, colorPickerGroups = _a.colorPickerGroups, theme = _a.theme, selectedColor = _a.selectedColor;
        this._colorPickerGroups = [];
        for (var i = 0; i < colorPickerGroups.length; i++) {
            var colorPickerGroup = void 0;
            switch (colorPickerGroups[i]) {
                case "standardColors" :
                    colorPickerGroup = PropertyPane_SPRteColorGroups.standardColorGroup(theme);
                    break;
                case "themeColors" :
                    colorPickerGroup = PropertyPane_SPRteColorGroups.themeColorGroup(theme);
                    break;
                case "highlightColors" :
                    colorPickerGroup = PropertyPane_SPRteColorGroups.highlightColorGroup;
                    break;
            }
            var color = PropertyPane_SPRteColorGroups.getSelectedColor(colorPickerGroup.groupColors, selectedColor);
            if (color) {
                this._selectedColor = color;
            }
            this._colorPickerGroups.push(this._getColorPickerGroup(colorPickerGroup, color));
        }
    };
    SPRteColorPicker.prototype._getColorPickerGroup = function (colorPickerGroup, color) {
        return (external__react_["createElement"](PropertyPane_ColorPickerGroup, { key: colorPickerGroup.groupText, groupText: colorPickerGroup.groupText, groupColors: colorPickerGroup.groupColors, onColorChanged: this._handleColorChanged, selectedColor: color }));
    };
    SPRteColorPicker.prototype._handleColorChanged = function (id, color) {
        this.props.onColorChanged(id, color, this.props.theme);
    };
    SPRteColorPicker.prototype._handleShowMenuClicked = function () {
        this.setState({
            isCalloutVisible: !this.state.isCalloutVisible
        });
    };
    SPRteColorPicker.prototype._handleCalloutDismiss = function () {
        this.setState({
            isCalloutVisible: false
        });
    };
    tslib_es6["b" /* __decorate */]([
        office_ui_fabric_react_bundle_["autobind"]
    ], SPRteColorPicker.prototype, "_handleColorChanged", null);
    tslib_es6["b" /* __decorate */]([
        office_ui_fabric_react_bundle_["autobind"]
    ], SPRteColorPicker.prototype, "_handleShowMenuClicked", null);
    tslib_es6["b" /* __decorate */]([
        office_ui_fabric_react_bundle_["autobind"]
    ], SPRteColorPicker.prototype, "_handleCalloutDismiss", null);
    SPRteColorPicker = tslib_es6["b" /* __decorate */]([
        Object(office_ui_fabric_react_bundle_["customizable"])('SPRteColorPicker', ['theme'])
    ], SPRteColorPicker);
    return SPRteColorPicker;
}(office_ui_fabric_react_bundle_["BaseComponent"]));
/* harmony default export */ var PropertyPane_SPRteColorPicker = (SPRteColorPicker_SPRteColorPicker);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRtePropertyPaneButton.module.scss.js
__webpack_require__(329);
var SPRtePropertyPaneButton_module_scss_styles = {
    sprtePropertyPaneButton: 'sprtePropertyPaneButton_7b16c658',
    iconButton: 'iconButton_7b16c658',
    imageButton: 'imageButton_7b16c658',
    tableButtonInPane: 'tableButtonInPane_7b16c658',
    sprtePropertyPaneButtonActive: 'sprtePropertyPaneButtonActive_7b16c658',
    textButton: 'textButton_7b16c658',
    disabled: 'disabled_7b16c658',
    sprtePropertyPaneButtonImage: 'sprtePropertyPaneButtonImage_7b16c658',
};
/* harmony default export */ var SPRtePropertyPaneButton_module_scss = (SPRtePropertyPaneButton_module_scss_styles);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRtePropertyPaneButton.js





var SPRtePropertyPaneButton_SPRtPropertyPaneButton = (function (_super) {
    tslib_es6["c" /* __extends */](SPRtPropertyPaneButton, _super);
    function SPRtPropertyPaneButton(props) {
        return _super.call(this, props) || this;
    }
    SPRtPropertyPaneButton.prototype.render = function () {
        var _a = this.props, onClick = _a.onClick, fabricIconKey = _a.fabricIconKey, title = _a.title, ariaLabel = _a.ariaLabel, isActive = _a.isActive, imageSrc = _a.imageSrc, isDisabled = _a.isDisabled, text = _a.text, automationId = _a.automationId, className = _a.className;
        var buttonAutomationId = automationId + '-propertyPaneButton';
        var buttonClassName = Object(office_ui_fabric_react_bundle_["css"])(SPRtePropertyPaneButton_module_scss.sprtePropertyPaneButton, (_b = {},
            _b[SPRtePropertyPaneButton_module_scss.sprtePropertyPaneButtonActive] = isActive,
            _b[SPRtePropertyPaneButton_module_scss.imageButton] = !!imageSrc,
            _b[SPRtePropertyPaneButton_module_scss.iconButton] = fabricIconKey,
            _b[SPRtePropertyPaneButton_module_scss.textButton] = !!text && !imageSrc,
            _b[SPRtePropertyPaneButton_module_scss.disabled] = isDisabled,
            _b[SPRtePropertyPaneButton_module_scss[className]] = className,
            _b));
        var imageButtonClassName = Object(office_ui_fabric_react_bundle_["css"])(SPRtePropertyPaneButton_module_scss.sprtePropertyPaneButtonImage);
        return (external__react_["createElement"](TooltipHost["TooltipHost"], { content: title, id: buttonAutomationId, calloutProps: { gapSpace: 0 } }, fabricIconKey ?
            external__react_["createElement"](office_ui_fabric_react_bundle_["IconButton"], { ariaLabel: ariaLabel, "aria-describedby": buttonAutomationId, onClick: onClick, iconProps: { iconName: fabricIconKey }, disabled: isDisabled, className: buttonClassName, "data-automation-id": buttonAutomationId }) :
            external__react_["createElement"](office_ui_fabric_react_bundle_["DefaultButton"], { ariaLabel: ariaLabel, "aria-describedby": buttonAutomationId, onClick: onClick, disabled: isDisabled, className: buttonClassName, "data-automation-id": buttonAutomationId }, (imageSrc && !fabricIconKey) ?
                external__react_["createElement"]("div", { className: imageButtonClassName }, imageSrc)
                :
                    text)));
        var _b;
    };
    return SPRtPropertyPaneButton;
}(external__react_["Component"]));
/* harmony default export */ var SPRtePropertyPaneButton = (SPRtePropertyPaneButton_SPRtPropertyPaneButton);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/tableIcons/TableIcons.module.scss.js
__webpack_require__(331);
var TableIcons_module_scss_styles = {
    themeableSvgTableIcon: 'themeableSvgTableIcon_24371c77',
    cls1: 'cls1_24371c77',
    cls6: 'cls6_24371c77',
    cls2: 'cls2_24371c77',
    cls3: 'cls3_24371c77',
    cls4: 'cls4_24371c77',
    cls5: 'cls5_24371c77',
    deleteTable: 'deleteTable_24371c77',
    neutralPrimaryIcon: 'neutralPrimaryIcon_24371c77',
    neutralRows: 'neutralRows_24371c77',
    neutralLeftColumn: 'neutralLeftColumn_24371c77',
    themePrimaryIcon: 'themePrimaryIcon_24371c77',
    themeRows: 'themeRows_24371c77',
    themeLeftColumn: 'themeLeftColumn_24371c77',
};
/* harmony default export */ var TableIcons_module_scss = (TableIcons_module_scss_styles);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/tableIcons/TableIcons.js



var TableIcons_TableIcons = (function () {
    function TableIcons() {
    }
    TableIcons.getSvgFromSrc = function (imageSrc) {
        switch (imageSrc) {
            case 'deleteColumn.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-19988 21347 28 33' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-21036.9 20446)' },
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls1, transform: 'translate(1057.4 915)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls2, transform: 'translate(1048.9 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '10', height: '28' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '9', height: '27' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls3, transform: 'translate(1057.9 929) rotate(-90)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls3, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls2, transform: 'translate(1066.9 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '10', height: '28' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '9', height: '27' })),
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls4, d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1066.591 905.704) rotate(45)' }))));
            case 'deleteRow.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-20039 21352 33 28' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-19105 20303.6) rotate(90)' },
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls1, transform: 'translate(1057.4 915)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls1, transform: 'translate(1066.4 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '10', height: '28' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '9', height: '27' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls2, transform: 'translate(1048.4 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '10', height: '28' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '9', height: '27' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls3, transform: 'translate(1057.4 929) rotate(-90)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls4, d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1066.347 905.704) rotate(45)' }))));
            case 'deleteTable.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-19947 21352 28 28' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-19013 20303.6) rotate(90)' },
                        external__react_["createElement"]("path", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls1), d: 'M.5.5H27.379V9.667H.5Z', transform: 'translate(1048.4 914.917)' }),
                        external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls2), width: '28', height: '27', transform: 'translate(1048.4 906)' }),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1048.4 915)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1048.4 906)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1057.4 915)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1057.4 906)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1066.4 915)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1066.4 906)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1048.4 924)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1057.4 924)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("g", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls3), transform: 'translate(1066.4 924)' },
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls5), width: '10', height: '10' }),
                            external__react_["createElement"]("rect", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls6), x: '0.5', y: '0.5', width: '9', height: '9' })),
                        external__react_["createElement"]("path", { className: Object(office_ui_fabric_react_bundle_["css"])(TableIcons_module_scss.deleteTable, TableIcons_module_scss.cls4), d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1075.091 910.961) rotate(45)' }))));
            case 'insertRowAbove.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-20219 21352 28 28' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-21267.4 20446)' },
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls1, d: 'M.5.5H27.379V9.667H.5Z', transform: 'translate(1048.4 914.917)' }),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls2, transform: 'translate(1048.4 924)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls3, transform: 'translate(1048.4 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls4, d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1064.981 908)' })),
                    ">"));
            case 'insertRowBelow.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-20171 21352 28 28' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-19094.6 22286) rotate(180)' },
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls1, d: 'M.5.5H27.379V9.667H.5Z', transform: 'translate(1048.4 914.917)' }),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls2, transform: 'translate(1048.4 924)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls3, transform: 'translate(1048.4 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls4, d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1064.981 908)' }))));
            case 'insertColumnLeft.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-20128 21352 28 28' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-21034 22428.4) rotate(-90)' },
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls1, d: 'M.5.5H27.379V9.667H.5Z', transform: 'translate(1048.4 914.917)' }),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls1, transform: 'translate(1048.4 924)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls3, transform: 'translate(1048.4 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls4, d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1064.981 908)' }))));
            case 'insertColumnRight.svg':
                return (external__react_["createElement"]("svg", { className: TableIcons_module_scss.themeableSvgTableIcon, xmlns: 'http://www.w3.org/2000/svg', viewBox: '-20085 21352 28 28' },
                    external__react_["createElement"]("g", { id: 'Table', transform: 'translate(-19151 20303.6) rotate(90)' },
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls1, d: 'M.5.5H27.379V9.667H.5Z', transform: 'translate(1048.4 914.917)' }),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls2, transform: 'translate(1048.4 924)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("g", { className: TableIcons_module_scss.cls3, transform: 'translate(1048.4 906)' },
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls5, width: '28', height: '10' }),
                            external__react_["createElement"]("rect", { className: TableIcons_module_scss.cls6, x: '0.5', y: '0.5', width: '27', height: '9' })),
                        external__react_["createElement"]("path", { className: TableIcons_module_scss.cls4, d: 'M-2.837,0V2.748H-5.581v.516h2.744V6h.512V3.264H.419V2.748H-2.325V0Z', transform: 'translate(1064.981 908)' }))));
            case 'simpleTableStyleNeutral.svg':
                return TableIcons.simpleTableIcon(TableIcons_module_scss.neutralPrimaryIcon);
            case 'borderHeaderTableStyleNeutral.svg':
                return TableIcons.borderHeaderTableStyle(TableIcons_module_scss.neutralPrimaryIcon);
            case 'filledHeaderTableStyleNeutral.svg':
                return TableIcons.filledHeaderTableStyle(TableIcons_module_scss.neutralPrimaryIcon);
            case 'bandedRowTableStyleNeutral.svg':
                return TableIcons.bandedRowTableStyle(TableIcons_module_scss.neutralPrimaryIcon, TableIcons_module_scss.neutralRows);
            case 'bandedRowColumnTableStyleNeutral.svg':
                return TableIcons.bandedRowColumnTableStyle(TableIcons_module_scss.neutralPrimaryIcon, TableIcons_module_scss.neutralRows, TableIcons_module_scss.neutralLeftColumn);
            case 'simpleTableStyleTheme.svg':
                return TableIcons.simpleTableIcon(TableIcons_module_scss.themePrimaryIcon);
            case 'borderHeaderTableStyleTheme.svg':
                return TableIcons.borderHeaderTableStyle(TableIcons_module_scss.themePrimaryIcon);
            case 'filledHeaderTableStyleTheme.svg':
                return TableIcons.filledHeaderTableStyle(TableIcons_module_scss.themePrimaryIcon);
            case 'bandedRowTableStyleTheme.svg':
                return TableIcons.bandedRowTableStyle(TableIcons_module_scss.themePrimaryIcon, TableIcons_module_scss.themeRows);
            case 'bandedRowColumnTableStyleTheme.svg':
                return TableIcons.bandedRowColumnTableStyle(TableIcons_module_scss.themePrimaryIcon, TableIcons_module_scss.themeRows, TableIcons_module_scss.themeLeftColumn);
            default:
                return;
        }
    };
    TableIcons.simpleTableIcon = function (className) {
        return (external__react_["createElement"]("svg", { xmlns: 'http://www.w3.org/2000/svg', x: '0px', y: '0px', viewBox: '0 0 48 48', className: TableIcons_module_scss.themeableSvgTableIcon },
            external__react_["createElement"]("g", null,
                external__react_["createElement"]("path", { className: className, 
                    d: 'M0,0v11.5v1v11v1v11v1V48h48V36.5v-1v-11v-1v-11v-1V0H0z M17,35.5v-11h15v11H17z M32,36.5V47H17V36.5H32zM17,23.5v-11h15v11H17z M17,11.5V1h15v10.5H17z M1,1h15v10.5H1V1z M1,12.5h15v11H1V12.5z M1,24.5h15v11H1V24.5z M1,36.5h15V47H1V36.5z M47,47H33V36.5h14V47z M47,35.5H33v-11h14V35.5z M47,23.5H33v-11h14V23.5z M33,11.5V1h14v10.5H33z' }))));
    };
    TableIcons.borderHeaderTableStyle = function (className) {
        return (external__react_["createElement"]("svg", { xmlns: 'http://www.w3.org/2000/svg', x: '0px', y: '0px', viewBox: '0 0 48 48', className: TableIcons_module_scss.themeableSvgTableIcon },
            external__react_["createElement"]("g", null,
                external__react_["createElement"]("path", { className: className, 
                    d: 'M33,0h-1H17h-1H0v12v11.5v1v11v1V48h48V36.5v-1v-11v-1V12V0H33z M47,1v10H33V1H47z M32,1v10H17V1H32z M17,35.5v-11h15v11H17z M32,36.5V47H17V36.5H32z M17,23.5V13h15v10.5H17z M1,1h15v10H1V1z M1,13h15v10.5H1V13z M1,24.5h15v11H1V24.5zM1,36.5h15V47H1V36.5z M47,47H33V36.5h14V47z M47,35.5H33v-11h14V35.5z M33,23.5V13h14v10.5H33z' }))));
    };
    TableIcons.filledHeaderTableStyle = function (className) {
        return (external__react_["createElement"]("svg", { xmlns: 'http://www.w3.org/2000/svg', x: '0px', y: '0px', viewBox: '0 0 48 48', className: TableIcons_module_scss.themeableSvgTableIcon },
            external__react_["createElement"]("g", null,
                external__react_["createElement"]("path", { className: className, 
                    d: 'M48,12V0H0v12h0v11.5v1v11v1V48h48V36.5v-1v-11v-1L48,12L48,12z M17,35.5v-11h15v11H17z M32,36.5V47H17V36.5H32z M17,23.5V13h15v10.5H17z M1,13h15v10.5H1V13z M1,24.5h15v11H1V24.5z M1,36.5h15V47H1V36.5z M47,47H33V36.5h14V47z M47,35.5H33v-11h14V35.5z M33,23.5V13h14v10.5H33z' }))));
    };
    TableIcons.bandedRowTableStyle = function (primaryClassName, rowClassName) {
        return (external__react_["createElement"]("svg", { xmlns: 'http://www.w3.org/2000/svg', x: '0px', y: '0px', viewBox: '0 0 48 48', className: TableIcons_module_scss.themeableSvgTableIcon },
            external__react_["createElement"]("rect", { x: '1', y: '36.5', className: rowClassName, width: '46', height: '11' }),
            external__react_["createElement"]("rect", { x: '1', y: '13', className: rowClassName, width: '46', height: '11' }),
            external__react_["createElement"]("g", null,
                external__react_["createElement"]("path", { className: primaryClassName, 
                    d: 'M48,12V0H0v12h0v11.5v1v11v1V48h48V36.5v-1v-11v-1L48,12L48,12z M17,35.5v-11h15v11H17z M32,36.5V47H17V36.5H32z M17,23.5V13h15v10.5H17z M1,13h15v10.5H1V13z M1,24.5h15v11H1V24.5z M1,36.5h15V47H1V36.5z M47,47H33V36.5h14V47z M47,35.5H33v-11h14V35.5z M33,23.5V13h14v10.5H33z' }))));
    };
    TableIcons.bandedRowColumnTableStyle = function (primaryClassName, rowClassName, columnClassName) {
        return (external__react_["createElement"]("svg", { xmlns: 'http://www.w3.org/2000/svg', x: '0px', y: '0px', viewBox: '0 0 48 48', className: TableIcons_module_scss.themeableSvgTableIcon },
            external__react_["createElement"]("rect", { x: '16.9', y: '36.5', className: rowClassName, width: '30.1', height: '10.5' }),
            external__react_["createElement"]("rect", { x: '17', y: '13', className: rowClassName, width: '30', height: '11' }),
            external__react_["createElement"]("rect", { x: '1', y: '13', className: columnClassName, width: '15', height: '34' }),
            external__react_["createElement"]("g", null,
                external__react_["createElement"]("path", { className: primaryClassName, 
                    d: 'M48,12V0H0v12h0v11.5v1v11v1V48h48V36.5v-1v-11v-1L48,12L48,12z M17,35.5v-11h15v11H17z M32,36.5V47H17V36.5H32z M17,23.5V13h15v10.5H17z M1,13h15v10.5H1V13z M1,24.5h15v11H1V24.5z M1,36.5h15V47H1V36.5z M47,47H33V36.5h14V47z M47,35.5H33v-11h14V35.5z M33,23.5V13h14v10.5H33z' }))));
    };
    return TableIcons;
}());
/* harmony default export */ var tableIcons_TableIcons = (TableIcons_TableIcons);

// EXTERNAL MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-canvas/common/Flights.js
var Flights = __webpack_require__(24);

// EXTERNAL MODULE: C:/agent/1/_work/60/s/common/temp/node_modules/.onedrive.pkgs.visualstudio.com/@ms/odsp-utilities-bundle/1.1.9/node_modules/@ms/odsp-utilities-bundle/lib/index.js
var lib = __webpack_require__(12);
var lib_default = /*#__PURE__*/__webpack_require__.n(lib);

// EXTERNAL MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/rte/BaseRte.js + 4 modules
var BaseRte = __webpack_require__(49);

// CONCATENATED MODULE: C:/agent/1/_work/60/s/libraries/sp-canvas/lib/sp-rte/PropertyPane/SPRtePropertyPaneControl.js














var FontStyle;
(function (FontStyle) {
    FontStyle["NormalText"] = "normalText";
    FontStyle["H2"] = "h2";
    FontStyle["H3"] = "h3";
    FontStyle["H4"] = "h4";
    FontStyle["Blockquote"] = "blockquote";
})(FontStyle || (FontStyle = {}));
var PROPERTY_PATH = new Set(['fontSize', 'fontStyle']);
var TABLE_BUTTON_CLASS = 'tableButtonInPane';
var SPRtePropertyPaneControl_SPRtePropertyPaneControl = (function (_super) {
    tslib_es6["c" /* __extends */](SPRtePropertyPaneControl, _super);
    function SPRtePropertyPaneControl(sprte) {
        var _this = _super.call(this) || this;
        _this._sprte = sprte;
        _this._async = new office_ui_fabric_react_bundle_["Async"](_this);
        _this._platform = new lib["PlatformDetection"]();
        return _this;
    }
    SPRtePropertyPaneControl._getArrayOfElements = function (props, component) {
        return props.map(function (prop) { return external__react_["createElement"](component, prop); });
    };
    SPRtePropertyPaneControl._controlsInOneRow = function (props, component, spaceInBetween) {
        return Object(sp_webpart_base_["PropertyPaneCustomField"])({
            key: 'ControlsInOneRow',
            onRender: function (element) { return external__react_dom_["render"](external__react_["createElement"](PropertyPane_ControlsInOneRow, { spaceInBetween: spaceInBetween }, SPRtePropertyPaneControl._getArrayOfElements(props, component)), element); }
        });
    };
    SPRtePropertyPaneControl._addTableButtonClassName = function (buttons) {
        return buttons.map(function (button) {
            button.className = TABLE_BUTTON_CLASS;
            return button;
        });
    };
    SPRtePropertyPaneControl.prototype._getPropertyPaneData = function () {
        return Promise.resolve({
            webPartId: this._sprte.props.id,
            title: RteStrings_resx["a" /* default */].RTESettingsText,
            configuration: {
                pages: this._getProperyPanePages()
            },
            onPropertyPaneFieldChanged: undefined,
            onConfigurationEvent: undefined,
            onRendered: undefined,
            properties: {},
            isReactive: !this.disableReactivePropertyChanges
        });
    };
    SPRtePropertyPaneControl.prototype._internalDispose = function () {
    };
    SPRtePropertyPaneControl.prototype._onPropertyPaneFieldChanged = function (propertyPath, newValue) {
        var _this = this;
        if (this._platform.isEdge && (PROPERTY_PATH.has(propertyPath))) {
            this._async.setTimeout(function () {
                _this._onFieldChanged(propertyPath, newValue);
            }, 0);
        }
        else {
            this._onFieldChanged(propertyPath, newValue);
        }
    };
    SPRtePropertyPaneControl.prototype._onFieldChanged = function (propertyPath, newValue) {
        if (propertyPath === 'fontSize') {
            this._sprte.getFormatHandler(BaseRte["b" /* Format */].FontSize)({
                format: BaseRte["b" /* Format */].FontSize,
                value: {
                    'fontSize': newValue
                }
            });
        }
        else {
            this._sprte.getFormatHandler(newValue)();
        }
    };
    SPRtePropertyPaneControl.prototype._propertyPaneButton = function (formatKey, automationId, value, imageSrc, disabled, className) {
        var _this = this;
        var button = this._sprte.config.formats.get(formatKey).button;
        return {
            ariaLabel: button.title,
            fabricIconKey: button.fabricIconKey,
            imageSrc: tableIcons_TableIcons.getSvgFromSrc(imageSrc),
            isActive: this._sprte.isFormatActive(formatKey),
            isDisabled: disabled,
            key: button.title,
            onClick: value ?
                function () { return _this._sprte.getFormatHandler(formatKey)(value); } :
                this._sprte.getFormatHandler(formatKey),
            text: (!button.fabricIconKey && !button.imageSrc) ? button.label : undefined,
            title: button.title,
            automationId: automationId
        };
    };
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_firstRowButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].Undo, 'undo'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Redo, 'redo'),
                {
                    ariaLabel: RteStrings_resx["a" /* default */].FormattingBarClearFormattingButtonTitle,
                    fabricIconKey: 'ClearFormatting',
                    key: RteStrings_resx["a" /* default */].FormattingBarClearFormattingButtonTitle,
                    onClick: this._sprte.getFormatHandler(BaseRte["b" /* Format */].ClearFormatting),
                    title: RteStrings_resx["a" /* default */].FormattingBarClearFormattingButtonTitle,
                    automationId: 'clearFormatting-button'
                }
            ];
            return buttons;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_styleCommandButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].Bold, 'bold'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Italic, 'italic'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Underline, 'underline'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Strike, 'strike'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Superscipt, 'superscript'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Subscript, 'subscript')
            ];
            return buttons;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_hyperlinkButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].Link, 'link'),
                this._propertyPaneButton(BaseRte["b" /* Format */].UnLink, 'unlink', undefined , SPRtePropertyPaneControl._noImgSrc, !this._sprte.isLinkSelected)
            ];
            return buttons;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_paragraphButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].AlignLeft, 'alignLeft'),
                this._propertyPaneButton(BaseRte["b" /* Format */].AlignCenter, 'alignCenter'),
                this._propertyPaneButton(BaseRte["b" /* Format */].AlignRight, 'alignRight'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Justify, 'justify'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Indent, 'indent'),
                this._propertyPaneButton(BaseRte["b" /* Format */].Outdent, 'outdent')
            ];
            return buttons;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_bulletButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].Bullet, 'bullet'),
                this._propertyPaneButton(BaseRte["b" /* Format */].List, 'list')
            ];
            return buttons;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_tableButton", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].CustomTable, 'createTable', {
                    format: BaseRte["b" /* Format */].CustomTable,
                    value: {
                        'customTable': {
                            rows: 3,
                            columns: 3
                        }
                    }
                }, SPRtePropertyPaneControl._noImgSrc, this._sprte.isTableContentsSelected)
            ];
            return buttons;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_tableOperationButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].RowInserBefore, 'insertRowAbove', { format: BaseRte["b" /* Format */].RowInserBefore }),
                this._propertyPaneButton(BaseRte["b" /* Format */].RowInsertAfter, 'insertRowBelow', { format: BaseRte["b" /* Format */].RowInsertAfter }),
                this._propertyPaneButton(BaseRte["b" /* Format */].ColumnInsertBefore, 'insertColumnLeft', { format: BaseRte["b" /* Format */].ColumnInsertBefore }),
                this._propertyPaneButton(BaseRte["b" /* Format */].ColumnInsertAfter, 'insertColumnRight', { format: BaseRte["b" /* Format */].ColumnInsertAfter }),
                this._propertyPaneButton(BaseRte["b" /* Format */].RowDelete, 'deleteRow', { format: BaseRte["b" /* Format */].RowDelete }),
                this._propertyPaneButton(BaseRte["b" /* Format */].ColumnDelete, 'deleteColumn', { format: BaseRte["b" /* Format */].ColumnDelete }),
                this._propertyPaneButton(Flights["a" /* default */].isRTEAdvancedTableSettingsFlightEnabled() ? BaseRte["b" /* Format */].TableCustomDelete : BaseRte["b" /* Format */].TableDelete, 'deleteTable', { format: Flights["a" /* default */].isRTEAdvancedTableSettingsFlightEnabled() ? BaseRte["b" /* Format */].TableCustomDelete
                        : BaseRte["b" /* Format */].TableDelete })
            ];
            return SPRtePropertyPaneControl._addTableButtonClassName(buttons);
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_tableStyleNeutralButtons", {
        get: function () {
            return this._tableStyleButtons("Neutral" );
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_tableStyleThemeButtons", {
        get: function () {
            return this._tableStyleButtons("Theme" );
        },
        enumerable: true,
        configurable: true
    });
    SPRtePropertyPaneControl.prototype._tableStyleButtons = function (styleColor) {
        var isTheme = styleColor === "Theme" ;
        var buttons = [
            this._propertyPaneButton(isTheme ? BaseRte["b" /* Format */].SimpleTableStyleTheme : BaseRte["b" /* Format */].SimpleTableStyleNeutral, "simpleTableStyle" + styleColor, {
                format: BaseRte["b" /* Format */].TableStyle,
                value: {
                    'tableStyle': {
                        'tableStyle': "simpleTableStyle" ,
                        'styleColor': styleColor
                    }
                }
            }, "simpleTableStyle" + styleColor + ".svg"),
            this._propertyPaneButton(isTheme ? BaseRte["b" /* Format */].BorderHeaderTableStyleTheme : BaseRte["b" /* Format */].BorderHeaderTableStyleNeutral, "borderHeaderTableStyle" + styleColor, {
                format: BaseRte["b" /* Format */].TableStyle,
                value: {
                    'tableStyle': {
                        'tableStyle': "borderHeaderTableStyle" ,
                        'styleColor': styleColor
                    }
                }
            }, "borderHeaderTableStyle" + styleColor + ".svg"),
            this._propertyPaneButton(isTheme ? BaseRte["b" /* Format */].FilledHeaderTableStyleTheme : BaseRte["b" /* Format */].FilledHeaderTableStyleNeutral, "filledHeaderTableStyle" + styleColor, {
                format: BaseRte["b" /* Format */].TableStyle,
                value: {
                    'tableStyle': {
                        'tableStyle': "filledHeaderTableStyle" ,
                        'styleColor': styleColor
                    }
                }
            }, "filledHeaderTableStyle" + styleColor + ".svg"),
            this._propertyPaneButton(isTheme ? BaseRte["b" /* Format */].BandedRowTableStyleTheme : BaseRte["b" /* Format */].BandedRowTableStyleNeutral, "bandedRowTableStyle" + styleColor, {
                format: BaseRte["b" /* Format */].TableStyle,
                value: {
                    'tableStyle': {
                        'tableStyle': "bandedRowTableStyle" ,
                        'styleColor': styleColor
                    }
                }
            }, "bandedRowTableStyle" + styleColor + ".svg"),
            this._propertyPaneButton(isTheme ? BaseRte["b" /* Format */].BandedRowColumnTableStyleTheme : BaseRte["b" /* Format */].BandedRowColumnTableStyleNeutral, "bandedRowColumnTableStyle" + styleColor, {
                format: BaseRte["b" /* Format */].TableStyle,
                value: {
                    'tableStyle': {
                        'tableStyle': "bandedRowColumnTableStyle" ,
                        'styleColor': styleColor
                    }
                }
            }, "bandedRowColumnTableStyle" + styleColor + ".svg")
        ];
        return buttons;
    };
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_fontStyleDropDown", {
        get: function () {
            var selectedKey = 'normalText';
            for (var i = 0; i < SPRtePropertyPaneControl._fontStyleOptions.length; i++) {
                if (this._sprte.isFormatActive(SPRtePropertyPaneControl._fontStyleOptions[i].key)) {
                    selectedKey = SPRtePropertyPaneControl._fontStyleOptions[i].key;
                    break;
                }
            }
            return Object(sp_webpart_base_["PropertyPaneDropdown"])('fontStyle', {
                label: RteStrings_resx["a" /* default */].FontDropDownText,
                options: SPRtePropertyPaneControl._fontStyleOptions,
                selectedKey: selectedKey
            });
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_fontSizeDropDown", {
        get: function () {
            return Object(sp_webpart_base_["PropertyPaneDropdown"])('fontSize', {
                label: RteStrings_resx["a" /* default */].FontSizeDropDownLabel,
                options: SPRtePropertyPaneControl._fontSizeOptions,
                selectedKey: this._sprte.fontSizeStyle ? this._sprte.fontSizeStyle : ''
            });
        },
        enumerable: true,
        configurable: true
    });
    SPRtePropertyPaneControl.prototype._handleFontColorChange = function (id, color, theme) {
        this._sprte.getFormatHandler(BaseRte["b" /* Format */].FontColor)({
            format: BaseRte["b" /* Format */].FontColor,
            value: {
                'fontColor': PropertyPane_SPRteColorGroups.buildFontColorMap(theme).get(id)
            }
        });
    };
    SPRtePropertyPaneControl.prototype._getColorPickerProps = function (colorPickerGroups, onColorChanged, buttonLabel, previewColor, switchToDefaultColor, defaultButtonLabel, selectedColor, automationId, fillThemeColor) {
        return {
            buttonLabel: buttonLabel,
            colorPickerGroups: colorPickerGroups,
            onColorChanged: onColorChanged,
            previewColor: previewColor,
            key: buttonLabel,
            switchToDefaultColor: switchToDefaultColor,
            defaultButtonLabel: defaultButtonLabel,
            selectedColor: selectedColor,
            fillThemeColor: fillThemeColor,
            automationId: automationId
        };
    };
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_colorPickers", {
        get: function () {
            var colorPickers = [];
            colorPickers.push(this._getColorPickerProps(["themeColors" , "standardColors" ], this._handleFontColorChange, RteStrings_resx["a" /* default */].FontColorLabel, this._sprte.fontColor, this._sprte.removeFontColorFormat, RteStrings_resx["a" /* default */].DefaultColorLabel, this._sprte.fontColor, 'fontColor', true));
            colorPickers.push(this._getColorPickerProps(["highlightColors" ], this._handleHighlightColorChange, RteStrings_resx["a" /* default */].HightlightLabel, this._sprte.highlightColor, this._sprte.removeHighlightColor, RteStrings_resx["a" /* default */].RemoveHighlightColor, this._sprte.highlightColor && this._sprte.highlightColor.toUpperCase(), 'highlightLabel'));
            return colorPickers;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_alignTableButtons", {
        get: function () {
            var buttons = [
                this._propertyPaneButton(BaseRte["b" /* Format */].AlignTableLeft, 'alignTableLeft'),
                this._propertyPaneButton(BaseRte["b" /* Format */].AlignTableCenter, 'alignTableCenter'),
                this._propertyPaneButton(BaseRte["b" /* Format */].AlignTableRight, 'alignTableRight')
            ];
            return SPRtePropertyPaneControl._addTableButtonClassName(buttons);
        },
        enumerable: true,
        configurable: true
    });
    SPRtePropertyPaneControl.prototype._handleHighlightColorChange = function (id, color) {
        this._sprte.getFormatHandler(BaseRte["b" /* Format */].FontColor)({
            format: BaseRte["b" /* Format */].HighlightColor,
            value: {
                'highlightColor': PropertyPane_SPRteColorGroups.HighlightColorMap.get(color)
            }
        });
    };
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_tableGroupLabel", {
        get: function () {
            return Flights["a" /* default */].isRTEAdvancedTableSettingsFlightEnabled() ?
                RteStrings_resx["a" /* default */].TableGroup :
                RteStrings_resx["a" /* default */].TableGroupLabel;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(SPRtePropertyPaneControl.prototype, "_shouldAddTableButtons", {
        get: function () {
            return Flights["a" /* default */].isRTEAdvancedTableSettingsFlightEnabled() && this._sprte.isTableContentsSelected;
        },
        enumerable: true,
        configurable: true
    });
    SPRtePropertyPaneControl.prototype._getProperyPanePages = function () {
        var groups = [
            {
                groupName: '',
                groupFields: [
                    SPRtePropertyPaneControl._controlsInOneRow(this._firstRowButtons, SPRtePropertyPaneButton),
                    this._fontStyleDropDown,
                    this._fontSizeDropDown,
                    SPRtePropertyPaneControl._controlsInOneRow(this._styleCommandButtons, SPRtePropertyPaneButton),
                    SPRtePropertyPaneControl._controlsInOneRow(this._colorPickers, PropertyPane_SPRteColorPicker)
                ]
            },
            {
                groupName: RteStrings_resx["a" /* default */].ParagraphGroupText,
                groupFields: [
                    SPRtePropertyPaneControl._controlsInOneRow(this._paragraphButtons, SPRtePropertyPaneButton),
                    SPRtePropertyPaneControl._controlsInOneRow(this._bulletButtons, SPRtePropertyPaneButton)
                ]
            },
            {
                groupName: RteStrings_resx["a" /* default */].HyperlinkGroupLabel,
                groupFields: [SPRtePropertyPaneControl._controlsInOneRow(this._hyperlinkButtons, SPRtePropertyPaneButton)]
            },
            {
                groupName: this._tableGroupLabel,
                groupFields: [SPRtePropertyPaneControl._controlsInOneRow(this._tableButton, SPRtePropertyPaneButton)]
            }
        ];
        if (this._shouldAddTableButtons) {
            groups.push({
                groupName: RteStrings_resx["a" /* default */].TableStylesGroupLabel,
                groupFields: [
                    SPRtePropertyPaneControl._controlsInOneRow(this._tableStyleNeutralButtons, SPRtePropertyPaneButton),
                    SPRtePropertyPaneControl._controlsInOneRow(this._tableStyleThemeButtons, SPRtePropertyPaneButton)
                ]
            });
        }
        if (this._sprte.isTableContentsSelected) {
            groups.push({
                groupName: RteStrings_resx["a" /* default */].InsertAndDeleteGroupLabel,
                groupFields: [
                    SPRtePropertyPaneControl._controlsInOneRow(this._tableOperationButtons, SPRtePropertyPaneButton)
                ]
            });
        }
        if (this._shouldAddTableButtons) {
            groups.push({
                groupName: RteStrings_resx["a" /* default */].AlignTableGroupLabel,
                groupFields: [
                    SPRtePropertyPaneControl._controlsInOneRow(this._alignTableButtons, SPRtePropertyPaneButton)
                ]
            });
        }
        return [{
                groups: groups
            }];
    };
    SPRtePropertyPaneControl._fontStyleOptions = [
        { key: FontStyle.NormalText, text: RteStrings_resx["a" /* default */].FormattingBarNormalTextButtonTitle },
        { key: FontStyle.H2, text: RteStrings_resx["a" /* default */].FormattingBarHeading2ButtonTitle },
        { key: FontStyle.H3, text: RteStrings_resx["a" /* default */].FormattingBarHeading3ButtonTitle },
        { key: FontStyle.H4, text: RteStrings_resx["a" /* default */].FormattingBarHeading4ButtonTitle },
        { key: FontStyle.Blockquote, text: RteStrings_resx["a" /* default */].FormattingBarQuoteButtonTitle }
    ];
    SPRtePropertyPaneControl._fontSizeOptions = [
        { key: "12" , text: "12" .toString() },
        { key: "14" , text: "14" .toString() },
        { key: "15" , text: "15" .toString() },
        { key: "17" , text: "17" .toString() },
        { key: "21" , text: "21" .toString() },
        { key: "24" , text: "24" .toString() },
        { key: "28" , text: "28" .toString() },
        { key: "32" , text: "32" .toString() },
        { key: "36" , text: "36" .toString() },
        { key: "42" , text: "42" .toString() }
    ];
    tslib_es6["b" /* __decorate */]([
        office_ui_fabric_react_bundle_["autobind"]
    ], SPRtePropertyPaneControl.prototype, "_handleFontColorChange", null);
    tslib_es6["b" /* __decorate */]([
        office_ui_fabric_react_bundle_["autobind"]
    ], SPRtePropertyPaneControl.prototype, "_handleHighlightColorChange", null);
    return SPRtePropertyPaneControl;
}(sp_webpart_base_["BasePropertyPaneConsumer"]));
/* harmony default export */ var PropertyPane_SPRtePropertyPaneControl = __webpack_exports__["default"] = (SPRtePropertyPaneControl_SPRtePropertyPaneControl);


/***/ })

});
//# sourceMappingURL=1.1_c9f3a4c0b2c394334ea0.js.map