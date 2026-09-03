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
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var strings = require("GridControlWebPartStrings");
var GridRichTextEditor = (function (_super) {
    __extends(GridRichTextEditor, _super);
    function GridRichTextEditor() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    GridRichTextEditor.prototype.componentDidMount = function () {
        this.updateEditorHtml();
    };
    GridRichTextEditor.prototype.componentDidUpdate = function () {
        this.updateEditorHtml();
    };
    GridRichTextEditor.prototype.updateEditorHtml = function () {
        if (this._editor && this._editor.innerHTML !== String(this.props.value || '')) {
            this._editor.innerHTML = String(this.props.value || '');
        }
    };
    GridRichTextEditor.prototype.applyCommand = function (command) {
        if (!this._editor) {
            return;
        }
        this._editor.focus();
        document.execCommand(command, false, undefined);
        this.props.onChange(this._editor.innerHTML);
    };
    GridRichTextEditor.prototype.handleInput = function () {
        if (this._editor) {
            this.props.onChange(this._editor.innerHTML);
        }
    };
    GridRichTextEditor.prototype.renderToolbarButton = function (command, iconName, label) {
        var _this = this;
        return (React.createElement("button", { type: "button", title: label, "aria-label": label, onMouseDown: function (ev) {
                ev.preventDefault();
                _this.applyCommand(command);
            } },
            React.createElement("i", { className: 'ms-Icon ms-Icon--' + iconName, "aria-hidden": "true" })));
    };
    GridRichTextEditor.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: "gc-rich-text-editor", title: this.props.title },
            React.createElement("div", { className: "gc-rich-text-toolbar", role: "toolbar", "aria-label": strings.RuntimeRichTextToolbar },
                this.renderToolbarButton('bold', 'Bold', strings.RuntimeRichTextBold),
                this.renderToolbarButton('italic', 'Italic', strings.RuntimeRichTextItalic),
                this.renderToolbarButton('underline', 'Underline', strings.RuntimeRichTextUnderline),
                this.renderToolbarButton('insertUnorderedList', 'BulletedList', strings.RuntimeRichTextBulletedList),
                this.renderToolbarButton('insertOrderedList', 'NumberedList', strings.RuntimeRichTextNumberedList),
                this.renderToolbarButton('justifyLeft', 'AlignLeft', strings.RuntimeRichTextAlignLeft),
                this.renderToolbarButton('justifyCenter', 'AlignCenter', strings.RuntimeRichTextAlignCenter),
                this.renderToolbarButton('justifyRight', 'AlignRight', strings.RuntimeRichTextAlignRight),
                this.renderToolbarButton('removeFormat', 'ClearFormatting', strings.RuntimeRichTextClearFormatting)),
            React.createElement("div", { className: "gc-rich-text-content", contentEditable: true, "data-placeholder": this.props.placeholder, ref: function (element) { _this._editor = element; }, onInput: function () { return _this.handleInput(); }, onBlur: function () { return _this.handleInput(); } })));
    };
    return GridRichTextEditor;
}(React.Component));
exports.GridRichTextEditor = GridRichTextEditor;

//# sourceMappingURL=GridRichTextEditor.js.map
