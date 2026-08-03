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
import { Validate } from '@microsoft/sp-core-library';
import SystemJsLoader from './SystemJsLoader';
import { BaseComponentLoader } from '../BaseComponentLoader';
import _loadComponent from './loadComponent';
import ComponentOverrider from './ComponentOverrider';
import ErrorBuilder from '../error/ErrorBuilder';
var SPSystemJsComponentLoader = (function (_super) {
    __extends(SPSystemJsComponentLoader, _super);
    function SPSystemJsComponentLoader() {
        var _this = _super.call(this) || this;
        _this._systemJsLoader = SystemJsLoader.instance;
        return _this;
    }
    SPSystemJsComponentLoader.prototype.loadScript = function (url, options) {
        Validate.isNonemptyString(url, 'url');
        if (typeof options === 'string') {
            throw ErrorBuilder.buildLoadScriptWithStringError();
        }
        var globalMetaConfig = {
            meta: {}
        };
        globalMetaConfig.meta[url] = {
            scriptLoad: false
        };
        if (options) {
            if (options.globalExportsName) {
                globalMetaConfig.meta[url] = {
                    format: 'global',
                    exports: options.globalExportsName
                };
            }
        }
        this._systemJsLoader.systemConfig(globalMetaConfig);
        return this._systemJsLoader.systemImport(url);
    };
    SPSystemJsComponentLoader.prototype.loadComponent = function (manifest) {
        return _loadComponent(manifest);
    };
    SPSystemJsComponentLoader.prototype._overrideComponent = function (componentId, componentModule) {
        ComponentOverrider.overrideComponent(componentId, componentModule);
    };
    SPSystemJsComponentLoader.prototype._listViewHostWorkaround = function (preloadedData) {
        if (preloadedData.clientSideApplicationId === 'b1ab4aaa-f779-405c-8683-d3a750b5d18d') {
            this._systemJsLoader._baseSystemConfig(SystemJsLoader.pluginName,  false);
        }
    };
    return SPSystemJsComponentLoader;
}(BaseComponentLoader));
export default SPSystemJsComponentLoader;
