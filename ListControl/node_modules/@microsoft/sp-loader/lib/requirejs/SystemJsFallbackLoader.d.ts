import { IClientSideComponentManifest } from '@microsoft/sp-module-interfaces';
import SPRequireJsComponentLoader from './SPRequireJsComponentLoader';
import LoadComponentExecutor from '../utilities/LoadComponentExecutor';
export default class SystemJsFallbackLoader {
    private static _window;
    private _systemJsComponentLoader;
    private _systemJsLoader;
    private _requireJsComponentLoader;
    private _isInitialized;
    private _executor;
    private define;
    private require;
    private requirejs;
    constructor();
    readonly executor: LoadComponentExecutor;
    setRequireJsComponentLoader(requireJsComponentLoader: SPRequireJsComponentLoader): void;
    loadComponent<T>(manifest: IClientSideComponentManifest): Promise<T>;
    private _loadComponentImpl<T>(manifest);
    private _ensureInitialized();
    private _saveGlobals();
    private _restoreGlobals();
}
