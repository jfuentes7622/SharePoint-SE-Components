import { IClientSideComponentManifest } from '@microsoft/sp-module-interfaces';
import { BaseComponentLoader } from '../BaseComponentLoader';
import { ILoadScriptOptions } from '../ILoadScriptOptions';
import { ISPComponentLoader } from '../ISPComponentLoader';
import SystemJsFallbackLoader from './SystemJsFallbackLoader';
/**
 * The component loader for SPFx.
 * This implementation uses RequireJS as the internal module loader.
 *
 * @alpha
 */
export default class SPRequireJsComponentLoader extends BaseComponentLoader implements ISPComponentLoader {
    private _requireJsLoader;
    private _systemJsFallbackLoader;
    private _loadComponentExecutor;
    /**
     * @internal
     */
    constructor(systemJsFallbackLoader: SystemJsFallbackLoader);
    loadScript<TModule>(url: string, options?: ILoadScriptOptions | undefined): Promise<TModule>;
    loadComponent<TComponent>(manifest: IClientSideComponentManifest): Promise<TComponent>;
    protected _overrideComponent<TComponent>(componentId: string, componentModule: TComponent): void;
    private _shouldSideLoadSystemJs();
    /**
     * This functions tries to load the component with RequireJS, and if the killswitch is not enabled and the load fails,
     * will set up the component to be loaded with SystemJS.
     */
    private _loadComponentWithExecutor<TComponent>(manifest);
    private _buildQosExtraData(manifest);
    private _handleLoadComponentError(error, manifest, qosMonitor, qosExtraData);
}
