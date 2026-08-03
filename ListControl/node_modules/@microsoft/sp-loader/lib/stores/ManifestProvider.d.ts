import { IClientSideComponentManifest } from '@microsoft/sp-module-interfaces';
import { ServiceScope } from '@microsoft/sp-core-library';
export default class ManifestProvider {
    private static _restApiUrl;
    private _logSource;
    private _pageContext;
    private _httpClient;
    constructor(serviceScope: ServiceScope);
    /**
     * Given a component id and version, requests its manifest (and all its dependencies) to SharePoint
     * through a REST API.
     * Returns a promise with all the manifests.
     * @param componentId - Id of the requested component
     * @param version - Optional. Version of the requested component
     */
    tryGetManifest(componentId: string, version?: string): Promise<IClientSideComponentManifest[]>;
    private _buildRequest(componentId, version?);
    private _extractManifests(response);
}
