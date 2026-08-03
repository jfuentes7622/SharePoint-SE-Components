//import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, SPHttpClientConfiguration } from "@microsoft/sp-http";
//import { ISpfxFetch } from "./../Interfaces/ISpfxFetch";
/* 
export class SpfxFetch implements ISpfxFetch {

    private _spHttpClient: SPHttpClient;

    constructor(spHttpClient: SPHttpClient) {
        this._spHttpClient = spHttpClient;
    }


    public async get(endpoint: string, config: SPHttpClientConfiguration, p0: { headers: { accept: string; }; }): Promise<SPHttpClientResponse> {
        const cors: ISPHttpClientOptions = {
            headers: new Headers(),
            method: "Get",
            mode: "cors"
        };
        const res = await this._spHttpClient.get(endpoint, config, cors);
        return res;
    }
} */