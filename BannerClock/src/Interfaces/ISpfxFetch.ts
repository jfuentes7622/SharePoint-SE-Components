import { SPHttpClientConfiguration, SPHttpClientResponse } from "@microsoft/sp-http";

export interface ISpfxFetch{
    get: (endpoint:string, config?: SPHttpClientConfiguration) => Promise<SPHttpClientResponse>;
}
