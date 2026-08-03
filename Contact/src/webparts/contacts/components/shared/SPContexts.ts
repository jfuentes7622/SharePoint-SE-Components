import { SPHttpClient } from '@microsoft/sp-http';

export type SPContexts = {
    spHttpClient: SPHttpClient;
    absUrl: string;
    relUrl: string;
};