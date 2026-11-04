import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
export interface IInstanceNameProperties {
    instanceName?: string;
    instanceNameGeneratedFor?: string;
}
export declare function ensureInstanceName(properties: IInstanceNameProperties, prefix: string, context: WebPartContext, element: HTMLElement, getDisplayMode: () => DisplayMode): Promise<boolean>;
