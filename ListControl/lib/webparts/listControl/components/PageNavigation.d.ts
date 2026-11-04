export interface IPageNavigationTarget {
    instanceId: string;
    title: string;
    element: HTMLElement;
}
export declare function getPageNavigationTargets(ownerDocument: Document, sourceInstanceId: string): IPageNavigationTarget[];
export declare function navigateToPageTarget(ownerDocument: Document, targetInstanceId: string, sourceInstanceId: string): void;
