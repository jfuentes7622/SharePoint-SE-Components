export interface ILibraryFolder {
    path: string;
    name: string;
}
export interface ILibraryCondition {
    field: string;
    operator?: string;
    logical?: string;
    valueType?: string;
    value: any;
}
export interface ILibraryLoadOptions {
    webUrl: string;
    listName: string;
    viewXml: string;
    fields: string[];
    conditions: ILibraryCondition[];
    time: number;
    batchSize: number;
    folder?: string;
    get: (url: string) => Promise<any>;
    post: (url: string, body: any) => Promise<any>;
    extract: (data: any) => {
        rows: any[];
        fields: any[];
        nextHref: string;
    };
    current: () => boolean;
}
export declare function validateLibraryFilterJson(sources: string[]): void;
export declare function resolveNowExpression(expression: string, time: number): string;
export declare function documentUrl(webUrl: string, path: string): string;
export declare function folderDocuments(current: string, documents: any[], recursive?: boolean): any[];
export declare function folderDocumentCount(current: string, documents: any[]): number;
export declare function ancestorPaths(rootPath: string, rows: any[]): string[];
export declare function folderLevel(root: string, current: string, documents: any[], folders: ILibraryFolder[], prune?: boolean): any[];
export declare function nearestFolder(root: string, current: string, paths: string[]): string;
export declare function recursiveView(viewXml: string, fields: string[], conditions: ILibraryCondition[], time: number, batchSize: number, dateFields?: string[], recursive?: boolean): string;
export declare function detectDocumentLibrary(webUrl: string, listName: string, get: (url: string) => Promise<any>): Promise<boolean>;
export declare function loadLibraryFolders(webUrl: string, listName: string, root: string, documents: any[], get: (url: string) => Promise<any>, current: () => boolean, batchSize: number): Promise<ILibraryFolder[]>;
export declare function defaultLibraryViewId(webUrl: string, listName: string, get: (url: string) => Promise<any>): Promise<string>;
export declare function loadLibraryDocuments(options: ILibraryLoadOptions): Promise<{
    root: string;
    folder: string;
    flat: boolean;
    rows: any[];
    fields: any[];
    metadata: any[];
}>;
export declare function loadChildFolders(webUrl: string, root: string, folder: string, get: (url: string) => Promise<any>, current: () => boolean, batchSize: number): Promise<ILibraryFolder[]>;
export declare function loadAncestorFolders(webUrl: string, root: string, rows: any[], get: (url: string) => Promise<any>): Promise<ILibraryFolder[]>;
