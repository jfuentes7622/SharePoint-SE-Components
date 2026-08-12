"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var spfxWindow = window;
spfxWindow.DEBUG = true;
spfxWindow.ENVIRONMENTTYPE = 'Local';
// This is a hack to work around a react-redux issue.
spfxWindow.process = { env: { NODE_ENV: 'production' } };
var manifests = spfxWindow.debugManifests.getManifests();
var manifestIds = [];
for (var _i = 0, manifests_1 = manifests; _i < manifests_1.length; _i++) {
    var manifest = manifests_1[_i];
    manifestIds.push(manifest.id.toUpperCase());
}
for (var _a = 0, _b = (spfxWindow.syntheticManifests || []); _a < _b.length; _a++) {
    var syntheticManifest = _b[_a];
    var manifestAlreadyExists = manifestIds.indexOf(syntheticManifest.id.toUpperCase()) !== -1;
    if (!manifestAlreadyExists) {
        manifests.push(syntheticManifest);
    }
}
spfxWindow.preloadedData = {
    'clientSideApplicationId': '8be81a5c-af38-4bb2-af97-afa3b64dfbed',
    manifests: manifests,
    /* tslint:enable:no-any */
    spPageContextInfo: {
        listBaseTemplate: 119,
        cdnPrefix: 'test.sharepoint',
        CorrelationId: '00000000-0000-4000-b000-000000000000',
        currentCultureName: 'en-US',
        currentUICultureName: 'en-US',
        groupId: 0,
        groupType: 'Public',
        isAnonymousGuestUser: false,
        isAppWeb: true,
        isExternalGuestUser: false,
        isNoScriptEnabled: false,
        listId: '00000000-0000-4000-b000-000000000000',
        listPermsMask: { High: 0x7FFFFFFF, Low: 0xFFFFFFFF },
        listUrl: 'https://wwww.contoso.com/sites/workbench/lists/todo',
        listTitle: 'Todo List',
        pageItemId: -1,
        pagePermsMask: { High: 0x7FFFFFFF, Low: 0xFFFFFFFF },
        RecycleBinItemCount: -1,
        serverRequestPath: '/workbencch.aspx',
        siteAbsoluteUrl: 'https://wwww.contoso.com/',
        siteId: '00000000-0000-4000-b000-111111111111',
        siteClassification: 'ABC',
        sitePagesEnabled: true,
        sitePagesFeatureVersion: 4,
        socialBarEnabled: true,
        siteServerRelativeUrl: '/',
        userDisplayName: 'User 1',
        userEmail: 'user1@contoso.com',
        userLoginName: 'user1@contoso.com',
        webAbsoluteUrl: 'https://wwww.contoso.com/sites/workbench',
        webId: '00000000-0000-4000-b000-222222222222',
        webLanguage: 'en-US',
        webLogoUrl: 'https://wwww.contoso.com/sites/workbench/test.jpg',
        webPermMasks: { High: 0x7FFFFFFF, Low: 0xFFFFFFFF },
        webServerRelativeUrl: '/sites/workbench',
        webTemplate: 'Blog',
        webTitle: 'Local Workbench'
    },
    contextWebInfo: {
        FormDigestTimeoutSeconds: 1800,
        FormDigestValue: 'Mock Digest'
    }
};

//# sourceMappingURL=workbenchInit.js.map
