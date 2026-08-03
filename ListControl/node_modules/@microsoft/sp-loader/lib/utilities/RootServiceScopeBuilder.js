import { ServiceScope, Environment, EnvironmentType, Validate } from '@microsoft/sp-core-library';
import { _DynamicDataManager } from '@microsoft/sp-dynamic-data';
import { PageContext } from '@microsoft/sp-page-context';
import { HttpClient, SPHttpClient, _GraphHttpClientContext, AadTokenProvider, DigestCache, _AadTokenProviders, _AadConstants } from '@microsoft/sp-http';
import { _logSourceServiceKey, _LogSource, _TraceLogger } from '@microsoft/sp-diagnostics';
var RootServiceScopeBuilder = (function () {
    function RootServiceScopeBuilder() {
    }
    RootServiceScopeBuilder.build = function (preloadedData) {
        var serviceScope = ServiceScope.startNewRoot();
        serviceScope.provide(_logSourceServiceKey, RootServiceScopeBuilder._logSource);
        serviceScope.createDefaultAndProvide(PageContext.serviceKey);
        serviceScope.createDefaultAndProvide(HttpClient.serviceKey);
        serviceScope.createDefaultAndProvide(SPHttpClient.serviceKey);
        serviceScope.createDefaultAndProvide(_DynamicDataManager.serviceKey);
        var _graphContext = serviceScope.createDefaultAndProvide(_GraphHttpClientContext.serviceKey);
        var digestCache = serviceScope.createDefaultAndProvide(DigestCache.serviceKey);
        serviceScope.finish();
        this._initializeGraphHttpClient(_graphContext, preloadedData);
        this._initializeDigestCache(digestCache, preloadedData);
        return serviceScope;
    };
    RootServiceScopeBuilder._initializeGraphHttpClient = function (graphContext, preloadedData) {
        if (DATACENTER && Environment.type !== EnvironmentType.Local) {
            graphContext.initialize(preloadedData.spPageContextInfo.webServerRelativeUrl, preloadedData.spPageContextInfo.msGraphEndpointUrl);
            var redirectPartialUri = window.location.origin + '/_forms/';
            try {
                _AadTokenProviders._initialize(new AadTokenProvider({
                    aadInstanceUrl: preloadedData.spPageContextInfo.aadInstanceUrl,
                    aadTenantId: preloadedData.spPageContextInfo.aadTenantId,
                    aadUserId: preloadedData.spPageContextInfo.aadUserId,
                    popupRedirectUri: redirectPartialUri + _AadConstants.SPFX_SINGLE_SIGN_ON_REPLY_URL,
                    redirectUri: redirectPartialUri +
                        _AadConstants.SPFX_SINGLE_SIGN_ON_REPLY_URL +
                        '?' +
                        _AadConstants.REDIRECT_QUERY_PARAM,
                    servicePrincipalId: _AadConstants.SPO_CE_WEB_APP_PRINCIPAL_ID
                }));
            }
            catch (e) {
                _TraceLogger.logVerbose(this._logSource, 'AadTokenProviders: Failed to initialize');
            }
        }
    };
    RootServiceScopeBuilder._initializeDigestCache = function (digestCache, preloadedData) {
        Validate.isNotNullOrUndefined(preloadedData, 'preloadedData');
        Validate.isNotNullOrUndefined(preloadedData.contextWebInfo, 'preloadedData.contextWebInfo');
        Validate.isNotNullOrUndefined(preloadedData.spPageContextInfo, 'preloadedData.spPageContextInfo');
        _TraceLogger.logVerbose(this._logSource, 'ServiceScopeBuilder: Added preloaded FormDigestValue to cache');
        if (preloadedData.contextWebInfo && preloadedData.spPageContextInfo) {
            var expirationTimestamp = (1000 * preloadedData.contextWebInfo.FormDigestTimeoutSeconds) -
                this.PRELOAD_DIGEST_EXPIRATION_SLOP_MS;
            digestCache.addDigestToCache(preloadedData.spPageContextInfo.webServerRelativeUrl, preloadedData.contextWebInfo.FormDigestValue, expirationTimestamp);
            digestCache.addDigestToCache(preloadedData.spPageContextInfo.webAbsoluteUrl, preloadedData.contextWebInfo.FormDigestValue, expirationTimestamp);
        }
    };
    RootServiceScopeBuilder._logSource = _LogSource.create('RootServiceScope');
    RootServiceScopeBuilder.PRELOAD_DIGEST_EXPIRATION_SLOP_MS = 30000; 
    return RootServiceScopeBuilder;
}());
export default RootServiceScopeBuilder;
