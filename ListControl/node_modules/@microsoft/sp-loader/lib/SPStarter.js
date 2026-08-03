import { Environment, EnvironmentType, Log, Session, UrlQueryParameterCollection, _SPFlight, _SPKillSwitch, Validate, Guid } from '@microsoft/sp-core-library';
import { _SPTelemetryLogHandler, _Telemetry, _PerformanceLogger, _QosMonitor } from '@ms/sp-telemetry';
import { _Diagnostics } from '@microsoft/sp-diagnostics';
import * as developerToolsLoader from './DeveloperTools/DeveloperToolsLoader';
import SPSystemJsComponentLoader from './systemjs/SPSystemJsComponentLoader';
import SPRequireJsComponentLoader from './requirejs/SPRequireJsComponentLoader';
import LocaleStore from './stores/LocaleStore';
import BrowserSupport from './utilities/BrowserSupport';
import { showBrowserDevToolsWarning } from './DeveloperTools/BrowserDeveloperToolsWarning/showBrowserDevToolsWarning';
import { DebugManager } from './debug/DebugManager';
import * as componentConstants from './utilities/componentConstants';
import strings from './SPLoader.resx';
import SystemJsFallbackLoader from './requirejs/SystemJsFallbackLoader';
if (!DEBUG && !DEPRECATED_UNIT_TEST) {
    showBrowserDevToolsWarning();
}
var REQUIREJS_FLIGHT_ID = 125;
var DEBUG_FLIGHTS_QUERY_PARAM = 'debugFlights';
var DEBUG_KILLSWITCHES_QUERY_PARAM = 'debugKillSwitches';
var COMPONENT_LOADER_GLOBAL_VARIABLE = '_spComponentLoader';
var startQosScenarioName = 'SPComponentLoader.start';
var SPStarter = (function () {
    function SPStarter() {
    }
    SPStarter._setBundledComponents = function (bundledComponents) {
        if (!SPStarter._bundledComponents) {
            SPStarter._bundledComponents = bundledComponents;
        }
        else {
            throw new Error('Bundled components can be only set once'); 
        }
    };
    SPStarter.start = function (preloadedData, handleFailure, debugData) {
        var _this = this;
        _PerformanceLogger.devMark('SPStarter.start');
        if (!preloadedData || !preloadedData.spPageContextInfo) {
            var error = new Error(strings.invalidPreloadedDataError);
            handleFailure({
                error: error,
                message: strings.loaderUserFriendlyError,
                operationName: 'InvalidPreloadedData'
            });
            return Promise.reject(error);
        }
        var correlationId = preloadedData.spPageContextInfo.CorrelationId;
        var browserCompatibility = SPStarter.getBrowserCompatibility();
        if (browserCompatibility.supportLevel === 3 ) {
            return this.handleError(new Error(browserCompatibility.warning), strings.browserNotSupportedError, 'BrowserNotSupported', correlationId, preloadedData);
        }
        this._initializeFlightsAndKillswitches(preloadedData);
        var spLoader = require('./index');
        var componentLoader = this._useRequireJs(preloadedData)
            ? new SPRequireJsComponentLoader(new SystemJsFallbackLoader())
            : new SPSystemJsComponentLoader();
        var qosMonitor = undefined;
        try {
            Log._initialize(new _SPTelemetryLogHandler());
            SPStarter._initializeEnvironment(preloadedData);
            Session._initialize({ applicationId: preloadedData.spPageContextInfo.CorrelationId });
            SPStarter._initializeTelemetry(preloadedData);
            _PerformanceLogger.markSpLoaderStart();
            if (_SPKillSwitch.isActivated(SPStarter.enableDeferredLogging, '11/27/2017', 'enableCircularBufferLogging')) {
                SPStarter._logDataInRealTime();
            }
            else {
                window.setTimeout(function () {
                    SPStarter._logDataInRealTime();
                }, SPStarter._realTimeProcessingWaitTime);
            }
            qosMonitor = new _QosMonitor(startQosScenarioName);
            LocaleStore.setLocale(preloadedData.spPageContextInfo.currentUICultureName);
            componentLoader._initialize(preloadedData, SPStarter._getBundledComponents(), !!(debugData || {}).debugLoader);
            spLoader.SPComponentLoader._initialize(componentLoader);
            spLoader.DeveloperTools.initialize(developerToolsLoader);
            window[COMPONENT_LOADER_GLOBAL_VARIABLE] = spLoader.SPComponentLoader;
        }
        catch (error) {
            return Promise.reject(error);
        }
        return DebugManager.initialize(componentLoader, debugData).then(function (debugLoadResult) {
            if (debugLoadResult.debugLoader) {
                return debugLoadResult.debugLoader.start(preloadedData, handleFailure, debugLoadResult);
            }
            if (preloadedData.clientSideApplicationId) {
                return spLoader.SPComponentLoader._startApplication(preloadedData).then(function (application) {
                    qosMonitor.writeSuccess();
                    return application;
                }).catch(function (error) {
                    qosMonitor.writeUnexpectedFailure(undefined, error);
                    return _this.handleError(error, strings.loaderUserFriendlyError, 'StartApplication', correlationId, preloadedData);
                });
            }
            else {
                qosMonitor.writeSuccess();
                return Promise.resolve();
            }
        }).then(function (application) {
            SPStarter._logDataInRealTime();
            return application;
        }).catch(function (error) {
            if (qosMonitor) {
                qosMonitor.writeUnexpectedFailure(undefined, error);
            }
            SPStarter._logDataInRealTime();
            return _this.handleError(error, strings.loaderUserFriendlyError, _Telemetry.isInitialized ? 'ClientError' : 'ClientErrorBeforeTelemetry', correlationId, preloadedData);
        });
    };
    SPStarter.getBrowserCompatibility = function () {
        return BrowserSupport.getBrowserCompatibility();
    };
    SPStarter._logDataInRealTime = function () {
        if (!SPStarter._isTelemetryDisabled() && !SPStarter._isTelemetryLoggingInRealTime) {
            _Telemetry.startRealTimeProcessing();
            SPStarter._isTelemetryLoggingInRealTime = true;
        }
    };
    SPStarter._getBundledComponents = function () {
        Validate.isNotNullOrUndefined(SPStarter._bundledComponents, 'bundledComponents');
        return SPStarter._bundledComponents;
    };
    SPStarter._initializeEnvironment = function (preloadedData) {
        var type;
        if (window.location.hostname === 'localhost' ||
            window['ENVIRONMENTTYPE'] &&
                window['ENVIRONMENTTYPE'] === 'Local') {
            type = EnvironmentType.Local;
        }
        else if (preloadedData.clientSideApplicationId === 'eb4b666b-5c29-4dad-9a99-23613f21a2b7') {
            type = EnvironmentType.ClassicSharePoint;
        }
        else {
            type = EnvironmentType.SharePoint;
        }
        Environment._initialize({ type: type });
    };
    SPStarter._initializeTelemetry = function (preloadedData) {
        if (SPStarter._isTelemetryDisabled()) {
            return;
        }
        var pageContextInfo = preloadedData.spPageContextInfo;
        _Telemetry.initialize(preloadedData.manifests[0].alias, {
            environment: pageContextInfo.env,
            farmLabel: pageContextInfo.farmLabel,
            clientSideApplicationId: preloadedData.clientSideApplicationId || '',
            siteSubscriptionId: pageContextInfo.siteSubscriptionId,
            version: !!BUILD_NUMBER ? BUILD_NUMBER : preloadedData.manifests[0].version,
            loginName: pageContextInfo.userLoginName,
            systemUserKey: pageContextInfo.systemUserKey,
            currentUICultureName: pageContextInfo.currentUICultureName,
            correlationId: pageContextInfo.CorrelationId,
            enableConsoleLog: SPStarter._isConsoleLogEnabled(),
            siteId: pageContextInfo.siteId,
            webId: pageContextInfo.webId
        });
        _Diagnostics.initialize({
            enableConsoleLog: SPStarter._isConsoleLogEnabled(),
            siteId: pageContextInfo.siteId,
            webId: pageContextInfo.webId
        });
    };
    SPStarter._initializeFlightsAndKillswitches = function (preloadedData) {
        _SPFlight.initialize(preloadedData.spPageContextInfo.ExpFeatures);
        _SPKillSwitch.initialize(preloadedData.spPageContextInfo.killSwitches);
        var allowDebugQueryParameter = _SPFlight.isDebugFlightEnabled
            || Environment.type === EnvironmentType.Local;
        if (allowDebugQueryParameter) {
            var queryParams = new UrlQueryParameterCollection(window.location.href);
            _SPFlight.setDebugFlights(queryParams.getValue(DEBUG_FLIGHTS_QUERY_PARAM));
            _SPKillSwitch.setDebugKillswitches(queryParams.getValue(DEBUG_KILLSWITCHES_QUERY_PARAM));
        }
    };
    SPStarter._isQueryParameterTrue = function (name) {
        var parameterValue;
        try {
            var queryParams = new UrlQueryParameterCollection(window.location.href);
            parameterValue = queryParams.getValue(name) === 'true';
        }
        catch (error) {
            parameterValue = false;
        }
        return parameterValue;
    };
    SPStarter._isConsoleLogEnabled = function () {
        return SPStarter._isQueryParameterTrue('enableConsoleLog');
    };
    SPStarter._isTelemetryDisabled = function () {
        return SPStarter._isQueryParameterTrue('disableTelemetry');
    };
    SPStarter._useRequireJs = function (preloadedData) {
        var isFlightEnabled = _SPFlight.isEnabled(REQUIREJS_FLIGHT_ID);
        var isListViewApplication = preloadedData.clientSideApplicationId === componentConstants.listViewHostComponentId;
        return isFlightEnabled || isListViewApplication;
    };
    SPStarter.handleError = function (error, userFriendlyMessage, operationName, correlationId, preloadedData) {
        var errorInformation = {
            message: userFriendlyMessage,
            correlationId: correlationId,
            error: error,
            operationName: operationName
        };
        if (this._isRedirectDisabled(preloadedData)) {
            this._consoleErrorHandleFailure(errorInformation);
        }
        else {
            this._errorAspxHandleFailure(errorInformation);
        }
        return Promise.reject(error);
    };
    SPStarter._isRedirectDisabled = function (preloadedData) {
        return SPStarter._isQueryParameterTrue('noredir') ||
            Environment.type === EnvironmentType.ClassicSharePoint ||
            Environment.type === EnvironmentType.Local ||
            preloadedData.clientSideApplicationId === componentConstants.listViewHostComponentId; 
    };
    SPStarter._consoleErrorHandleFailure = function (errorInformation) {
        console.error(errorInformation.message);
        if (errorInformation.correlationId) {
            console.error("Correlation Id: " + errorInformation.correlationId);
        }
        if (errorInformation.operationName) {
            console.error("Operation name: " + errorInformation.operationName);
        }
        if (errorInformation.error) {
            console.error(errorInformation.error.message);
            console.error("CALL STACK: " + errorInformation.error.stack);
        }
    };
    SPStarter._errorAspxHandleFailure = function (errorInformation) {
        window.location.href =
            '_layouts/15/error.aspx' +
                '?ErrorCorrelationId=' + encodeURIComponent(errorInformation.correlationId || '') +
                '&ErrorText=' + encodeURIComponent(errorInformation.message) +
                '&ErrorDetails=' + encodeURIComponent(errorInformation.error ? errorInformation.error.toString() : '') +
                '&Name=' + encodeURIComponent(errorInformation.operationName || 'DefaultOperation') +
                '&ErrorCategory=spfx';
    };
    SPStarter._bundledComponents = undefined;
    SPStarter._isTelemetryLoggingInRealTime = false;
    SPStarter._realTimeProcessingWaitTime = 1 * 1000; 
    SPStarter.enableDeferredLogging = Guid.parse('4b6e1a58-3e0c-43e2-b9cb-1a470f6fe402');
    return SPStarter;
}());
export default SPStarter;
