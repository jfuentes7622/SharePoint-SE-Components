import { Validate, Text } from '@microsoft/sp-core-library';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { _QosMonitor } from '@ms/sp-telemetry';
import { _TraceLogger } from '@microsoft/sp-diagnostics';
import SystemJsLoader from './SystemJsLoader';
import ManifestStore from '../stores/ManifestStore';
import ComponentStore from '../stores/ComponentStore';
import * as telemetryConstants from '../utilities/telemetryConstants';
import normalizeName, { normalizeFailoverPathName } from './normalizeName';
import { resolvePath } from '../utilities/resolveAddress';
import ResourceUrlChecker from '../utilities/ResourceUrlChecker';
import ErrorBuilder from '../error/ErrorBuilder';
import SPLoaderError from '../error/SPLoaderError';
import strings from '../SPLoader.resx';
var FIRST_RETRY = 1;
var MAX_NUMBER_RETRIES = 3;
var _systemJsLoader = SystemJsLoader.instance;
var loadComponentImplEventName = 'loadComponentImpl';
export default function loadComponent(manifest) {
    Validate.isNotNullOrUndefined(manifest, 'manifest');
    var cachedModule = ComponentStore.instance.tryGetComponent(manifest.id, manifest.version);
    if (cachedModule) {
        return cachedModule;
    }
    var qosMonitor = new _QosMonitor(telemetryConstants.loadComponentQosScenarioName);
    var qosExtraData = _buildQosExtraData(manifest);
    var componentPromise = _loadComponentRetryStrategy(manifest, FIRST_RETRY, MAX_NUMBER_RETRIES)
        .then(function (component) {
        qosMonitor.writeSuccess(qosExtraData);
        return component;
    })
        .catch(function (error) {
        if (error instanceof SPLoaderError && error.isExpected) {
            qosMonitor.writeExpectedFailure(undefined, error, qosExtraData);
        }
        else {
            qosMonitor.writeUnexpectedFailure(undefined, error, qosExtraData);
        }
        ComponentStore.instance.deleteComponent(manifest.id, manifest.version);
        throw error;
    });
    ComponentStore.instance.storeComponent(manifest.id, manifest.version, componentPromise);
    return componentPromise;
}
function _loadComponentRetryStrategy(manifest, currentRetryNumber, maxNumberRetries) {
    if (currentRetryNumber === 1) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, Text.format(strings.loadComponentLog, manifest.id, manifest.alias, manifest.version));
    }
    else {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, Text.format(strings.loadComponentRetryLog, manifest.id, manifest.alias, currentRetryNumber, maxNumberRetries));
    }
    return _loadComponentImpl(manifest)
        .then(function (component) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, Text.format(strings.loadComponentEndLog, manifest.id, manifest.alias, manifest.version));
        return component;
    })
        .catch(function (error) {
        _systemJsLoader.systemDelete(manifest);
        if (currentRetryNumber < maxNumberRetries) {
            return _loadComponentRetryStrategy(manifest, currentRetryNumber + 1, maxNumberRetries);
        }
        else {
            _TraceLogger.logError(telemetryConstants.loadComponentLogSource, new Error(Text.format(strings.loadComponentMaxRetriesError, manifest.id, manifest.alias, maxNumberRetries)));
            throw error;
        }
    });
}
function _loadComponentImpl(manifest) {
    try {
        _systemJsLoader.configure(manifest);
    }
    catch (error) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, error.message, loadComponentImplEventName);
        return Promise.reject(error);
    }
    var componentDeps = [];
    var pathDeps = [];
    try {
        componentDeps = _loadComponentDependencies(manifest);
    }
    catch (error) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, error.message, loadComponentImplEventName);
        return Promise.reject(error);
    }
    try {
        pathDeps = _loadPathDependencies(manifest);
    }
    catch (error) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, error.message, loadComponentImplEventName);
        return Promise.reject(error);
    }
    return Promise.all(componentDeps.concat(pathDeps)).then(function (components) {
        if (!manifest.loaderConfig.entryModuleId) {
            _systemJsLoader.ensure(normalizeName(manifest), {});
            return {};
        }
        return _loadEntryPoint(manifest).then(function (entryPoint) {
            _validateComponentIsNotEmptyOrThrow(entryPoint, manifest);
            return entryPoint;
        });
    }).catch(function (e) {
        throw ErrorBuilder.buildLoadComponentError(manifest, e);
    });
}
function _validateComponentIsNotEmptyOrThrow(component, manifest) {
    if (isEmpty(component)) {
        throw ErrorBuilder.buildLoadComponentReturnsEmptyError(manifest);
    }
    var defaultObject = component.default; 
    if (defaultObject && isEmpty(defaultObject) && !defaultObject.prototype) {
        throw ErrorBuilder.buildLoadComponentReturnsDefaultEmptyError(manifest);
    }
}
function _loadComponentDependencies(manifest) {
    var depPromises = [];
    var resources = manifest.loaderConfig.scriptResources;
    var _loop_1 = function (name_1) {
        if (resources[name_1].type === 'component' && !resources[name_1].shouldNotPreload) {
            var moduleConfiguration_1 = resources[name_1];
            var resourceManifest = ManifestStore.instance.tryGetManifest(moduleConfiguration_1.id, moduleConfiguration_1.version);
            if (resourceManifest) {
                var dep = loadComponent(resourceManifest).catch(function (e) {
                    throw ErrorBuilder.buildLoadComponentDependencyError(manifest, e);
                });
                depPromises.push(dep);
            }
            else {
                if (moduleConfiguration_1.failoverPath) {
                    var dep = _systemJsLoader.systemImport(normalizeFailoverPathName(name_1))
                        .catch(function (e) {
                        return _processSystemImportErrors(manifest, name_1, [ResourceUrlChecker.checkResourceUrl], function () { return ErrorBuilder.buildLoadComponentDependencyFailoverPathError(manifest, name_1, resolvePath(moduleConfiguration_1.failoverPath), e); });
                    });
                    depPromises.push(dep);
                }
                else {
                    depPromises.push(Promise.reject(ErrorBuilder.buildManifestNotFoundError(moduleConfiguration_1)));
                }
            }
        }
    };
    for (var name_1 in resources) {
        _loop_1(name_1);
    }
    return depPromises;
}
function _loadPathDependencies(manifest) {
    var resources = manifest.loaderConfig.scriptResources;
    var loadedPathDependencies = new Map();
    for (var name_2 in resources) {
        if ((resources[name_2].type === 'path' || resources[name_2].type === 'localizedPath')
            && !resources[name_2].shouldNotPreload) {
            if (name_2 !== manifest.loaderConfig.entryModuleId) {
                _loadPathDependency(manifest, name_2, loadedPathDependencies);
            }
        }
    }
    var loadedPathDependenciesValues = [];
    loadedPathDependencies.forEach(function (value) {
        loadedPathDependenciesValues.push(value);
    });
    return loadedPathDependenciesValues;
}
function _loadPathDependency(manifest, name, loadedPathDependencies) {
    var loadedPathDependency = loadedPathDependencies.get(name);
    if (loadedPathDependency) {
        return loadedPathDependency;
    }
    var qosMonitor = new _QosMonitor(telemetryConstants.loadPathDependencyQosScenarioName);
    var qosExtraData = {
        name: name,
        manifestId: manifest.id,
        version: manifest.version,
        alias: manifest.alias,
        isInternal: manifest.isInternal
    };
    _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, Text.format(strings.loadPathDependencyLog, name, manifest.id, manifest.alias));
    var resources = manifest.loaderConfig.scriptResources;
    var pathConfig = resources[name];
    var loadPromise;
    if (pathConfig && pathConfig.globalDependencies) {
        var depPromises = pathConfig.globalDependencies.map(function (dep) { return _loadPathDependency(manifest, dep, loadedPathDependencies); });
        loadPromise = Promise.all(depPromises).then(function () {
            return _systemImportPathDependency(manifest, name);
        }, function () {
            throw ErrorBuilder.buildLoadPathDependencyBlockedError(manifest, name);
        });
    }
    else {
        loadPromise = _systemImportPathDependency(manifest, name);
    }
    loadedPathDependencies.set(name, loadPromise);
    return loadPromise.then(function (load) {
        qosMonitor.writeSuccess(qosExtraData);
        return load;
    }, function (error) {
        qosMonitor.writeUnexpectedFailure(undefined, error, qosExtraData);
        throw error;
    });
}
function _systemImportPathDependency(manifest, name) {
    return _systemJsLoader.systemImport(normalizeName(manifest, name)).catch(function (e) {
        return _processSystemImportErrors(manifest, name, [ResourceUrlChecker.checkResourceUrl], function () { return ErrorBuilder.buildLoadPathDependencyError(manifest, name, e); });
    });
}
function _loadEntryPoint(manifest) {
    var entryPointModule = _systemJsLoader.systemImport(normalizeName(manifest)).catch(function (e) {
        return _processSystemImportErrors(manifest, manifest.loaderConfig.entryModuleId, [ResourceUrlChecker.checkResourceUrl, _checkEntryPointDependenciesError], function () { return ErrorBuilder.buildLoadEntryPointError(manifest, e); });
    });
    return entryPointModule.then(function (module) {
        return _getExportFromModule(manifest, module);
    });
}
function _processSystemImportErrors(manifest, name, errorProcessors, buildDefaultError) {
    return Promise.race(errorProcessors.map(function (errorProcessor) { return errorProcessor(manifest, name); }))
        .then(
    function () {
        throw buildDefaultError();
    }, function (e) { throw e; });
}
function _checkEntryPointDependenciesError(manifest, name) {
    var dependencies = _systemJsLoader.getDependencies(manifest);
    var resources = manifest.loaderConfig.scriptResources;
    dependencies.forEach(function (depName) {
        if (!resources[depName]) {
            throw ErrorBuilder.buildModuleHasUndeclaredDependencyError(manifest, depName);
        }
    });
    return Promise.resolve();
}
function _getExportFromModule(manifest, module) {
    var retValue = module;
    if (manifest.loaderConfig.exportName) {
        retValue = module[manifest.loaderConfig.exportName];
        _systemJsLoader.ensure(normalizeName(manifest, manifest.loaderConfig.exportName), retValue);
    }
    return retValue;
}
function _buildQosExtraData(manifest) {
    return {
        manifestId: manifest.id,
        version: manifest.version,
        alias: manifest.alias,
        isInternal: manifest.isInternal,
        isDebug: manifest._isDebug,
        loader: 'systemjs'
    };
}
