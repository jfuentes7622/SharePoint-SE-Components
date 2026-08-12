import { Text } from '@microsoft/sp-core-library';
import { isEmpty } from '@microsoft/sp-lodash-subset';
import { _QosMonitor } from '@ms/sp-telemetry';
import { _TraceLogger } from '@microsoft/sp-diagnostics';
import RequireJsLoader from './RequireJsLoader';
import ManifestStore from '../stores/ManifestStore';
import ComponentStore from '../stores/ComponentStore';
import * as telemetryConstants from '../utilities/telemetryConstants';
import { resolvePath } from '../utilities/resolveAddress';
import normalizeName, { normalizeFailoverPathName } from './normalizeName';
import ResourceUrlChecker from '../utilities/ResourceUrlChecker';
import ErrorBuilder from '../error/ErrorBuilder';
import strings from '../SPLoader.resx';
var FIRST_RETRY = 1;
var MAX_NUMBER_RETRIES = 3;
var LOAD_COMPONENT_IMPL_EVENT_NAME = 'loadComponentImpl';
var _requireJsLoader = RequireJsLoader.instance;
export default function loadComponent(manifest) {
    if (!manifest) {
        return Promise.reject(new Error(Text.format(strings.isUndefinedValidateError, 'manifest')));
    }
    var cachedModule = ComponentStore.instance.tryGetComponent(manifest.id, manifest.version);
    if (cachedModule) {
        return cachedModule;
    }
    var componentPromise = _loadComponentRetryStrategy(manifest, FIRST_RETRY, MAX_NUMBER_RETRIES)
        .catch(function (requireJsError) {
        ComponentStore.instance.deleteComponent(manifest.id, manifest.version);
        throw requireJsError;
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
        _requireJsLoader.requireDelete(manifest);
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
        _requireJsLoader.configure(manifest);
    }
    catch (error) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, error.message, LOAD_COMPONENT_IMPL_EVENT_NAME);
        return Promise.reject(error);
    }
    var componentDeps = [];
    var pathDeps = [];
    try {
        componentDeps = _loadComponentDependencies(manifest);
    }
    catch (error) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, error.message, LOAD_COMPONENT_IMPL_EVENT_NAME);
        return Promise.reject(error);
    }
    try {
        pathDeps = _loadPathDependencies(manifest);
    }
    catch (error) {
        _TraceLogger.logVerbose(telemetryConstants.loadComponentLogSource, error.message, LOAD_COMPONENT_IMPL_EVENT_NAME);
        return Promise.reject(error);
    }
    return Promise.all(componentDeps.concat(pathDeps)).then(function (components) {
        if (!manifest.loaderConfig.entryModuleId) {
            _requireJsLoader.ensure(normalizeName(manifest), {});
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
                    var dep = _requireJsLoader.requireLoad(normalizeFailoverPathName(name_1))
                        .catch(function (e) {
                        return _processLoadErrors(manifest, name_1, [ResourceUrlChecker.checkResourceUrl], function () { return ErrorBuilder.buildLoadComponentDependencyFailoverPathError(manifest, name_1, resolvePath(moduleConfiguration_1.failoverPath), e); });
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
    if (!pathConfig) {
        return Promise.reject(ErrorBuilder.buildMissingPathDependencyError(manifest, name));
    }
    var loadPromise;
    if (pathConfig.globalDependencies) {
        var depPromises = pathConfig.globalDependencies.map(function (dep) { return _loadPathDependency(manifest, dep, loadedPathDependencies); });
        loadPromise = Promise.all(depPromises).then(function () {
            return _requireLoadPathDependency(manifest, name);
        }, function () {
            throw ErrorBuilder.buildLoadPathDependencyBlockedError(manifest, name);
        });
    }
    else {
        loadPromise = _requireLoadPathDependency(manifest, name);
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
function _requireLoadPathDependency(manifest, name) {
    var globalName = manifest.loaderConfig.scriptResources[name].globalName;
    return _requireJsLoader.requireLoad(normalizeName(manifest, name), globalName)
        .catch(function (e) {
        return _processLoadErrors(manifest, name, [ResourceUrlChecker.checkResourceUrl,
            checkDependencies], function () { return ErrorBuilder.buildLoadPathDependencyError(manifest, name, e); });
    });
}
function _loadEntryPoint(manifest) {
    var entryPointModule = _requireJsLoader.requireLoad(normalizeName(manifest)).catch(function (e) {
        return _processLoadErrors(manifest, manifest.loaderConfig.entryModuleId, [ResourceUrlChecker.checkResourceUrl,
            checkDependencies], function () { return ErrorBuilder.buildLoadEntryPointError(manifest, e); });
    });
    return entryPointModule.then(function (module) {
        return _getExportFromModule(manifest, module);
    });
}
function _getExportFromModule(manifest, module) {
    var retValue = module;
    if (manifest.loaderConfig.exportName) {
        retValue = module[manifest.loaderConfig.exportName];
    }
    return retValue;
}
function _processLoadErrors(manifest, name, errorProcessors, buildDefaultError) {
    return Promise.race(errorProcessors.map(function (errorProcessor) { return errorProcessor(manifest, name); }))
        .then(
    function () {
        throw buildDefaultError();
    }, function (e) { throw e; });
}
function checkDependencies(manifest, name) {
    var moduleName = normalizeName(manifest, name);
    var missingDependencies = _requireJsLoader.getMissingDependencies(moduleName);
    if (missingDependencies.length > 0) {
        return Promise.reject(ErrorBuilder.buildModuleHasFailedDependencyError(moduleName, missingDependencies.join(', ')));
    }
    return Promise.resolve();
}
