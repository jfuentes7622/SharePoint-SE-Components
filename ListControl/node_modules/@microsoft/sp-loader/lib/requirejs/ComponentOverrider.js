import { Validate } from '@microsoft/sp-core-library';
import ManifestStore from '../stores/ManifestStore';
import ComponentStore from '../stores/ComponentStore';
import RequireJsLoader from './RequireJsLoader';
import normalizeName from './normalizeName';
var ComponentOverrider = (function () {
    function ComponentOverrider() {
    }
    ComponentOverrider.overrideComponent = function (componentId, componentModule) {
        Validate.isNonemptyString(componentId, 'componentId');
        Validate.isNotNullOrUndefined(componentModule, 'componentModule');
        var manifest = ManifestStore.instance.tryGetManifest(componentId);
        if (!manifest) {
            return;
        }
        var normalizedName = normalizeName(manifest);
        RequireJsLoader.instance.ensure(normalizedName, componentModule);
        ComponentStore.instance.storeComponent(manifest.id, manifest.version, Promise.resolve(componentModule));
    };
    return ComponentOverrider;
}());
export default ComponentOverrider;
