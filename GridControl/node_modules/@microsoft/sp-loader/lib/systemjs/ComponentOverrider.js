import { Validate } from '@microsoft/sp-core-library';
import SystemJsLoader from './SystemJsLoader';
import ManifestStore from '../stores/ManifestStore';
import ComponentStore from '../stores/ComponentStore';
import normalizeName from './normalizeName';
var ComponentOverrider = (function () {
    function ComponentOverrider() {
    }
    ComponentOverrider.overrideComponent = function (componentId, componentModule) {
        Validate.isNotNullOrUndefined(componentModule, 'componentModule');
        var manifest = ManifestStore.instance.tryGetManifest(componentId);
        if (!manifest) {
            return;
        }
        var normalizedName = normalizeName(manifest);
        SystemJsLoader.instance.ensure(normalizedName, componentModule);
        ComponentStore.instance.storeLoadedComponent(manifest.id, manifest.version, componentModule);
    };
    return ComponentOverrider;
}());
export default ComponentOverrider;
