"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var path = require("path");
var fsx = require("fs-extra");
var node_core_library_1 = require("@microsoft/node-core-library");
var Resolve_1 = require("@microsoft/sp-build-core-tasks/lib/utilities/Resolve");
var ASSEMBLY_MANIFEST_ID = '44bae1a2-d2eb-4e10-8c21-a1dbdce1036f'; // loader assembly
var REACT_MANIFEST_ID = '0d910c1c-13b9-4e1c-9aa4-b008c5e42d7d';
var REACT_DOM_MANIFEST_ID = 'aa0a46ec-1505-43cd-a44a-93f3a5aa460a';
var WorkbenchPage = (function () {
    function WorkbenchPage() {
        this._packageNameSymlinks = new Map();
        this._projectRootPath = process.cwd();
        this._tempFolderPath = path.join(this._projectRootPath, 'temp');
        this._workbenchPackagesTempPath = path.join(this._tempFolderPath, 'workbench-packages');
        this._workbenchRootPath = path.resolve(__dirname, '..', '..');
        this.handleWorkbenchRequest = this.handleWorkbenchRequest.bind(this);
        fsx.ensureDirSync(this._workbenchPackagesTempPath);
        fsx.emptyDirSync(this._workbenchPackagesTempPath);
    }
    /* tslint:disable-next-line:no-any */
    WorkbenchPage.prototype.handleWorkbenchRequest = function (request, response) {
        var workbenchFilePath = path.join(this._tempFolderPath, 'workbench.html');
        fsx.ensureDirSync(path.dirname(workbenchFilePath));
        fsx.writeFileSync(workbenchFilePath, this._generateWorkbenchPageContent());
        var redirectPath = this._getUrlForPath(workbenchFilePath);
        response.redirect(redirectPath);
    };
    WorkbenchPage.prototype._generateWorkbenchPageContent = function () {
        // Get the manifests.js URL.
        var manifestsJsPath = path.join(this._tempFolderPath, 'manifests.js');
        var manifestJsUrl = this._getUrlForPath(manifestsJsPath);
        var workbenchInitUrl = this._getUrlForPackagePath('@microsoft/sp-webpart-workbench', path.join('lib', 'api', 'workbenchInit.js'));
        var assemblyUrl = this._assemblyUrl;
        var head;
        var body;
        if (!assemblyUrl) {
            // This doesn't need to be localized because this string can only show up when the workbench
            //  has been NPM installed, and we only ship English modules to NPM
            body = this._getErrorBody('The script containing the initialization code could not be resolved. Unable to load workbench');
        }
        else {
            head = "\n    <script type=\"text/javascript\" src=\"" + manifestJsUrl + "\"></script>\n    <script type=\"text/javascript\" src=\"" + assemblyUrl + "\"></script>\n    <script type=\"text/javascript\">\n      // The workbenchInit.js script is commonJS and expects an \"exports\" object\n      var exports = {};\n\n      var syntheticManifests = " + JSON.stringify(this._generateSyntheticManifests(), undefined, 2) + "\n    </script>\n    <script type=\"text/javascript\" src=\"" + workbenchInitUrl + "\"></script>\n    <script type=\"text/javascript\">\n      // Clean up the \"exports\" object\n      exports = undefined;\n    </script>";
            body = "\n    <script type=\"text/javascript\">\n      window.spModuleLoader.start(window.preloadedData);\n    </script>";
        }
        return this._getPageWithBodyAndHead(body, head);
    };
    Object.defineProperty(WorkbenchPage.prototype, "_faviconUrl", {
        /**
         * Generate a favicon URL. Chrome will request one if we don't give a URL, so in order to avoid an error in the
         * console we'll give a URL.
         */
        get: function () {
            return this._getUrlForPackagePath('@microsoft/sp-webpart-workbench', path.join('lib', 'api', 'assets', 'server-icon.png'));
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(WorkbenchPage.prototype, "_assemblyUrl", {
        /**
         * Find the assembly.
         */
        get: function () {
            try {
                var assemblyPath = this._getAssemblyPath();
                return this._getUrlForPackagePath('@microsoft/sp-loader', assemblyPath);
            }
            catch (e) {
                // Unable to get the assembly URL. We'll inform the user if they try to load the workbench page.
                return undefined;
            }
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(WorkbenchPage.prototype, "_serveConfig", {
        /**
         * Get the running project's serve.json
         */
        get: function () {
            if (!this._cachedServeConfig) {
                try {
                    this._cachedServeConfig = node_core_library_1.JsonFile.load(path.join(this._projectRootPath, 'config', 'serve.json'));
                }
                catch (e) {
                    // Could not load serve config, setting defaults.
                    this._cachedServeConfig = {
                        port: 4321,
                        initialPage: ''
                    };
                }
            }
            return this._cachedServeConfig;
        },
        enumerable: true,
        configurable: true
    });
    WorkbenchPage.prototype._getErrorBody = function (errorMessage) {
        return "\n    <span style=\"font-weight:bold;color:red;font-size:20px;\">\n      " + errorMessage + "\n    </span>";
    };
    WorkbenchPage.prototype._getPageWithBodyAndHead = function (body, head) {
        return "<!doctype html>\n  <html dir=\"ltr\">\n  <head>\n    <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\n\n    <title>SharePoint Web Part Workbench</title>\n\n    <link rel=\"shortcut icon\" href=\"" + this._faviconUrl + "\" />\n  " + (head || '') + "\n  </head>\n  <body>\n  " + body + "\n  </body>\n  </html>\n  ";
    };
    WorkbenchPage.prototype._getAssemblyPath = function () {
        var loaderRootPath = Resolve_1.resolvePackage('@microsoft/sp-loader', this._workbenchRootPath);
        var distFolder = path.join(loaderRootPath, 'dist');
        var assemblyManifestPath = path.join(distFolder, ASSEMBLY_MANIFEST_ID + ".manifest.json");
        var assemblyManifest = node_core_library_1.JsonFile.load(assemblyManifestPath);
        var assemblyEntryModuleConfig = assemblyManifest.loaderConfig.scriptResources[assemblyManifest.loaderConfig.entryModuleId];
        if (!assemblyEntryModuleConfig) {
            throw 'Unable to find assembly entrypoint';
        }
        var foundPath;
        switch (assemblyEntryModuleConfig.type) {
            case 'path':
                foundPath = this._disambiguateManifestPath(assemblyEntryModuleConfig.path);
                break;
            case 'localizedPath':
                foundPath = this._disambiguateManifestPath(assemblyEntryModuleConfig.defaultPath);
                break;
            default:
                throw "Invalid module assembly manifest found at " + assemblyManifestPath;
        }
        if (!foundPath) {
            throw 'Unable to find assembly path';
        }
        return foundPath;
    };
    WorkbenchPage.prototype._disambiguateManifestPath = function (pathToDisambiguate) {
        if (typeof pathToDisambiguate === 'string') {
            return pathToDisambiguate;
        }
        else {
            return pathToDisambiguate.debug || pathToDisambiguate.default;
        }
    };
    WorkbenchPage.prototype._getUrlForPackagePath = function (packageName, pathInsidePackage) {
        var normalizedPackageName = packageName.toUpperCase();
        if (!this._packageNameSymlinks.has(normalizedPackageName)) {
            var packagePath = Resolve_1.resolvePackage(packageName, this._workbenchRootPath);
            var symlinkPath = path.join(this._workbenchPackagesTempPath, packageName.replace(/\//g, '_') // Replace slashes with underscores so we don't create nested directories
            );
            fsx.symlinkSync(packagePath, symlinkPath, 'junction');
            this._packageNameSymlinks.set(normalizedPackageName, symlinkPath);
        }
        var linkedPackagePath = this._packageNameSymlinks.get(normalizedPackageName);
        var filePath = path.join(linkedPackagePath, pathInsidePackage);
        return this._getUrlForPath(filePath);
    };
    WorkbenchPage.prototype._getUrlForPath = function (filePath) {
        var relativePath = path.relative(this._projectRootPath, filePath);
        return (this._serveConfig.https ? 'https' : 'http') +
            ("://localhost:" + this._serveConfig.port + "/" + relativePath.replace(/\\/g, '/'));
    };
    WorkbenchPage.prototype._generateSyntheticManifests = function () {
        // Figure out where React is
        var reactPackagePath = Resolve_1.resolvePackage('react', __dirname);
        var reactPackage = node_core_library_1.JsonFile.load(path.join(reactPackagePath, 'package.json'));
        var reactUrl = this._getUrlForPackagePath('react', path.join('dist', 'react.js'));
        // Figure out where React-DOM is
        var reactDomPackagePath = Resolve_1.resolvePackage('react-dom', __dirname);
        var reactDomPackage = node_core_library_1.JsonFile.load(path.join(reactDomPackagePath, 'package.json'));
        var reactDomUrl = this._getUrlForPackagePath('react-dom', path.join('dist', 'react-dom.js'));
        var reactManifest = {
            id: REACT_MANIFEST_ID,
            manifestVersion: 2,
            componentType: 'Library',
            alias: 'react',
            version: reactPackage.version,
            loaderConfig: {
                internalModuleBaseUrls: [],
                entryModuleId: 'react',
                scriptResources: {
                    'react': {
                        type: 'path',
                        path: reactUrl
                    }
                }
            }
        };
        var reactDomManifest = {
            id: REACT_DOM_MANIFEST_ID,
            manifestVersion: 2,
            componentType: 'Library',
            alias: 'react-dom',
            version: reactDomPackage.version,
            loaderConfig: {
                internalModuleBaseUrls: [],
                entryModuleId: 'react-dom',
                scriptResources: {
                    'react-dom': {
                        type: 'path',
                        path: reactDomUrl
                    },
                    'react': {
                        type: 'component',
                        id: REACT_MANIFEST_ID,
                        version: reactPackage.version
                    }
                }
            }
        };
        return [
            reactManifest,
            reactDomManifest
        ];
    };
    return WorkbenchPage;
}());
exports.WorkbenchPage = WorkbenchPage;

//# sourceMappingURL=WorkbenchPage.js.map
