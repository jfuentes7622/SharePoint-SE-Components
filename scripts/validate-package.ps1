param(
    [Parameter(Mandatory = $true)][string]$Project,
    [string]$RepositoryRoot = (Split-Path $PSScriptRoot -Parent)
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = [System.IO.Path]::GetFullPath($RepositoryRoot)
$projectPath = Join-Path $root $Project
if ($Project -notmatch '^[A-Za-z][A-Za-z0-9]*$') {
    throw 'Project must be a top-level component name.'
}
$config = Get-Content (Join-Path $projectPath 'config\package-solution.json') -Raw | ConvertFrom-Json
$package = Join-Path (Join-Path $projectPath 'sharepoint') $config.paths.zippedPackage
if (-not (Test-Path -LiteralPath $package -PathType Leaf)) {
    throw "Missing solution package: $package"
}
$archive = [System.IO.Compression.ZipFile]::OpenRead($package)
try {
    $app = $archive.Entries | Where-Object { $_.FullName -eq 'AppManifest.xml' }
    if (@($app).Count -ne 1) { throw 'Package must contain exactly one AppManifest.xml.' }
    $features = @()
    $components = @()
    foreach ($entry in $archive.Entries) {
        if ($entry.FullName -notmatch '\.xml$') { continue }
        $reader = [System.IO.StreamReader]::new($entry.Open())
        try {
            $settings = [System.Xml.XmlReaderSettings]::new()
            $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
            $settings.XmlResolver = $null
            $xmlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($reader.ReadToEnd()), $settings)
            try {
                $xml = [System.Xml.XmlDocument]::new()
                $xml.XmlResolver = $null
                $xml.Load($xmlReader)
            } finally { $xmlReader.Dispose() }
        } finally { $reader.Dispose() }
        if ($entry.FullName -eq 'AppManifest.xml') {
            $appNode = $xml.DocumentElement
            if ($appNode.GetAttribute('ProductID').Trim('{}') -ne $config.solution.id) {
                throw 'Packaged solution ID differs from package-solution.json.'
            }
            if ($appNode.GetAttribute('Version') -ne $config.solution.version) {
                throw 'Packaged solution version differs from package-solution.json.'
            }
        }
        foreach ($feature in $xml.SelectNodes("//*[local-name()='Feature' and @Id and @Version]")) {
            $features += $feature
        }
        foreach ($component in $xml.SelectNodes("//*[@ComponentManifest]")) {
            $manifest = $component.GetAttribute('ComponentManifest') | ConvertFrom-Json
            if ($component.GetAttribute('Id').Trim('{}') -ne $manifest.id) {
                throw 'Packaged component declaration and embedded manifest IDs differ.'
            }
            $components += $manifest
            foreach ($resource in $manifest.loaderConfig.scriptResources.PSObject.Properties) {
                if ($resource.Name -like '@microsoft/sp-*' -and $resource.Value.version -ne '1.5.1') {
                    throw "Packaged SPFx resource $($resource.Name) is not 1.5.1."
                }
                if (-not $config.solution.includeClientSideAssets) { continue }
                $paths = @($resource.Value.path, $resource.Value.defaultPath)
                $paths += @($resource.Value.paths.PSObject.Properties | ForEach-Object { $_.Value })
                foreach ($asset in $paths) {
                    if (-not $asset -or $asset -match '^(https?:)?//') { continue }
                    $assetPath = 'ClientSideAssets/' + $asset.TrimStart('./')
                    if (-not ($archive.Entries | Where-Object { $_.FullName -eq $assetPath })) {
                        throw "Packaged component references missing asset: $assetPath"
                    }
                }
            }
        }
    }
    foreach ($expected in $config.solution.features) {
        $found = @($features | Where-Object { $_.GetAttribute('Id').Trim('{}') -eq $expected.id })
        if ($found.Count -ne 1 -or $found[0].GetAttribute('Version') -ne $expected.version) {
            throw "Generated feature XML is missing, duplicated or stale for $($expected.id)."
        }
    }
    $sourceManifests = @(Get-ChildItem (Join-Path $projectPath 'src') -Filter '*.manifest.json' -Recurse -File)
    if ($sourceManifests.Count -eq 0) { throw 'No source component manifests found.' }
    if ($components.Count -ne $sourceManifests.Count) {
        throw 'Package contains an unexpected number of component manifests.'
    }
    foreach ($file in $sourceManifests) {
        $source = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $found = @($components | Where-Object { $_.id -eq $source.id })
        if ($found.Count -ne 1) {
            throw "Packaged component is missing or duplicated: $($source.id)"
        }
        $expectedVersion = if ($source.version -eq '*') {
            (Get-Content (Join-Path $projectPath 'package.json') -Raw | ConvertFrom-Json).version
        } else { $source.version }
        if ($found[0].version -ne $expectedVersion) {
            throw "Packaged component version is stale: $($source.id)"
        }
    }
    Write-Output "SE package validation passed: $Project ($($config.solution.version)). Test-farm deployment is still required."
} finally {
    $archive.Dispose()
}
