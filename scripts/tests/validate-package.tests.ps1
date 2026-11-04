$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Join-Path ([System.IO.Path]::GetTempPath()) ('se-package-test-' + [guid]::NewGuid())
$project = Join-Path $root 'Example'
$staging = Join-Path $root 'staging'
$validator = Join-Path (Split-Path $PSScriptRoot -Parent) 'validate-package.ps1'
$solutionId = '11111111-1111-1111-1111-111111111111'
$componentId = '22222222-2222-2222-2222-222222222222'
$featureId = '33333333-3333-3333-3333-333333333333'
try {
    foreach ($directory in @('config', 'src', 'sharepoint\solution')) {
        New-Item -ItemType Directory -Force (Join-Path $project $directory) | Out-Null
    }
    New-Item -ItemType Directory -Force (Join-Path $staging 'ClientSideAssets') | Out-Null
    @{
        solution = @{
            id = $solutionId
            version = '1.0.0.2'
            includeClientSideAssets = $true
            features = @(@{ id = $featureId; version = '1.0.0.1' })
        }
        paths = @{ zippedPackage = 'solution/example.sppkg' }
    } | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $project 'config\package-solution.json')
    @{ version = '1.0.1' } | ConvertTo-Json | Set-Content (Join-Path $project 'package.json')
    @{ id = $componentId; version = '*' } | ConvertTo-Json | Set-Content (Join-Path $project 'src\Example.manifest.json')
    'asset' | Set-Content (Join-Path $staging 'ClientSideAssets\example.js')
    $app = '<App xmlns="http://schemas.microsoft.com/sharepoint/2012/app/manifest" ProductID="' + $solutionId + '" Version="1.0.0.2" />'
    $feature = '<Feature xmlns="http://schemas.microsoft.com/sharepoint/" Id="' + $featureId + '" Version="1.0.0.1" />'
    $manifest = @{
        id = $componentId
        version = '1.0.1'
        loaderConfig = @{ scriptResources = @{ example = @{ type = 'path'; path = 'example.js' } } }
    }
    function Write-Package {
        param([string]$AppXml = $app, [string]$FeatureXml = $feature)
        $AppXml | Set-Content (Join-Path $staging 'AppManifest.xml')
        $FeatureXml | Set-Content (Join-Path $staging 'feature.xml')
        $componentXml = [System.Xml.XmlDocument]::new()
        $componentXml.LoadXml('<Elements><ClientSideComponent /></Elements>')
        $componentXml.DocumentElement.FirstChild.SetAttribute('Id', $componentId)
        $componentXml.DocumentElement.FirstChild.SetAttribute('ComponentManifest', ($manifest | ConvertTo-Json -Depth 8 -Compress))
        $componentXml.Save((Join-Path $staging 'component.xml'))
        $target = Join-Path $project 'sharepoint\solution\example.sppkg'
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $target)
    }
    function Assert-Fails {
        param([string]$Pattern)
        $failure = $null
        try { & $validator -Project Example -RepositoryRoot $root | Out-Null }
        catch { $failure = $_.Exception.Message }
        if (-not $failure -or $failure -notmatch $Pattern) {
            throw "Expected validation failure matching '$Pattern', got '$failure'"
        }
    }
    Write-Package
    & $validator -Project Example -RepositoryRoot $root
    Write-Package -AppXml ($app.Replace('1.0.0.2', '1.0.0.1'))
    Assert-Fails 'solution version differs'
    Write-Package -AppXml ($app.Replace($solutionId, $componentId))
    Assert-Fails 'solution ID differs'
    Write-Package -FeatureXml ($feature.Replace('1.0.0.1', '1.0.0.0'))
    Assert-Fails 'feature XML'
    Write-Package -AppXml '<App>'
    Assert-Fails 'XML|Unexpected|unclosed'
    Write-Package -AppXml ('<!DOCTYPE App [<!ENTITY x SYSTEM "file:///nonexistent">]>' + $app)
    Assert-Fails 'DTD'
    $manifest.version = '1.0.0'
    Write-Package
    Assert-Fails 'component version is stale'
    $manifest.version = '1.0.1'
    $manifest.loaderConfig.scriptResources.example.path = 'missing.js'
    Write-Package
    Assert-Fails 'missing asset'
    $manifest.loaderConfig.scriptResources.example.path = 'example.js'
    $manifest.loaderConfig.scriptResources['@microsoft/sp-core-library'] = @{ type = 'component'; version = '1.6.0' }
    Write-Package
    Assert-Fails 'not 1.5.1'
    Write-Output 'Package tests passed: valid package; bad solution ID/version, feature version, XML, DTD, component version, assets and SPFx rejected.'
} finally {
    # Delete only the unique fixture directory created by this test.
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
