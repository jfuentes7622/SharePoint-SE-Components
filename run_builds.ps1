$projects = @("Accordion", "Calendar", "Carousel", "Contact", "DynamicForms", "FormNavigator", "GridControl", "LinkButton", "ListControl", "Marquee", "PrintControl", "ReportForms", "ScriptEditor", "Tab", "Tiles")
$baseDir = "d:\Source\SharePoint-SE-Components"
$results = @()

foreach ($proj in $projects) {
    Write-Host "=== Starting Project: $proj ==="
    $projDir = Join-Path $baseDir $proj
    if (-not (Test-Path $projDir)) {
        Write-Host "Directory $projDir does not exist!"
        Continue
    }
    
    $res = [PSCustomObject]@{
        Project = $proj
        CleanExit = $null
        CleanOutput = ""
        BundleExit = $null
        BundleOutput = ""
        PackageExit = $null
        PackageOutput = ""
        PackagePath = ""
        AllErrorsWarnings = ""
    }
    
    Set-Location $projDir
    
    # 1. Clean
    Write-Host "Running clean..."
    $cleanOut = cmd /c "npx gulp clean 2>&1"
    $res.CleanExit = $LASTEXITCODE
    # Extract line-by-line errors/warnings
    $cleanLines = $cleanOut | Where-Object { $_ -match "error" -or $_ -match "warning" -or $_ -match "Failed" -or $_ -match "Err" -or $_ -match "Warn" }
    
    # 2. Bundle
    Write-Host "Running bundle..."
    $bundleOut = cmd /c "npx gulp bundle --ship 2>&1"
    $res.BundleExit = $LASTEXITCODE
    $bundleLines = $bundleOut | Where-Object { $_ -match "error" -or $_ -match "warning" -or $_ -match "Failed" -or $_ -match "Err" -or $_ -match "Warn" }
    
    # 3. Package
    Write-Host "Running package-solution..."
    $packageOut = cmd /c "npx gulp package-solution --ship 2>&1"
    $res.PackageExit = $LASTEXITCODE
    $packageLines = $packageOut | Where-Object { $_ -match "error" -or $_ -match "warning" -or $_ -match "Failed" -or $_ -match "Err" -or $_ -match "Warn" }
    
    # Find package path
    if ($res.PackageExit -eq 0) {
        $sppkgDir = Join-Path $projDir "sharepoint\solution"
        if (Test-Path $sppkgDir) {
            $sppkg = Get-ChildItem -Path $sppkgDir -Filter "*.sppkg" | Select-Object -First 1
            if ($sppkg) {
                # relative package path
                $res.PackagePath = "$proj\sharepoint\solution\$($sppkg.Name)"
            }
        }
    }
    
    # Combine errs and warnings
    $combined = @()
    if ($cleanLines) { $combined += ($cleanLines | ForEach-Object { "[Clean] $_" }) }
    if ($bundleLines) { $combined += ($bundleLines | ForEach-Object { "[Bundle] $_" }) }
    if ($packageLines) { $combined += ($packageLines | ForEach-Object { "[Package] $_" }) }
    
    # Remove duplicates or empty items
    $combined = $combined | Where-Object { $_ -and $_.Trim() -ne "" }
    $res.AllErrorsWarnings = $combined -join " | "
    
    $results += $res
    Write-Host "=== Finished Project: $proj -> Clean: $($res.CleanExit), Bundle: $($res.BundleExit), Package: $($res.PackageExit) ==="
}

Set-Location $baseDir
# Output results as JSON for easy reading and robust formatting
$results | ConvertTo-Json -Depth 3 | Set-Content (Join-Path $baseDir "build_results.json")
$results | Format-Table -Property Project, CleanExit, BundleExit, PackageExit, PackagePath -AutoSize
