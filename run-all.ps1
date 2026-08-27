$projects = @("Accordion", "BannerClock", "Calendar", "Carousel", "Contact", "DynamicForms", "FullWidthControl", "GridControl", "LinkButton", "ListControl", "Marquee", "PrintControl", "ReportForms", "ScriptEditor", "Tab", "Tiles")
$rootDir = "D:\Source\SharePoint-SE-Components"
$results = @()

foreach ($proj in $projects) {
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "Processing $proj..." -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    
    $projDir = Join-Path $rootDir $proj
    if (-not (Test-Path $projDir)) {
        $results += [PSCustomObject]@{
            Project = $proj
            Clean = "Skipped (Dir Not Found)"
            Bundle = "Skipped"
            Package = "Skipped"
            SppkgPath = ""
            SppkgSize = ""
            SppkgDate = ""
            ErrorDetails = "Directory not found: $projDir"
        }
        continue
    }
    
    Set-Location $projDir
    
    $cleanStatus = "Not Run"
    $bundleStatus = "Not Run"
    $packageStatus = "Not Run"
    $errorDetails = ""
    $sppkgPath = ""
    $sppkgSize = ""
    $sppkgDate = ""
    
    # 1. Clean
    Write-Host "Running gulp clean..."
    $cleanOut = & npx gulp clean 2>&1
    $cleanCode = $LASTEXITCODE
    if ($cleanCode -eq 0) {
        $cleanStatus = "Success"
        # 2. Bundle
        Write-Host "Running gulp bundle..."
        $bundleOut = & npx gulp bundle --ship 2>&1
        $bundleCode = $LASTEXITCODE
        if ($bundleCode -eq 0) {
            $bundleStatus = "Success"
            # 3. Package
            Write-Host "Running gulp package-solution..."
            $packageOut = & npx gulp package-solution --ship 2>&1
            $packageCode = $LASTEXITCODE
            if ($packageCode -eq 0) {
                $packageStatus = "Success"
                
                # Search for sppkg
                $sppkgFiles = Get-ChildItem -Path (Join-Path $projDir "sharepoint/solution") -Filter "*.sppkg" -ErrorAction SilentlyContinue
                if ($sppkgFiles) {
                    $sppkg = $sppkgFiles[0]
                    $sppkgPath = $sppkg.FullName
                    $sppkgSize = $sppkg.Length
                    $sppkgDate = $sppkg.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                } else {
                    $errorDetails += "Package-solution succeeded, but no .sppkg found in sharepoint/solution!`n"
                }
            } else {
                $packageStatus = "Failed ($packageCode)"
                $errExcerpts = $packageOut | Out-String
                $errorDetails += "--- Package-solution Failure Excerpts ---`n$errExcerpts`n"
            }
        } else {
            $bundleStatus = "Failed ($bundleCode)"
            $errExcerpts = $bundleOut | Out-String
            $errorDetails += "--- Bundle Failure Excerpts ---`n$errExcerpts`n"
        }
    } else {
        $cleanStatus = "Failed ($cleanCode)"
        $errExcerpts = $cleanOut | Out-String
        $errorDetails += "--- Clean Failure Excerpts ---`n$errExcerpts`n"
    }
    
    $results += [PSCustomObject]@{
        Project = $proj
        Clean = $cleanStatus
        Bundle = $bundleStatus
        Package = $packageStatus
        SppkgPath = $sppkgPath
        SppkgSize = $sppkgSize
        SppkgDate = $sppkgDate
        ErrorDetails = $errorDetails
    }
}

# Write results
$results | Export-Clixml -Path (Join-Path $rootDir "results.xml")

Write-Host "Done processing all projects!" -ForegroundColor Green
