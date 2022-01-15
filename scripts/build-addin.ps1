param(
    # Package names to put into add-in
    [Parameter(Position=0)]
    [string[]]
    $PackageNames,
    # Generate add-in, or presentation
    [Parameter(Mandatory=$false)]
    [ValidateSet('AddIn','Presentation')]
    [string]
    $Generate = 'AddIn',
    # Output file name, without extension
    [Parameter(Mandatory=$false)]
    [string]
    $OutFileName = "Maxoffice-PowerPoint-Macros-Collection",
    # The directory where the output file will be created
    [string]
    $OutPath = "..\out\",
    # The directory where packages can be found
    [Parameter(Mandatory=$false)]
    [string]
    $PackagesPath = "../packages"
)
try {
    $packagedirs = (Get-ChildItem -Path $PackagesPath -ErrorAction Stop) 
}
catch {
    Write-Host "Could not find the packages directory. Exiting."
    return
}

if ($packagedirs.Count -eq 0) {
    Write-Error "No packages found."
    return
}

# Package directories have been found. 

# If the "PackageNames" parameter has
# been passed, filter them by that.
if($PackageNames.Count -gt 0) {
    # Sanitize parameter
    $PackageNames = ($PackageNames | ForEach-Object { $_.ToLowerInvariant() })

    # Filter packages
    $packagedirs = ($packagedirs | Where-Object { $PackageNames.Contains($_.Name.ToLowerInvariant()) })
} 

# Store the paths in an array
$packagepaths = $packagedirs | ForEach-Object { $_.FullName }
    
# Create PowerPoint object
$ppa = New-Object -ComObject PowerPoint.Application

$newPpt = $ppa.Presentations.Add($false)

try {
    # Try to get the VBA project object.
    # If VBA Object model access is not trusted, $vbp
    # will contain $null
    $vbp = $newPpt.VBProject
    
    if($null -eq $vbp) {
        Write-Error "Access to VBA Object Model not trusted. Please check the Trust Access to the VBA Object model checkbox in the PowerPoint Trust Centre." -ErrorAction Stop
    }

    $packagepaths | ForEach-Object {
        $packagepathspec = "$PSItem\*"
        $basfiles = (Get-ChildItem -Path $packagepathspec -Include *.bas,*.frm,*.cls)
        $basfiles | ForEach-Object {
            # try {

            $newComponent = $vbp.VBComponents.Import($_.FullName)
            Write-Output "$($newComponent.Name)"
            #}
            #catch {
            #    Write-Error "Horrible error: $PSItem" -ErrorAction Stop
            #}        
        }
    }

    # Check output path and create if needed
    if(-not [System.IO.Path]::IsPathRooted($OutPath)) {
        $OutPath = Join-Path -Path $PWD -ChildPath $OutPath
    }

    $outdirexists = (Test-Path -PathType Container $OutPath)
    if($outdirexists -eq $false) {
        New-Item -ItemType Directory -Force -Path $OutPath
    }

    if($Generate -eq "Presentation") {
        $savefileFullName = (Join-Path -Path $OutPath -ChildPath $OutFileName)
        # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
        $newPpt.SaveAs($savefileFullName, 25, $false)
    } else {
        $savefileFullName = (Join-Path -Path $OutPath -ChildPath $OutFileName)
        # 30 = ppSaveAsOpenXMLAddin 
        $newPpt.SaveAs($savefileFullName, 30, $false)
    }
}
finally {
    $newPpt.Close()
    $newPpt = $null

    $ppa.Quit()
    $ppa = $null
}
