param(
    # Package names to put into add-in
    [Parameter(Position = 0)]
    [string[]]
    $PackageNames,
    # Generate add-in, or presentation
    [Parameter(Mandatory = $false)]
    [ValidateSet('AddIn', 'Presentation')]
    [string]
    $Generate = 'AddIn',
    # Output file name, without extension
    [Parameter(Mandatory = $false)]
    [string]
    $OutFileName = "Maxoffice-PowerPoint-Macros-Collection",
    # The directory where the output file will be created
    [string]
    $OutPath = "..\out\",
    # The directory where packages can be found
    [Parameter(Mandatory = $false)]
    [string]
    $PackagesPath = "../packages"
)



function ensureDirectory {
    param (
        # Parameter help description
        [Parameter(Mandatory = $true, Position = 0)]
        [string]
        $OutPath
    )

    $outdirexists = (Test-Path -PathType Container $OutPath)
    if ($outdirexists -eq $false) {
        New-Item -ItemType Directory -Force -Path $OutPath
    }
}

function Build-PPTFile {
    param(
        # Package names to put into add-in
        [Parameter(Position = 0)]
        [string[]]
        $PackageNames,
        # Generate add-in, or presentation
        [Parameter(Mandatory = $false)]
        [ValidateSet('AddIn', 'Presentation')]
        [string]
        $Generate = 'AddIn',
        # Output file name, without extension
        [Parameter(Mandatory = $false)]
        [string]
        $OutFileName = "Maxoffice-PowerPoint-Macros-Collection",
        # The directory where the output file will be created
        [string]
        $OutPath = "..\out\",
        # The directory where packages can be found
        [Parameter(Mandatory = $false)]
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
    if ($PackageNames.Count -gt 0) {
        # Sanitize parameter
        $PackageNames = ($PackageNames | ForEach-Object { $_.ToLowerInvariant() })
    
        # Filter packages
        $packagedirs = ($packagedirs | Where-Object { $PackageNames.Contains($_.Name.ToLowerInvariant()) })
    } 
    
    # Store the paths in an array
    $packagepaths = ($packagedirs | ForEach-Object { $_.FullName })

    $mcx = [xml]@"
    <customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
    <ribbon startFromScratch="false">
    <tabs>
    </tabs>
    </ribbon>
    </customUI>
"@

    $packagepaths | ForEach-Object {
        processPackageDir $_ $mcx
    }

    Write-Host "Resulting XML is: $($mcx.InnerXml)"

    return

    ### Create PowerPoint Document/AddIn file

    # Create PowerPoint object
    $ppa = New-Object -ComObject PowerPoint.Application
    
    $newPpt = $ppa.Presentations.Add($false)
    
    try {
        # Try to get the VBA project object.
        # If VBA Object model access is not trusted, $vbp
        # will contain $null
        $vbp = $newPpt.VBProject
        
        if ($null -eq $vbp) {
            Write-Error "Access to VBA Object Model not trusted. Please check the Trust Access to the VBA Object model checkbox in the PowerPoint Trust Centre." -ErrorAction Stop
        }
    
        $packagepaths | ForEach-Object {
            $packagepathspec = "$PSItem\*"
            $basfiles = (Get-ChildItem -Path $packagepathspec -Include *.bas, *.frm, *.cls)
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
        if (-not [System.IO.Path]::IsPathRooted($OutPath)) {
            $OutPath = Join-Path -Path $PWD -ChildPath $OutPath
        }
    
        ensureDirectory $OutPath
    
        if ($Generate -eq "Presentation") {
            $savefileFullName = (Join-Path -Path $OutPath -ChildPath $OutFileName)
            # 25 = ppSaveAsOpenXMLPresentationMacroEnabled
            $newPpt.SaveAs($savefileFullName, 25, $false)
        }
        else {
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

    ### End Create PowerPoint Document/AddIn file
}

function processPackageDir {
    param (
        [string]
        $packageDir,
        [xml]
        $mergedCuiXML
    )
    
    # Check whether there is a file called customUI14.xml in 
    # a subdirectory called CustomUI in the package directory
    $cuifilePath = Join-Path $packageDir "CustomUI\customUI14.xml"
    $cuiExists = Test-Path $cuifilePath
    if (-not $cuiExists) {
        Write-Output "No customUI in $cuifilePath"
        return
    }
    
    # Check if it a valid XML file
    try {
        $cuiXML = [xml] (Get-Content $cuifilePath)
    }
    catch {
        Write-Error "Could not read custom UI XML file $cuifilePath" -ErrorAction Stop
    }

    # Validate schema
    $cuins = @{cui = "http://schemas.microsoft.com/office/2009/07/customui" }
    $ribbonElement = $cuiXML | Select-Xml "/cui:customUI/cui:ribbon[@startFromScratch]" -Namespace $cuins
    if ($null -eq $ribbonElement ) {
        Write-Error "Wrong schema found in custom UI XML file $cuifilePath" -ErrorAction Stop
    }

    # Validate that there is only one ribbon element in the right place
    if ($ribbonElement.Count -ne 1) {
        Write-Error "There can be only one ribbon element." -ErrorAction Stop
    }

    # Validate that the ribbon element does not start from scratch
    if ($ribbonElement.Node.startFromScratch -ne "false") {
        Write-Error "Custom UIs are not allowed to have a from-scratch ribbon." -ErrorAction Stop
    }

    # Validate that all button image files are present 
    $buttons = $cuiXML | Select-Xml "/cui:customUI/cui:ribbon/cui:tabs//cui:button" -Namespace $cuins
    $buttons | ForEach-Object {
        $imagename = "$($PSItem.Node.image)"
        if ("" -eq $imagename) {
            return
        }
        $imagepath = Join-Path $packageDir "CustomUI\$imagename.png"
        if (-not (Test-Path $imagepath)) {
            Write-Error "Image $imagepath not present in $packageDir." -ErrorAction Stop
        }
    }

    # Debug
    # Write-Output "Validations passed."
    # if($null -eq $mergedCuiXML) {
    #     return
    # }
    # End Debug

    # Iterate tabs
    $tabs = ($cuiXML | Select-Xml "//cui:tabs/cui:tab" -Namespace $cuins)
    $tabs | ForEach-Object {
        # Write-Output "Processing tab '$($_.Node.idMso)'"
        processTab -currentTab $_.Node -mergedCuiXML $mergedCuiXML -cuins $cuins
    }
}

function processTab {
    param (
        $currentTab,
        [xml]
        $mergedCuiXML,
        $cuins
    )
    
    $currentTabQuery = "//cui:tabs/cui:tab[@idMso='$($currentTab.idMso)']"

    # If current tab does not exist in merged,
    # add it as is
    $existingTab = ($mergedCuiXML | Select-Xml $currentTabQuery -Namespace $cuins)
    if ($null -eq $existingTab) {
            
        $inode = $mergedCuiXML.ImportNode($currentTab, $true)
        $parentNode = ($mergedCuiXML | Select-Xml "//cui:ribbon/cui:tabs" -Namespace $cuins).Node
        [void] $parentNode.AppendChild($inode) # Do not put result in pipeline
            
        return
    }

    # If the tab is already there, iterate groups
    $groups = ($currentTab | Select-Xml "//cui:tab/cui:group" -Namespace $cuins)
    $groups | ForEach-Object {
        processTabGroup -currentGroup $_.Node -mergedCuiXML $mergedCuiXML -currentTabQuery $currentTabQuery -cuins $cuins
    }
}

function processTabGroup {
    param (
        $currentGroup,
        [xml]
        $mergedCuiXML,
        $currentTabQuery,
        $cuins
    )

    $currentGroupQuery = "$currentTabQuery/cui:group[@id='$($currentGroup.id)']"
            
    # if current group does not exist in merged.
    # add it as is
    $existingGroup = ($mergedCuiXML | Select-Xml $currentGroupQuery -Namespace $cuins)
    if ($null -eq $existingGroup) {

        $inode = $mergedCuiXML.ImportNode($currentGroup, $true)
        $parentnode = ($mergedCuiXML | Select-Xml $currentTabQuery -Namespace $cuins).Node
        [void] $parentnode.AppendChild($inode) # Do not put result in pipeline

        return    
    }

    # If group is already there, iterate buttons
    $buttons = ($currentGroup | Select-Xml "//cui:group/cui:button" -Namespace $cuins)
    $buttons | ForEach-Object {
        processButton -currentButton $_.Node -mergedCuiXML $mergedCuiXML -currentGroupQuery $currentGroupQuery -cuins $cuins
    }   
}

function processButton {
    param (
        $currentButton,
        [xml]
        $mergedCuiXML,
        $currentGroupQuery,
        $cuins
    )

    $currentButtonQuery = "$currentGroupQuery/cui:button[@id='$($currentButton.id)']"
    
    $existingbutton = ($mergedCuiXML | Select-Xml $currentButtonQuery -Namespace $cuins)
    
    if ($null -eq $existingbutton) {

        $inode = $mergedCuiXML.ImportNode($currentButton, $true)
        $parentnode = ($mergedCuiXML | Select-Xml $currentGroupQuery -Namespace $cuins).Node 
        [void] $parentnode.AppendChild($inode) # Do not put result in pipeline

        return    
    }

    # Current button should NOT exist in merged
    Write-Error "Duplicate button id: $($currentButton.id)" -ErrorAction Stop
}


Build-PPTFile $PackageNames -Generate $Generate -OutFileName $OutFileName -OutPath $OutPath -PackagesPath $PackagesPath