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
    $OutFileName = "MPMC",
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
        [void] (New-Item -ItemType Directory -Force -Path $OutPath)
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

    if($packagepaths.Count -eq 0) {
        Write-Output "No packages selected. Not doing anything"
        return
    }

    # Check output path and create if needed
    if (-not [System.IO.Path]::IsPathRooted($OutPath)) {
        $OutPath = Join-Path -Path $PWD -ChildPath $OutPath
    }
    
    ensureDirectory $OutPath

    
    Write-Output "Creating $Generate named $OutFileName..."
    createPPTFile -packagePaths $packagepaths -OutPath $OutPath -OutFileName $OutFileName -Generate $Generate

    Write-Output "Consolidating CustomUI for $Generate $OutFileName"
    consolidateCustomUI -packagePaths $packagepaths -OutPath $OutPath -OutFileName $OutFileName

    Write-Output "Done. For now, use a custom ui editor to insert the custom UI."
    # mergeUI -OutFileName $OutFileName -OutPath $OutPath -Generate $Generate
}

function mergeUI {
    param (
        $OutFileName,
        $OutPath,
        $Generate
    )
    
    if($Generate -eq "AddIn") {
        $extension = ".ppam"
    } else {
        $extension = ".pptm"
    }

    $pptfilename = Join-Path $OutPath "$($OutFileName)$extension"
    $pptzipfilename = "$pptfilename.zip"
    $expanddir = "$pptfilename.d"

    Remove-Item $pptzipfilename -Force -ErrorAction Ignore
    Remove-Item $expanddir -Recurse -Force -ErrorAction Ignore
    

    Move-Item $pptfilename $pptzipfilename
    Expand-Archive $pptzipfilename -DestinationPath $expanddir -ErrorAction Stop
    Remove-Item $pptzipfilename -Force -ErrorAction Ignore
    
    $cuiDir = Join-Path $OutPath "$OutFileName.UI"
    Move-Item $cuiDir -Destination "$expanddir\customUI"

    $relsFilePath = "$expanddir\_rels\.rels"
    $relsXML = [xml](Get-Content $relsFilePath)

    $relsnsurn = "http://schemas.openxmlformats.org/package/2006/relationships"
    $newNode = $relsXML.CreateElement("Relationship", $relsnsurn)
    $newNode.SetAttributeNode("Type", "").Value = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    $newNode.SetAttributeNode("Target", "").Value = "/customUI/customUI14.xml"
    $newNode.SetAttributeNode("Id", "").Value = "R" + [string](Get-Random)
    [void] $relsXML.DocumentElement.AppendChild($newNode)

    Set-Content $relsFilePath $relsXML.InnerXml
    
    Compress-Archive $expanddir $pptzipfilename -ErrorAction Stop

    Move-Item $pptzipfilename $pptfilename
}

function createPPTFile {
    param (
        $packagePaths,
        $OutPath,
        $OutFileName,
        $Generate
    )

    ### Create PowerPoint Document/AddIn file

    # Create PowerPoint object
    try {
        $ppa = New-Object -ComObject PowerPoint.Application
    
        $newPpt = $ppa.Presentations.Add($false)
    }
    catch {
        Write-Error "PowerPoint does not seem to be available." -ErrorAction Stop
    }
    
    try {
        # Try to get the VBA project object.
        # If VBA Object model access is not trusted, $vbp
        # will contain $null
        $vbp = $newPpt.VBProject
        
        if ($null -eq $vbp) {
            Write-Error "Access to VBA Object Model not trusted. Please check the Trust Access to the VBA Object model checkbox in the PowerPoint Trust Centre." -ErrorAction Stop
        }
    
        $packagePaths | ForEach-Object {
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

function consolidateCustomUI {
    param (
        $packagePaths,
        $OutPath,
        $OutFileName
    )

    $customUIOutPath = Join-Path $OutPath "$OutFileName.UI"
    if (Test-Path $customUIOutPath) {
        Remove-Item $customUIOutPath -Force -Recurse
    }
 
    ensureDirectory $customUIOutPath

    $customUIImagesPath = Join-Path $customUIOutPath "images"
    ensureDirectory $customUIImagesPath

    $customUIFilePath = Join-Path $customUIOutPath "customUI14.xml"

    $mcx = [xml]@"
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
<ribbon startFromScratch="false">
<tabs>
</tabs>
</ribbon>
</customUI>
"@

    $packagePaths | ForEach-Object {
        processPackageCustomUI -packageDir $_ -mergedCuiXML $mcx -outImagesDir $customUIImagesPath
    }

    $tabCount = ($mcx | Select-Xml "//cui:tabs/cui:tab" -Namespace @{cui = "http://schemas.microsoft.com/office/2009/07/customui" } ).Count
    if ($tabCount -gt 0) {
        $mcx.Save($customUIFilePath)
    }

    # processImageRels $mcx @{cui="http://schemas.microsoft.com/office/2009/07/customui"} $customUIOutPath
}

function  processImageRels {
    param (
        [xml]
        $mergedCuiXml,
        $cuins,
        $OutPath
    )
    
    $relsDir = Join-Path $OutPath "_rels"
    ensureDirectory $relsDir
    $relsFilePath = Join-Path $relsDir "customUI14.xml.rels"

    $relsnsurn = "http://schemas.openxmlformats.org/package/2006/relationships"
    
    $relsXML = [xml]"<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns='$relsnsurn'></Relationships>"
   
    $buttons = $mergedCuiXml | Select-Xml "//cui:button" -Namespace $cuins
    $buttons | ForEach-Object {
        $currentButton = $_.Node

        $newNode = $relsXML.CreateElement("Relationship", $relsnsurn)
        $newNode.SetAttributeNode("Type", "").Value = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        $newNode.SetAttributeNode("Target", "").Value = "images/$($currentButton.image).png"
        $newNode.SetAttributeNode("Id", "").Value = $currentButton.image
        [void] $relsXML.DocumentElement.AppendChild($newNode)
    }

    Set-Content $relsFilePath $relsXML.InnerXml
}

function processPackageCustomUI {
    param (
        [string]
        $packageDir,
        [xml]
        $mergedCuiXML,
        $outImagesDir
    )
    
    # Check whether there is a file called customUI14.xml in 
    # a subdirectory called CustomUI in the package directory
    $cuiDir = Join-Path $packageDir "CustomUI"
    $cuifilePath = Join-Path $cuiDir "customUI14.xml"
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
        $imagepath = Join-Path $cuiDir "$imagename.png"
        if (-not (Test-Path $imagepath)) {
            Write-Error "Image $imagepath not present in $packageDir." -ErrorAction Stop
        }
    }

    # Iterate tabs
    $tabs = ($cuiXML | Select-Xml "//cui:tabs/cui:tab" -Namespace $cuins)
    $tabs | ForEach-Object {
        # Write-Output "Processing tab '$($_.Node.idMso)'"
        processTab -currentTab $_.Node -mergedCuiXML $mergedCuiXML -cuins $cuins
    }

    # Copy images
    [void] (Copy-Item "$cuiDir\*png" -Destination $outImagesDir -Force)
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
