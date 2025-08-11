Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Folder Watcher"
$Form.Size = New-Object System.Drawing.Size(1000,700)

$CompanyLabel = New-Object System.Windows.Forms.Label
$CompanyLabel.Text = "Company Name:"
$CompanyLabel.Location = New-Object System.Drawing.Point(10,10)
$CompanyLabel.Size = New-Object System.Drawing.Size(100,20)
$Form.Controls.Add($CompanyLabel)

$CompanyTextBox = New-Object System.Windows.Forms.TextBox
$CompanyTextBox.Location = New-Object System.Drawing.Point(120,10)
$CompanyTextBox.Size = New-Object System.Drawing.Size(200,20)
$Form.Controls.Add($CompanyTextBox)

$IssueLabel = New-Object System.Windows.Forms.Label
$IssueLabel.Text = "Issue Number:"
$IssueLabel.Location = New-Object System.Drawing.Point(340,10) # <-- Move left to align
$IssueLabel.Size = New-Object System.Drawing.Size(90,20)
$Form.Controls.Add($IssueLabel)

$IssueTextBox = New-Object System.Windows.Forms.TextBox
$IssueTextBox.Location = New-Object System.Drawing.Point(440,10) # <-- Move left to align
$IssueTextBox.Size = New-Object System.Drawing.Size(100,20)
$Form.Controls.Add($IssueTextBox)

$FolderLabel = New-Object System.Windows.Forms.Label
$FolderLabel.Text = "Folder Path:"
$FolderLabel.Location = New-Object System.Drawing.Point(10,40)
$FolderLabel.Size = New-Object System.Drawing.Size(100,20)
$Form.Controls.Add($FolderLabel)

$FolderTextBox = New-Object System.Windows.Forms.TextBox
$FolderTextBox.Location = New-Object System.Drawing.Point(120,40)
$FolderTextBox.Size = New-Object System.Drawing.Size(400,20)
$Form.Controls.Add($FolderTextBox)

$InitButton = New-Object System.Windows.Forms.Button
$InitButton.Text = "Initialize JSON"
$InitButton.Location = New-Object System.Drawing.Point(550,40)
$InitButton.Size = New-Object System.Drawing.Size(120,20)
$Form.Controls.Add($InitButton)

$JsonLabel = New-Object System.Windows.Forms.Label
$JsonLabel.Text = "JSON Path:"
$JsonLabel.Location = New-Object System.Drawing.Point(550,10) # <-- Move right after Issue
$JsonLabel.Size = New-Object System.Drawing.Size(80,20)
$Form.Controls.Add($JsonLabel)

$JsonTextBox = New-Object System.Windows.Forms.TextBox
$JsonTextBox.Location = New-Object System.Drawing.Point(640,10) # <-- Move right after JSON label
$JsonTextBox.Size = New-Object System.Drawing.Size(300,20)
$Form.Controls.Add($JsonTextBox)

$PickJsonButton = New-Object System.Windows.Forms.Button
$PickJsonButton.Text = "Pick JSON"
$PickJsonButton.Location = New-Object System.Drawing.Point(680,40)
$PickJsonButton.Size = New-Object System.Drawing.Size(100,20)
$Form.Controls.Add($PickJsonButton)

$LoadJsonButton = New-Object System.Windows.Forms.Button
$LoadJsonButton.Text = "Load JSON"
$LoadJsonButton.Location = New-Object System.Drawing.Point(800,40)
$LoadJsonButton.Size = New-Object System.Drawing.Size(100,20)
$Form.Controls.Add($LoadJsonButton)

$TreeView = New-Object System.Windows.Forms.TreeView
$TreeView.Location = New-Object System.Drawing.Point(10, 70)
$TreeView.Size = New-Object System.Drawing.Size(960, 550)
$TreeView.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$Form.Controls.Add($TreeView)

$RefreshButton = New-Object System.Windows.Forms.Button
$RefreshButton.Location = New-Object System.Drawing.Point(920, 40)
$RefreshButton.Size = New-Object System.Drawing.Size(50, 20)
$RefreshButton.Text = "Refresh"
$Form.Controls.Add($RefreshButton)

# Global variables for JSON path and loaded state
$Global:LoadedJsonPath = $null
$Global:LoadedState = $null

# Helper Functions
function Test-JsonProperty {
    param($obj, $property)
    return $obj.PSObject.Properties.Name -contains $property
}

function ConvertToHashtable($obj) {
    Write-Host "[ConvertToHashtable] Input type: $($obj.GetType().FullName)"
    if ($obj -is [System.Collections.Hashtable]) {
        Write-Host "[ConvertToHashtable] Already hashtable."
        return $obj
    } elseif ($obj -is [PSCustomObject]) {
        $ht = @{}
        foreach ($prop in $obj.PSObject.Properties) {
            $ht[$prop.Name] = ConvertToHashtable $prop.Value
        }
        Write-Host "[ConvertToHashtable] Converted PSCustomObject to hashtable. Keys: $($ht.Keys -join ', ')"
        return $ht
    } elseif ($obj -is [System.Collections.IDictionary]) {
        $ht = @{}
        foreach ($key in $obj.Keys) {
            $ht[$key] = ConvertToHashtable $obj[$key]
        }
        Write-Host "[ConvertToHashtable] Converted IDictionary to hashtable. Keys: $($ht.Keys -join ', ')"
        return $ht
    } elseif ($obj -is [System.Collections.IEnumerable] -and !$obj -is [string]) {
        $arr = @($obj | ForEach-Object { ConvertToHashtable $_ })
        Write-Host "[ConvertToHashtable] Converted IEnumerable to array of $($arr.Count) items."
        return $arr
    } else {
        return $obj
    }
}

function DeepConvertState($state) {
    Write-Host "[DeepConvertState] Converting top-level state..."
    $state = ConvertToHashtable $state
    if ($state.ContainsKey('folders')) {
        $folders = @{}
        foreach ($folderKey in $state['folders'].Keys) {
            $folders[$folderKey] = ConvertToHashtable $state['folders'][$folderKey]
            Write-Host "[DeepConvertState] Folder '$folderKey' type: $($folders[$folderKey].GetType().FullName)"
            if ($folders[$folderKey].ContainsKey('Files')) {
                $folders[$folderKey]['Files'] = ConvertToHashtable $folders[$folderKey]['Files']
                Write-Host "[DeepConvertState] Folder '$folderKey' Files type: $($folders[$folderKey]['Files'].GetType().FullName)"
            }
        }
        $state['folders'] = $folders
        Write-Host "[DeepConvertState] Final folders keys: $($state['folders'].Keys -join ', ')"
    }
    if ($state.ContainsKey('files')) {
        $files = @{}
        foreach ($fileKey in $state['files'].Keys) {
            $files[$fileKey] = ConvertToHashtable $state['files'][$fileKey]
            Write-Host "[DeepConvertState] File '$fileKey' type: $($files[$fileKey].GetType().FullName)"
        }
        $state['files'] = $files
        Write-Host "[DeepConvertState] Final files keys: $($state['files'].Keys -join ', ')"
    }
    return $state
}

function UpdateFolderPath {
    $script:basePath = "G:\Shared drives\TriMech Solutions\"
    $company = $CompanyTextBox.Text.Trim()
    $issue = $IssueTextBox.Text.Trim()
    
    # Input validation
    if (-not $company -or $company.Length -lt 1) {
        $FolderTextBox.Text = ""
        $JsonTextBox.Text = ""
        return
    }
    
    $firstLetter = $company.Substring(0,1).ToUpper()
    $rest = if ($company.Length -gt 1) { $company.Substring(1) } else { "" }
    $script:companyFolder = "$firstLetter$rest"
    
    # Build folder path
    if ($script:companyFolder -eq "Halyard") {
        $folderPath = Join-Path $script:basePath "H"
        $folderPath = Join-Path $folderPath "Halyard"
        if ($issue) {
            $folderPath = Join-Path $folderPath $issue
        }
    } else {
        $folderPath = $script:basePath
        $folderPath = Join-Path $folderPath $firstLetter
        $folderPath = Join-Path $folderPath $script:companyFolder
        if ($issue) {
            $folderPath = Join-Path $folderPath $issue
        }
    }
    $FolderTextBox.Text = $folderPath
    $JsonTextBox.Text = Join-Path $PWD "$company.json"
}

# Core Functions
function Initialize-State {
    param($folderPath, $jsonPath)
    
    Write-Host "=== Starting Initialize-State ==="
    Write-Host "Folder path: $folderPath"
    Write-Host "JSON path: $jsonPath"
    
    # Create state object with proper type declarations
    $state = [PSCustomObject]@{
        rootPath = $folderPath
        folders = [System.Collections.Hashtable]::new()
        files = [System.Collections.Hashtable]::new()
        lastCheck = [DateTime]::UtcNow.ToString("o")
        isInitialized = $true
    }
    
    if (-not (Test-Path $folderPath)) {
        Write-Host "ERROR: Folder path does not exist"
        return [PSCustomObject]@{
            rootPath = $folderPath
            folders = [System.Collections.Hashtable]@{}
            files = [System.Collections.Hashtable]@{}
            lastCheck = (Get-Date).ToString("o")
            isInitialized = $false
        }
    }

    # Get all items (folders and files) in the directory
    Write-Host "Getting items from directory..."
    $allItems = Get-ChildItem -Path $folderPath -Recurse
    
    $state = [PSCustomObject]@{
        rootPath = $folderPath
        folders = [System.Collections.Hashtable]@{}
        files = [System.Collections.Hashtable]@{}
        lastCheck = (Get-Date).ToString("o")
        isInitialized = $true
    }

    # First, add all folders
    $allItems | Where-Object { $_.PSIsContainer } | ForEach-Object {
        $relativePath = $_.FullName.Substring($folderPath.Length).TrimStart('\')
        if ($relativePath -ne '') {
            $lastMod = $_.LastWriteTimeUtc.ToString("o")
            $state.folders[$relativePath] = @{
                path = $relativePath
                creation_date = $_.CreationTimeUtc.ToString("o")
                initial_modified = $lastMod
                current_modified = $lastMod
                status = "red"  # Initial status is red
                hasChanges = $false
            }
        }
    }

    # Add all files
    $allItems | Where-Object { -not $_.PSIsContainer } | ForEach-Object {
        $relativePath = $_.FullName.Substring($folderPath.Length).TrimStart('\')
        $lastMod = $_.LastWriteTimeUtc.ToString("o")
        $size = $_.Length
        $state.files[$relativePath] = @{
            path = $relativePath
            name = $_.Name
            creation_date = $_.CreationTimeUtc.ToString("o")
            initial_modified = $lastMod
            current_modified = $lastMod
            initial_size = $size
            current_size = $size
            status = "red"  # Initial status is red
            hasChanges = $false
        }
    }

    # Add the root folder if there are no subfolders
    if ($state.folders.Count -eq 0) {
        $rootInfo = Get-Item $folderPath
        $state.folders['.'] = @{
            path = '.'
            creation_date = $rootInfo.CreationTimeUtc.ToString("o")
            last_modified = $rootInfo.LastWriteTimeUtc.ToString("o")
            status = "red"  # Initial status is red
            hasChanges = $false
        }
    }
    
    # Ensure the folder exists
    $folder = Split-Path $jsonPath -Parent
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    # Save the state with explicit encoding
    try {
        $json = $state | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($jsonPath, $json)
        Write-Host "Successfully saved state to $jsonPath"
    } catch {
        Write-Host "Error saving state: $_"
    }
    
    return $state
}

function Load-State {
    param($jsonFileName)
    Write-Host "=== Starting Load-State ==="
    Write-Host "Loading state from: $jsonFileName"
    
    if ([string]::IsNullOrWhiteSpace($jsonFileName)) {
        Write-Host "ERROR: No JSON filename provided"
        return $null
    }
    
    $jsonPath = Join-Path $PWD $jsonFileName
    Write-Host "Constructed full JSON path: $jsonPath"
    Write-Host "File exists: $(Test-Path $jsonPath)"
    Write-Host "Is file (not directory): $(Test-Path $jsonPath -PathType Leaf)"
    
    if (Test-Path $jsonPath -PathType Leaf) {
        try {
            Write-Host "Reading JSON content..."
            $jsonContent = Get-Content $jsonPath -Raw
            Write-Host "JSON content length: $($jsonContent.Length)"
            Write-Host "First 100 chars of JSON: $($jsonContent.Substring(0, [Math]::Min(100, $jsonContent.Length)))"
            
            Write-Host "Converting JSON to object..."
            $state = $jsonContent | ConvertFrom-Json
            
            # Convert to proper types
            if ($state.folders -is [PSCustomObject]) {
                $foldersHash = [System.Collections.Hashtable]::new()
                $state.folders.PSObject.Properties | ForEach-Object {
                    $foldersHash[$_.Name] = $_.Value
                }
                $state.folders = $foldersHash
            }
            if ($state.files -is [PSCustomObject]) {
                $filesHash = [System.Collections.Hashtable]::new()
                $state.files.PSObject.Properties | ForEach-Object {
                    $filesHash[$_.Name] = $_.Value
                }
                $state.files = $filesHash
            }
            Write-Host "State object created and converted to proper types"
            Write-Host "Raw state properties: $($state.PSObject.Properties.Name -join ', ')"
            
            # Convert folders and files to hashtables
            if ($state.folders -is [PSCustomObject]) {
                $foldersHash = @{}
                $state.folders.PSObject.Properties | ForEach-Object {
                    $foldersHash[$_.Name] = $_.Value
                }
                $state.folders = $foldersHash
            }
            
            if ($state.files -is [PSCustomObject]) {
                $filesHash = @{}
                $state.files.PSObject.Properties | ForEach-Object {
                    $filesHash[$_.Name] = $_.Value
                }
                $state.files = $filesHash
            }
            
            # Validate state object structure
            $hasValidStructure = $true
            if (-not (Test-JsonProperty $state 'folders')) {
                Write-Host "ERROR: JSON missing 'folders' property"
                $hasValidStructure = $false
            }
            if (-not (Test-JsonProperty $state 'files')) {
                Write-Host "ERROR: JSON missing 'files' property"
                $hasValidStructure = $false
            }
            
            if ($hasValidStructure) {
                Write-Host "Folders in JSON: $($state.folders.PSObject.Properties.Name -join ', ')"
                Write-Host "Files in JSON: $($state.files.PSObject.Properties.Name -join ', ')"
            } else {
                Write-Host "Full JSON content for debugging:"
                Write-Host $jsonContent
                return $null
            }
            
            return $state
        }
        catch {
            Write-Host "ERROR loading JSON: $_"
            Write-Host "Error details: $($_.Exception.Message)"
            return $null
        }
    } else {
        Write-Host "ERROR: JSON file not found at: $jsonPath"
        return $null
    }
    Write-Host "=== End Load-State ==="
}

function Update-State($state) {
    if ($null -eq $state) {
        Write-Host "ERROR: Cannot update null state"
        return $null
    }

    $folderPath = $FolderTextBox.Text
    if (-not (Test-Path $folderPath)) {
        [System.Windows.Forms.MessageBox]::Show("Cannot update state: Folder not found: $folderPath")
        return $state
    }

    # Make a deep copy of the original state for comparison
    $originalState = $state | ConvertTo-Json -Depth 10 | ConvertFrom-Json

    # Ensure state has proper structure
    if ($null -eq $state.folders) {
        $state.folders = [System.Collections.Hashtable]@{}
    }
    if ($null -eq $state.files) {
        $state.files = [System.Collections.Hashtable]@{}
    }

    # Get current state of all items
    $currentItems = Get-ChildItem -Path $folderPath -Recurse
    
    Write-Host "=== Starting Update-State Check ==="
    # Check root folder
    $rootInfo = Get-Item $folderPath
    $rootModified = $rootInfo.LastWriteTimeUtc.ToString("o")
    
    if ($state.folders -and $state.folders.ContainsKey('.')) {
        Write-Host "Root folder initial time: $($state.folders['.'].initial_modified)"
        Write-Host "Root folder current time: $rootModified"
        
        try {
            $initialTime = [DateTime]::Parse($state.folders['.'].initial_modified)
            $currentTime = [DateTime]::Parse($rootModified)
            $hasChanged = $initialTime -ne $currentTime
            
            Write-Host "Root folder has changed: $hasChanged"
            
            if ($hasChanged) {
                $state.folders['.'].current_modified = $rootModified
                $state.folders['.'].status = "green"
                $state.folders['.'].hasChanges = $true
            } else {
                $state.folders['.'].status = "red"
                $state.folders['.'].hasChanges = $false
            }
        } catch {
            Write-Host "Error comparing root folder times: $_"
        }
    }

    # Check folders
    foreach ($folder in $currentItems | Where-Object { $_.PSIsContainer }) {
        $relativePath = $folder.FullName.Substring($folderPath.Length).TrimStart('\')
        $lastMod = $folder.LastWriteTimeUtc.ToString("o")
        
        Write-Host "Checking folder: $relativePath"
        
        if ($state.folders -and $state.folders.ContainsKey($relativePath)) {
            # Existing folder - check for changes
            Write-Host "  Initial time: $($state.folders[$relativePath].initial_modified)"
            Write-Host "  Current time: $lastMod"
            
            try {
                $initialTime = [DateTime]::Parse($state.folders[$relativePath].initial_modified)
                $currentTime = [DateTime]::Parse($lastMod)
                $hasChanged = $initialTime -ne $currentTime
                
                Write-Host "  Has changed: $hasChanged"
                
                $state.folders[$relativePath].current_modified = $lastMod
                if ($hasChanged) {
                    $state.folders[$relativePath].status = "green"
                    $state.folders[$relativePath].hasChanges = $true
                } else {
                    $state.folders[$relativePath].status = "red"
                    $state.folders[$relativePath].hasChanges = $false
                }
            } catch {
                Write-Host "  Error comparing times: $_"
            }
        } else {
            Write-Host "  New folder found"
            # New folder found
            $state.folders[$relativePath] = @{
                path = $relativePath
                creation_date = $folder.CreationTimeUtc.ToString("o")
                initial_modified = $lastMod
                current_modified = $lastMod
                status = "red"
                hasChanges = $false
            }
        }
    }

    # Check files
    foreach ($file in $currentItems | Where-Object { -not $_.PSIsContainer }) {
        $relativePath = $file.FullName.Substring($folderPath.Length).TrimStart('\')
        $lastMod = $file.LastWriteTimeUtc.ToString("o")
        $currentSize = $file.Length
        
        Write-Host "Checking file: $relativePath"
        
        if ($state.files.ContainsKey($relativePath)) {
            # Existing file - check for changes
            Write-Host "  Initial time: $($state.files[$relativePath].initial_modified)"
            Write-Host "  Initial size: $($state.files[$relativePath].initial_size)"
            Write-Host "  Current time: $lastMod"
            Write-Host "  Current size: $currentSize"
            
            try {
                $initialTime = [DateTime]::Parse($state.files[$relativePath].initial_modified)
                $currentTime = [DateTime]::Parse($lastMod)
                $hasChanged = ($initialTime -ne $currentTime) -or ($state.files[$relativePath].initial_size -ne $currentSize)
                
                Write-Host "  Has changed: $hasChanged"
                
                $state.files[$relativePath].current_modified = $lastMod
                $state.files[$relativePath].current_size = $currentSize
                if ($hasChanged) {
                    $state.files[$relativePath].status = "green"
                    $state.files[$relativePath].hasChanges = $true
                } else {
                    $state.files[$relativePath].status = "red"
                    $state.files[$relativePath].hasChanges = $false
                }
            } catch {
                Write-Host "  Error comparing times: $_"
            }
        } else {
            Write-Host "  New file found"
            # New file found
            $state.files[$relativePath] = @{
                path = $relativePath
                name = $file.Name
                creation_date = $file.CreationTimeUtc.ToString("o")
                initial_modified = $lastMod
                current_modified = $lastMod
                initial_size = $currentSize
                current_size = $currentSize
                status = "red"
                hasChanges = $false
            }
        }
    }

    # Detect missing folders (previously captured but now gone)
    foreach ($savedKey in $state.folders.Keys) {
        if (-not ($currentItems | Where-Object { $_.PSIsContainer -and $_.FullName.Substring($folderPath.Length).TrimStart('\') -eq $savedKey })) {
            Write-Host "MISSING FOLDER: $savedKey"
            if ($state.folders[$savedKey] -is [hashtable]) {
                $state.folders[$savedKey].status = "missing"
            } elseif ($state.folders[$savedKey] -is [PSCustomObject] -and $state.folders[$savedKey].GetType().Name -ne 'Hashtable') {
                $state.folders[$savedKey] = @{} + $state.folders[$savedKey]
                $state.folders[$savedKey].status = "missing"
            }
        }
    }

    # Detect missing files (previously captured but now gone)
    foreach ($savedKey in $state.files.Keys) {
        if (-not ($currentItems | Where-Object { -not $_.PSIsContainer -and $_.FullName.Substring($folderPath.Length).TrimStart('\') -eq $savedKey })) {
            Write-Host "MISSING FILE: $savedKey"
            if ($state.files[$savedKey] -is [hashtable]) {
                $state.files[$savedKey].status = "missing"
            } elseif ($state.files[$savedKey] -is [PSCustomObject] -and $state.files[$savedKey].GetType().Name -ne 'Hashtable') {
                $state.files[$savedKey] = @{} + $state.files[$savedKey]
                $state.files[$savedKey].status = "missing"
            }
        }
    }

    # Compare the current state with original state to detect real changes
    $hasRealChanges = $false
    
    # Compare folders
    foreach ($folderPath in $state.folders.Keys) {
        $currentFolder = $state.folders[$folderPath]
        $originalFolder = $originalState.folders.$folderPath
        
        if ($originalFolder) {
            # Parse dates to ensure consistent comparison
            try {
                $originalTime = [DateTime]::Parse($originalFolder.current_modified)
                $currentTime = [DateTime]::Parse($currentFolder.current_modified)
                
                if ($originalTime -ne $currentTime) {
                    $hasRealChanges = $true
                    $currentFolder.hasChanges = $true
                    $currentFolder.status = "green"
                    Write-Host "CHANGE DETECTED - Folder: $folderPath"
                    Write-Host "  Original Modified: $($originalTime.ToString('o'))"
                    Write-Host "  Current Modified:  $($currentTime.ToString('o'))"
                } else {
                    $currentFolder.hasChanges = $false
                    $currentFolder.status = "red"
                }
            } catch {
                Write-Host "Error comparing dates for folder $folderPath : $_"
            }
        }
    }
    
    # Compare files
    foreach ($filePath in $state.files.Keys) {
        $currentFile = $state.files[$filePath]
        $originalFile = $originalState.files.$filePath
        
        if ($originalFile) {
            try {
                # Parse dates to DateTime objects for consistent comparison
                $originalTime = [DateTime]::Parse($originalFile.current_modified)
                $currentTime = [DateTime]::Parse($currentFile.current_modified)
                
                $timeChanged = $originalTime -ne $currentTime
                $sizeChanged = $currentFile.current_size -ne $originalFile.current_size
                
                if ($timeChanged -or $sizeChanged) {
                    $hasRealChanges = $true
                    $currentFile.hasChanges = $true
                    $currentFile.status = "green"
                    Write-Host "CHANGE DETECTED - File: $filePath"
                    if ($timeChanged) {
                        # Use consistent ISO 8601 format for both dates
                        Write-Host "  Original Modified: $($originalTime.ToString('o'))"
                        Write-Host "  Current Modified:  $($currentTime.ToString('o'))"
                    }
                    if ($sizeChanged) {
                        Write-Host "  Original Size: $($originalFile.current_size) bytes"
                        Write-Host "  Current Size:  $($currentFile.current_size) bytes"
                    }
                } else {
                    $currentFile.hasChanges = $false
                    $currentFile.status = "red"
                }
            } catch {
                Write-Host "Error comparing dates for file $filePath : $_"
            }
        }
    }
    
    # Update lastCheck without triggering a save just for timestamp
    $state.lastCheck = (Get-Date).ToString("o")
    
    Write-Host "Has real changes to save: $hasRealChanges"
    
    if ($hasRealChanges) {
        $jsonFileName = [System.IO.Path]::GetFileName($JsonTextBox.Text)
        $jsonPath = Join-Path $PWD $jsonFileName
        Write-Host "Saving state to: $jsonPath"
        $state | ConvertTo-Json -Depth 5 | Set-Content $jsonPath
    } else {
        Write-Host "No real changes detected, skipping save"
    }
    
    return $state
}

function Build-Tree($state) {
    Write-Host "=== Starting Build-Tree ==="
    Write-Host "Building tree with state properties: $($state.PSObject.Properties.Name -join ', ')"
    
    if ($null -eq $state) {
        Write-Host "ERROR: State is null"
        return
    }
    
    if (-not (Test-JsonProperty $state 'folders')) {
        Write-Host "ERROR: State missing 'folders' property"
        Write-Host "Available properties: $($state.PSObject.Properties.Name -join ', ')"
        return
    }
    
    Write-Host "Folders in state: $($state.folders.PSObject.Properties.Name -join ', ')"
    $TreeView.Nodes.Clear()

    # Add root folder
    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Name = "root"
    if ($state.rootPath) {
        $rootNode.Text = Split-Path $state.rootPath -Leaf
    } else {
        $rootNode.Text = "Root"
    }
    if ($state.folders.ContainsKey('.')) {
        $rootNode.ForeColor = if ($state.folders['.'].status -eq "green") { 'Green' } else { 'Red' }
    }
    $TreeView.Nodes.Add($rootNode)

    # Create a hashtable to store all nodes for easy parent lookup
    $nodeMap = @{}
    $nodeMap['.'] = $rootNode

    # Add all subfolders first
    foreach ($folderKey in $state.folders.Keys | Where-Object { $_ -ne '.' } | Sort-Object) {
        $folder = $state.folders[$folderKey]
        $parts = $folder.path.Split('\')
        $parent = $rootNode
        $currentPath = ''
        for ($i = 0; $i -lt $parts.Length; $i++) {
            $part = $parts[$i]
            $currentPath = if ($currentPath) { "$currentPath\$part" } else { $part }
            if (-not $nodeMap.ContainsKey($currentPath)) {
                $node = New-Object System.Windows.Forms.TreeNode
                $node.Name = $part
                # Strikethrough for missing
                if ($folder.status -eq "missing") {
                    $node.NodeFont = New-Object System.Drawing.Font($TreeView.Font, [System.Drawing.FontStyle]::Strikeout)
                    $node.ForeColor = 'Gray'
                } else {
                    $node.ForeColor = if ($folder.status -eq "green") { 'Green' } elseif ($folder.status -eq "red") { 'Red' } else { 'Black' }
                }
                # If this is the last part of the path, set the color based on status
                if ($i -eq $parts.Length - 1) {
                    try {
                        $lastMod = [DateTime]::Parse($folder.current_modified).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                    } catch {
                        $lastMod = "Unknown"
                        Write-Host "Error parsing date: $($folder.current_modified)"
                    }
                    $node.Text = "$part (Modified: $lastMod)"
                } else {
                    $node.Text = $part
                }
                $parent.Nodes.Add($node)
                $nodeMap[$currentPath] = $node
                $parent = $node
            } else {
                $parent = $nodeMap[$currentPath]
            }
        }
    }

    # Add files under their respective folders
    if ($state.files) {
        foreach ($fileKey in $state.files.Keys | Sort-Object) {
            $file = $state.files[$fileKey]
            $folderPath = Split-Path $file.path -Parent
            $fileName = Split-Path $file.path -Leaf
            $parentNode = if ($folderPath) { $nodeMap[$folderPath] } else { $rootNode }
            $fileNode = New-Object System.Windows.Forms.TreeNode
            $fileNode.Name = $fileName
            # Strikethrough for missing
            if ($file.status -eq "missing") {
                $fileNode.NodeFont = New-Object System.Drawing.Font($TreeView.Font, [System.Drawing.FontStyle]::Strikeout)
                $fileNode.ForeColor = 'Gray'
            } else {
                $fileNode.ForeColor = if ($file.status -eq "green") { 'Green' } elseif ($file.status -eq "red") { 'Red' } else { 'Black' }
            }
            try {
                $lastMod = [DateTime]::Parse($file.current_modified).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
            } catch {
                $lastMod = "Unknown"
                Write-Host "Error parsing date: $($file.current_modified)"
            }
            $size = [math]::Round($file.current_size / 1KB, 2)
            $fileNode.Text = "$fileName (Modified: $lastMod, Size: $size KB)"
            $parentNode.Nodes.Add($fileNode)
        }
    }
    $rootNode.Expand()
}

function Show-FolderTree {
    param (
        [string]$Path,
        [int]$Indent = 0,
        [switch]$ClearFoundFiles = $false
    )
    
    if ($ClearFoundFiles) {
        $script:foundFiles = @()
    }

    if (-not (Test-Path $Path)) {
        return "Path not found: $Path`n"
    }
    
    $items = Get-ChildItem -Path $Path -ErrorAction SilentlyContinue
    if (-not $items) {
        return "Empty folder`n"
    }
    
    $result = ""
    foreach ($item in $items) {
        $prefix = " " * $Indent
        if ($item.PSIsContainer) {
            $result += "$prefix[DIR] $($item.Name)`n"
            $result += Show-FolderTree -Path $item.FullName -Indent ($Indent + 2)
        } else {
            $result += "$prefix[FILE] $($item.Name)`n"
            $script:foundFiles += $item
        }
    }
    return $result
}

# Event Handlers
$CompanyTextBox.Add_TextChanged({ UpdateFolderPath })
$IssueTextBox.Add_TextChanged({ UpdateFolderPath })

$PickJsonButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $PWD
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Only set the filename, not the full path, to ensure loading uses script directory
        $JsonTextBox.Text = [System.IO.Path]::GetFileName($dialog.FileName)
    }
})

$InitButton.Add_Click({
    $company = $CompanyTextBox.Text.Trim()
    if (-not $company) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a company name.")
        return
    }

    # Setup paths
    $folderPath = $FolderTextBox.Text
    $jsonPath = Join-Path $PWD "$company.json"
    $JsonTextBox.Text = $jsonPath

    try {
        Write-Host "Checking for existing JSON at: $jsonPath"
        if (Test-Path $jsonPath -PathType Leaf) {
            Write-Host "Found existing JSON, loading it..."
            $state = Load-State "$company.json"
            if ($state) {
                Write-Host "Loaded existing JSON successfully"
                Build-Tree $state
                [System.Windows.Forms.MessageBox]::Show(
                    "Loaded existing JSON with $($state.folders.Count) folders and $($state.files.Count) files",
                    "Load Complete"
                )
                return
            }
        }

        Write-Host "No existing JSON found or load failed, initializing new state..."
        # Ensure script directory exists
        if (-not (Test-Path $PSScriptRoot)) {
            New-Item -Path $PSScriptRoot -ItemType Directory -Force | Out-Null
        }

        # Handle non-existent folder
        if (-not (Test-Path $folderPath)) {
            $parentPath = Split-Path $folderPath -Parent
            $structure = Show-FolderTree -Path $parentPath
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Folder does not exist: $folderPath`n`nDo you want to create it?",
                "Folder Not Found",
                [System.Windows.Forms.MessageBoxButtons]::YesNo
            )
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
            } else {
                return
            }
        }

        # Initialize new state
        $state = Initialize-State $folderPath $jsonPath
        if ($null -ne $state -and ($state.folders.Count -gt 0 -or $state.files.Count -gt 0)) {
            Write-Host "Successfully initialized new state"
            Write-Host "State contains folders: $($state.folders.Count)"
            Write-Host "State contains files: $($state.files.Count)"
            Build-Tree $state
            [System.Windows.Forms.MessageBox]::Show(
                "Initialized JSON at: $jsonPath`nTracking $($state.folders.Count) folders and $($state.files.Count) files",
                "Initialization Complete"
            )
        } else {
            Write-Host "No folders or files found to initialize"
            [System.Windows.Forms.MessageBox]::Show(
                "No folders or files found to initialize.",
                "Initialization Complete"
            )
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_")
    }
})

# --- BEGIN: Fix parent branch marking in LoadJsonButton handler ---
$LoadJsonButton.Add_Click({
    try {
        $jsonFileName = [System.IO.Path]::GetFileName($JsonTextBox.Text)
        Write-Host "Loading JSON: $jsonFileName"
        if ([string]::IsNullOrWhiteSpace($jsonFileName)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify a JSON file name to load.")
            return
        }
        $jsonPath = Join-Path $PWD $jsonFileName
        $state = Get-Content $jsonPath -Raw | ConvertFrom-Json
        Write-Host "State loaded, checking content..."
        # Ensure folders and files are hashtables
        if ($state.folders -is [PSCustomObject]) {
            $foldersHash = @{}
            foreach ($prop in $state.folders.PSObject.Properties) {
                $foldersHash[$prop.Name] = $prop.Value
            }
            $state.folders = $foldersHash
        }
        if ($state.files -is [PSCustomObject]) {
            $filesHash = @{}
            foreach ($prop in $state.files.PSObject.Properties) {
                $filesHash[$prop.Name] = $prop.Value
            }
            $state.files = $filesHash
        }
        # Get current folder path
        $folderPath = $FolderTextBox.Text
        if (-not (Test-Path $folderPath)) {
            [System.Windows.Forms.MessageBox]::Show("Cannot compare: Folder not found: $folderPath")
            return
        }
        $currentItems = Get-ChildItem -Path $folderPath -Recurse
        $changedFolders = @{}
        $changedFiles = @{}
        # Compare folders
        foreach ($folder in $currentItems | Where-Object { $_.PSIsContainer }) {
            $relativePath = $folder.FullName.Substring($folderPath.Length).TrimStart('\')
            Write-Host "Comparing folder: $relativePath"
            if ($relativePath -ne '') {
                if ($state.folders.ContainsKey($relativePath)) {
                    Write-Host "Indexing into folders for key: $relativePath"
                    if ($state.folders[$relativePath] -isnot [hashtable] -and $state.folders[$relativePath] -is [PSCustomObject]) {
                        $state.folders[$relativePath] = @{} + $state.folders[$relativePath]
                    }
                    $state.folders[$relativePath].status = $state.folders[$relativePath].status
                } else {
                    Write-Host "Key not found in folders: $relativePath"
                    $state.folders[$relativePath] = @{
                        path = $relativePath
                        creation_date = $folder.CreationTimeUtc.ToString('o')
                        initial_modified = $folder.LastWriteTimeUtc.ToString('o')
                        current_modified = $folder.LastWriteTimeUtc.ToString('o')
                        status = "red"
                        hasChanges = $true
                    }
                    $changedFolders[$relativePath] = $true
                }
            }
        }
        # Compare files
        foreach ($file in $currentItems | Where-Object { -not $_.PSIsContainer }) {
            $relativePath = $file.FullName.Substring($folderPath.Length).TrimStart('\')
            Write-Host "Comparing file: $relativePath"
            if ($state.files.ContainsKey($relativePath)) {
                Write-Host "Indexing into files for key: $relativePath"
                if ($state.files[$relativePath] -isnot [hashtable] -and $state.files[$relativePath] -is [PSCustomObject]) {
                    $state.files[$relativePath] = @{} + $state.files[$relativePath]
                }
                $initialMod = $state.files[$relativePath].initial_modified
                $currentMod = $file.LastWriteTimeUtc.ToString('o')
                if ($initialMod -ne $currentMod) {
                    $state.files[$relativePath].status = "green"
                } else {
                    $state.files[$relativePath].status = $state.files[$relativePath].status
                }
            } else {
                Write-Host "Key not found in files: $relativePath"
                $state.files[$relativePath] = @{
                    path = $relativePath
                    name = $file.Name
                    creation_date = $file.CreationTimeUtc.ToString('o')
                    initial_modified = $file.LastWriteTimeUtc.ToString('o')
                    current_modified = $file.LastWriteTimeUtc.ToString('o')
                    initial_size = $file.Length
                    current_size = $file.Length
                    status = "red"
                    hasChanges = $true
                }
                $changedFiles[$relativePath] = $true
            }
        }
        # Mark only direct parent branch green if any child changed
        foreach ($folderKey in $changedFolders.Keys) {
            $parentBranch = Split-Path $folderKey -Parent
            if ($parentBranch -and $state.folders.ContainsKey($parentBranch) -and $state.folders[$parentBranch].status -ne "green") {
                $state.folders[$parentBranch].status = "green"
            }
        }
        foreach ($fileKey in $changedFiles.Keys) {
            $parent = Split-Path $fileKey -Parent
            if ($parent -and $state.folders[$parent]) {
                $state.folders[$parent].status = "green"
            }
        }
        Write-Host "Building tree with loaded and compared state..."
        Build-Tree $state
        Write-Host "Tree built successfully"
        $folderCount = if ($state.folders) { $state.folders.Count } else { 0 }
        $fileCount = if ($state.files) { $state.files.Count } else { 0 }
        [System.Windows.Forms.MessageBox]::Show(
            "Loaded and compared JSON successfully`nFolders: $folderCount`nFiles: $fileCount",
            "Load & Compare Complete"
        )
    } catch {
        Write-Host "Error in LoadJsonButton: $_"
        [System.Windows.Forms.MessageBox]::Show("Error loading JSON: $_")
    }
})
# --- END: Fix parent branch marking ---

$RefreshButton.Add_Click({
    $jsonFileName = [System.IO.Path]::GetFileName($JsonTextBox.Text)
    Write-Host "=== Refresh Button Clicked ==="
    Write-Host "Refreshing from file: $jsonFileName"
    if ([string]::IsNullOrWhiteSpace($jsonFileName)) {
        [System.Windows.Forms.MessageBox]::Show("No JSON file selected.")
        return
    }

    $jsonPath = Join-Path $PWD $jsonFileName
    Write-Host "Full JSON path: $jsonPath"
    Write-Host "JSON file exists: $(Test-Path $jsonPath)"

    try {
        $savedState = Get-Content $jsonPath -Raw | ConvertFrom-Json
        $savedState = DeepConvertState $savedState
        $savedState = Ensure-StateObject $savedState
        Write-Host "Loaded saved state from JSON"

        $folderPath = $FolderTextBox.Text
        if (-not (Test-Path $folderPath)) {
            [System.Windows.Forms.MessageBox]::Show("Cannot refresh: Folder not found: $folderPath")
            return
        }

        # Get current state
        $currentState = Initialize-State $folderPath $jsonPath
        $currentState = DeepConvertState $currentState
        $currentState = Ensure-StateObject $currentState

        # Set all folders/files to red by default
        foreach ($key in $currentState.folders.Keys) { $currentState.folders[$key].status = "red" }
        foreach ($key in $currentState.files.Keys) { $currentState.files[$key].status = "red" }

        # Track which folders changed
        $changedFolders = @{}
        foreach ($key in $currentState.folders.Keys) {
            if (-not $savedState.folders.ContainsKey($key)) {
                $currentState.folders[$key].status = "green"
                $changedFolders[$key] = $true
            } elseif ((ConvertTo-Json $currentState.folders[$key] -Compress) -ne (ConvertTo-Json $savedState.folders[$key] -Compress)) {
                $currentState.folders[$key].status = "green"
                $changedFolders[$key] = $true
            }
        }
        # Mark only direct parent branch green if any child changed
        foreach ($folderKey in $changedFolders.Keys) {
            $parentBranch = Split-Path $folderKey -Parent
            if ($parentBranch -and $currentState.folders.ContainsKey($parentBranch) -and $currentState.folders[$parentBranch].status -ne "green") {
                $currentState.folders[$parentBranch].status = "green"
            }
        }
        # Files: only mark green if changed/added
        foreach ($key in $currentState.files.Keys) {
            if (-not $savedState.files.ContainsKey($key)) {
                $currentState.files[$key].status = "green"
            } elseif ((ConvertTo-Json $currentState.files[$key] -Compress) -ne (ConvertTo-Json $savedState.files[$key] -Compress)) {
                $currentState.files[$key].status = "green"
            }
        }
        Build-Tree $currentState
        [System.Windows.Forms.MessageBox]::Show("Refresh complete. Only changed branches and their direct parent are green.", "Refresh Complete")
    } catch {
        Write-Host "Error during refresh: $_"
        [System.Windows.Forms.MessageBox]::Show("Error refreshing state: $_")
    }
})

function Compare-Hashtables {
    param($ht1, $ht2)
    return ((ConvertTo-Json $ht1 -Compress) -eq (ConvertTo-Json $ht2 -Compress))
}

function Merge-Hashtables {
    param(
        [hashtable]$ht1,
        [hashtable]$ht2
    )
    foreach ($key in $ht2.Keys) {
        $ht1[$key] = $ht2[$key]
    }
    return $ht1
}

function Refresh-State {
    param($state, $folderPath)
    
    Write-Host "[Refresh-State] Refreshing state for folder: $folderPath"
    $state = DeepConvertState $state
    Write-Host "[Refresh-State] State type after deep conversion: $($state.GetType().FullName)"
    
    # Get current state of all items
    $currentItems = Get-ChildItem -Path $folderPath -Recurse
    
    # Check root folder
    $rootInfo = Get-Item $folderPath
    $rootModified = $rootInfo.LastWriteTimeUtc.ToString("o")
    
    if ($state.folders -and $state.folders.ContainsKey('.')) {
        Write-Host "Root folder initial time: $($state.folders['.'].initial_modified)"
        Write-Host "Root folder current time: $rootModified"
        
        try {
            $initialTime = [DateTime]::Parse($state.folders['.'].initial_modified)
            $currentTime = [DateTime]::Parse($rootModified)
            $hasChanged = $initialTime -ne $currentTime
            
            Write-Host "Root folder has changed: $hasChanged"
            
            if ($hasChanged) {
                $state.folders['.'].current_modified = $rootModified
                $state.folders['.'].status = "green"
                $state.folders['.'].hasChanges = $true
            } else {
                $state.folders['.'].status = "red"
                $state.folders['.'].hasChanges = $false
            }
        } catch {
            Write-Host "Error comparing root folder times: $_"
        }
    }

    # Check folders
    foreach ($folder in $currentItems | Where-Object { $_.PSIsContainer }) {
        $relativePath = $folder.FullName.Substring($folderPath.Length).TrimStart('\')
        $lastMod = $folder.LastWriteTimeUtc.ToString("o")
        
        Write-Host "Checking folder: $relativePath"
        
        if ($state.folders -and $state.folders.ContainsKey($relativePath)) {
            # Existing folder - check for changes
            Write-Host "  Initial time: $($state.folders[$relativePath].initial_modified)"
            Write-Host "  Current time: $lastMod"
            
            try {
                $initialTime = [DateTime]::Parse($state.folders[$relativePath].initial_modified)
                $currentTime = [DateTime]::Parse($lastMod)
                $hasChanged = $initialTime -ne $currentTime
                
                Write-Host "  Has changed: $hasChanged"
                
                $state.folders[$relativePath].current_modified = $lastMod
                if ($hasChanged) {
                    $state.folders[$relativePath].status = "green"
                    $state.folders[$relativePath].hasChanges = $true
                } else {
                    $state.folders[$relativePath].status = "red"
                    $state.folders[$relativePath].hasChanges = $false
                }
            } catch {
                Write-Host "  Error comparing times: $_"
            }
        } else {
            Write-Host "  New folder found"
            # New folder found
            $state.folders[$relativePath] = @{
                path = $relativePath
                creation_date = $folder.CreationTimeUtc.ToString("o")
                initial_modified = $lastMod
                current_modified = $lastMod
                status = "red"
                hasChanges = $false
            }
        }
    }

    # Check files
    foreach ($file in $currentItems | Where-Object { -not $_.PSIsContainer }) {
        $relativePath = $file.FullName.Substring($folderPath.Length).TrimStart('\')
        $lastMod = $file.LastWriteTimeUtc.ToString("o")
        $currentSize = $file.Length
        
        Write-Host "Checking file: $relativePath"
        
        if ($state.files.ContainsKey($relativePath)) {
            # Existing file - check for changes
            Write-Host "  Initial time: $($state.files[$relativePath].initial_modified)"
            Write-Host "  Initial size: $($state.files[$relativePath].initial_size)"
            Write-Host "  Current time: $lastMod"
            Write-Host "  Current size: $currentSize"
            
            try {
                $initialTime = [DateTime]::Parse($state.files[$relativePath].initial_modified)
                $currentTime = [DateTime]::Parse($lastMod)
                $hasChanged = ($initialTime -ne $currentTime) -or ($state.files[$relativePath].initial_size -ne $currentSize)
                
                Write-Host "  Has changed: $hasChanged"
                
                $state.files[$relativePath].current_modified = $lastMod
                $state.files[$relativePath].current_size = $currentSize
                if ($hasChanged) {
                    $state.files[$relativePath].status = "green"
                    $state.files[$relativePath].hasChanges = $true
                } else {
                    $state.files[$relativePath].status = "red"
                    $state.files[$relativePath].hasChanges = $false
                }
            } catch {
                Write-Host "  Error comparing times: $_"
            }
        } else {
            Write-Host "  New file found"
            # New file found
            $state.files[$relativePath] = @{
                path = $relativePath
                name = $file.Name
                creation_date = $file.CreationTimeUtc.ToString("o")
                initial_modified = $lastMod
                current_modified = $lastMod
                initial_size = $currentSize
                current_size = $currentSize
                status = "red"
                hasChanges = $false
            }
        }
    }

    # Detect missing folders (previously captured but now gone)
    foreach ($savedKey in $state.folders.Keys) {
        if (-not ($currentItems | Where-Object { $_.PSIsContainer -and $_.FullName.Substring($folderPath.Length).TrimStart('\') -eq $savedKey })) {
            Write-Host "MISSING FOLDER: $savedKey"
            if ($state.folders[$savedKey] -is [hashtable]) {
                $state.folders[$savedKey].status = "missing"
            } elseif ($state.folders[$savedKey] -is [PSCustomObject] -and $state.folders[$savedKey].GetType().Name -ne 'Hashtable') {
                $state.folders[$savedKey] = @{} + $state.folders[$savedKey]
                $state.folders[$savedKey].status = "missing"
            }
        }
    }

    # Detect missing files (previously captured but now gone)
    foreach ($savedKey in $state.files.Keys) {
        if (-not ($currentItems | Where-Object { -not $_.PSIsContainer -and $_.FullName.Substring($folderPath.Length).TrimStart('\') -eq $savedKey })) {
            Write-Host "MISSING FILE: $savedKey"
            if ($state.files[$savedKey] -is [hashtable]) {
                $state.files[$savedKey].status = "missing"
            } elseif ($state.files[$savedKey] -is [PSCustomObject] -and $state.files[$savedKey].GetType().Name -ne 'Hashtable') {
                $state.files[$savedKey] = @{} + $state.files[$savedKey]
                $state.files[$savedKey].status = "missing"
            }
        }
    }

    # Compare the current state with original state to detect real changes
    $hasRealChanges = $false
    
    # Compare folders
    foreach ($folderPath in $state.folders.Keys) {
        $currentFolder = $state.folders[$folderPath]
        $originalFolder = $originalState.folders.$folderPath
        
        if ($originalFolder) {
            # Parse dates to ensure consistent comparison
            try {
                $originalTime = [DateTime]::Parse($originalFolder.current_modified)
                $currentTime = [DateTime]::Parse($currentFolder.current_modified)
                
                if ($originalTime -ne $currentTime) {
                    $hasRealChanges = $true
                    $currentFolder.hasChanges = $true
                    $currentFolder.status = "green"
                    Write-Host "CHANGE DETECTED - Folder: $folderPath"
                    Write-Host "  Original Modified: $($originalTime.ToString('o'))"
                    Write-Host "  Current Modified:  $($currentTime.ToString('o'))"
                } else {
                    $currentFolder.hasChanges = $false
                    $currentFolder.status = "red"
                }
            } catch {
                Write-Host "Error comparing dates for folder $folderPath : $_"
            }
        }
    }
    
    # Compare files
    foreach ($filePath in $state.files.Keys) {
        $currentFile = $state.files[$filePath]
        $originalFile = $originalState.files.$filePath
        
        if ($originalFile) {
            try {
                # Parse dates to DateTime objects for consistent comparison
                $originalTime = [DateTime]::Parse($originalFile.current_modified)
                $currentTime = [DateTime]::Parse($currentFile.current_modified)
                
                $timeChanged = $originalTime -ne $currentTime
                $sizeChanged = $currentFile.current_size -ne $originalFile.current_size
                
                if ($timeChanged -or $sizeChanged) {
                    $hasRealChanges = $true
                    $currentFile.hasChanges = $true
                    $currentFile.status = "green"
                    Write-Host "CHANGE DETECTED - File: $filePath"
                    if ($timeChanged) {
                        # Use consistent ISO 8601 format for both dates
                        Write-Host "  Original Modified: $($originalTime.ToString('o'))"
                        Write-Host "  Current Modified:  $($currentTime.ToString('o'))"
                    }
                    if ($sizeChanged) {
                        Write-Host "  Original Size: $($originalFile.current_size) bytes"
                        Write-Host "  Current Size:  $($currentFile.current_size) bytes"
                    }
                } else {
                    $currentFile.hasChanges = $false
                    $currentFile.status = "red"
                }
            } catch {
                Write-Host "Error comparing dates for file $filePath : $_"
            }
        }
    }
    
    # Update lastCheck without triggering a save just for timestamp
    $state.lastCheck = (Get-Date).ToString("o")
    
    Write-Host "Has real changes to save: $hasRealChanges"
    
    if ($hasRealChanges) {
        $jsonFileName = [System.IO.Path]::GetFileName($JsonTextBox.Text)
        $jsonPath = Join-Path $PWD $jsonFileName
        Write-Host "Saving state to: $jsonPath"
        $state | ConvertTo-Json -Depth 5 | Set-Content $jsonPath
    } else {
        Write-Host "No real changes detected, skipping save"
    }
    
    return $state
}

function Load-Json {
    param([string]$jsonPath)
    Write-Host "[Load-Json] Loading JSON from: $jsonPath"
    if (Test-Path $jsonPath) {
        $jsonContent = Get-Content $jsonPath -Raw
        $state = ConvertFrom-Json $jsonContent
        $state = DeepConvertState $state
        $Global:LoadedJsonPath = $jsonPath
        $Global:LoadedState = $state
        Write-Host "[Load-Json] Loaded and converted JSON. Path saved for refresh."
    } else {
        Write-Host "[Load-Json] File not found: $jsonPath"
    }
}

function Ensure-StateObject {
    param($state)
    # If state is a hashtable with only system properties, reconstruct as PSCustomObject
    if ($state -is [hashtable] -and -not ($state.PSObject.Properties.Name -contains 'folders')) {
        $newState = [PSCustomObject]@{
            folders = $null
            files = $null
            rootPath = $null
            lastCheck = $null
            isInitialized = $null
        }
        foreach ($key in $state.Keys) {
            if ($newState.PSObject.Properties.Name -contains $key) {
                $newState.$key = $state[$key]
            }
        }
        return $newState
    }
    return $state
}

# Initial State Load
if (-not (Test-Path $PWD)) {
    New-Item -Path $PWD -ItemType Directory -Force | Out-Null
}

# Only try to load initial state if we have a valid JSON path
$jsonFileName = [System.IO.Path]::GetFileName($JsonTextBox.Text)
if (-not [string]::IsNullOrWhiteSpace($jsonFileName)) {
    $state = Load-State $jsonFileName
    if ($state -and $state.Count -gt 0) {
        $state = Update-State $state
        Build-Tree $state
    }
}

$Form.ShowDialog()
