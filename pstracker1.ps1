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
$IssueLabel.Location = New-Object System.Drawing.Point(340,10)
$IssueLabel.Size = New-Object System.Drawing.Size(90,20)
$Form.Controls.Add($IssueLabel)

$IssueTextBox = New-Object System.Windows.Forms.TextBox
$IssueTextBox.Location = New-Object System.Drawing.Point(440,10)
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
$JsonLabel.Location = New-Object System.Drawing.Point(550,10)
$JsonLabel.Size = New-Object System.Drawing.Size(80,20)
$Form.Controls.Add($JsonLabel)

$JsonTextBox = New-Object System.Windows.Forms.TextBox
$JsonTextBox.Location = New-Object System.Drawing.Point(640,10)
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

# Global variables for state management
$Global:BaselineState = $null
$Global:CurrentState = $null

# Helper Functions
function ConvertTo-ConsistentHashtable($obj) {
    if ($obj -is [System.Collections.Hashtable]) {
        $result = @{}
        foreach ($key in $obj.Keys) {
            $result[$key] = ConvertTo-ConsistentHashtable $obj[$key]
        }
        return $result
    } 
    elseif ($obj -is [PSCustomObject]) {
        $result = @{}
        foreach ($prop in $obj.PSObject.Properties) {
            $result[$prop.Name] = ConvertTo-ConsistentHashtable $prop.Value
        }
        return $result
    } 
    elseif ($obj -is [System.Array]) {
        return @($obj | ForEach-Object { ConvertTo-ConsistentHashtable $_ })
    } 
    else {
        return $obj
    }
}

function Create-StateObject {
    param(
        [string]$rootPath,
        [hashtable]$folders = @{},
        [hashtable]$files = @{}
    )
    
    return @{
        rootPath = $rootPath
        folders = $folders
        files = $files
        lastCheck = (Get-Date).ToString("o")
        isInitialized = $true
    }
}

function Get-CurrentFileSystemState {
    param([string]$folderPath)
    
    Write-Host "Getting current filesystem state for: $folderPath"
    
    if (-not (Test-Path $folderPath)) {
        Write-Host "ERROR: Folder path does not exist"
        return Create-StateObject -rootPath $folderPath
    }
    
    $folders = @{}
    $files = @{}
    
    # Get all items recursively
    $allItems = Get-ChildItem -Path $folderPath -Recurse -ErrorAction SilentlyContinue
    
    # Process folders
    foreach ($folder in ($allItems | Where-Object { $_.PSIsContainer })) {
        $relativePath = $folder.FullName.Substring($folderPath.Length).TrimStart('\')
        if ($relativePath -ne '') {
            $folders[$relativePath] = @{
                path = $relativePath
                creation_date = $folder.CreationTimeUtc.ToString("o")
                last_modified = $folder.LastWriteTimeUtc.ToString("o")
                status = "red"
                hasChanges = $false
            }
        }
    }
    
    # Process files
    foreach ($file in ($allItems | Where-Object { -not $_.PSIsContainer })) {
        $relativePath = $file.FullName.Substring($folderPath.Length).TrimStart('\')
        $files[$relativePath] = @{
            path = $relativePath
            name = $file.Name
            creation_date = $file.CreationTimeUtc.ToString("o")
            last_modified = $file.LastWriteTimeUtc.ToString("o")
            size = $file.Length
            status = "red"
            hasChanges = $false
        }
    }
    
    # Add root folder if no subfolders exist
    if ($folders.Count -eq 0) {
        $rootInfo = Get-Item $folderPath
        $folders['.'] = @{
            path = '.'
            creation_date = $rootInfo.CreationTimeUtc.ToString("o")
            last_modified = $rootInfo.LastWriteTimeUtc.ToString("o")
            status = "red"
            hasChanges = $false
        }
    }
    
    Write-Host "Found $($folders.Count) folders and $($files.Count) files"
    return Create-StateObject -rootPath $folderPath -folders $folders -files $files
}

function Compare-States {
    param(
        [hashtable]$baselineState,
        [hashtable]$currentState
    )
    
    Write-Host "=== Comparing States ==="
    $hasChanges = $false
    
    # Compare folders
    foreach ($folderKey in $currentState.folders.Keys) {
        $currentFolder = $currentState.folders[$folderKey]
        
        if ($baselineState.folders.ContainsKey($folderKey)) {
            $baselineFolder = $baselineState.folders[$folderKey]
            
            # Compare last modified times
            $baselineTime = [DateTime]::Parse($baselineFolder.last_modified)
            $currentTime = [DateTime]::Parse($currentFolder.last_modified)
            
            if ($baselineTime -ne $currentTime) {
                Write-Host "CHANGE DETECTED - Folder: $folderKey"
                Write-Host "  Baseline: $($baselineTime.ToString('o'))"
                Write-Host "  Current:  $($currentTime.ToString('o'))"
                
                $currentFolder.status = "green"
                $currentFolder.hasChanges = $true
                $hasChanges = $true
                
                # Mark parent folders as green too
                $parentPath = Split-Path $folderKey -Parent
                while ($parentPath -and $currentState.folders.ContainsKey($parentPath)) {
                    if ($currentState.folders[$parentPath].status -ne "green") {
                        $currentState.folders[$parentPath].status = "green"
                        Write-Host "  Marking parent green: $parentPath"
                    }
                    $parentPath = Split-Path $parentPath -Parent
                }
            } else {
                $currentFolder.status = "red"
                $currentFolder.hasChanges = $false
            }
        } else {
            # New folder
            Write-Host "NEW FOLDER: $folderKey"
            $currentFolder.status = "green"
            $currentFolder.hasChanges = $true
            $hasChanges = $true
        }
    }
    
    # Compare files
    foreach ($fileKey in $currentState.files.Keys) {
        $currentFile = $currentState.files[$fileKey]
        
        if ($baselineState.files.ContainsKey($fileKey)) {
            $baselineFile = $baselineState.files[$fileKey]
            
            # Compare last modified times and sizes
            $baselineTime = [DateTime]::Parse($baselineFile.last_modified)
            $currentTime = [DateTime]::Parse($currentFile.last_modified)
            
            $timeChanged = $baselineTime -ne $currentTime
            $sizeChanged = $baselineFile.size -ne $currentFile.size
            
            if ($timeChanged -or $sizeChanged) {
                Write-Host "CHANGE DETECTED - File: $fileKey"
                if ($timeChanged) {
                    Write-Host "  Time - Baseline: $($baselineTime.ToString('o'))"
                    Write-Host "  Time - Current:  $($currentTime.ToString('o'))"
                }
                if ($sizeChanged) {
                    Write-Host "  Size - Baseline: $($baselineFile.size) bytes"
                    Write-Host "  Size - Current:  $($currentFile.size) bytes"
                }
                
                $currentFile.status = "green"
                $currentFile.hasChanges = $true
                $hasChanges = $true
                
                # Mark parent folder as green
                $parentPath = Split-Path $fileKey -Parent
                if ($parentPath -and $currentState.folders.ContainsKey($parentPath)) {
                    if ($currentState.folders[$parentPath].status -ne "green") {
                        $currentState.folders[$parentPath].status = "green"
                        Write-Host "  Marking parent green: $parentPath"
                    }
                }
            } else {
                $currentFile.status = "red"
                $currentFile.hasChanges = $false
            }
        } else {
            # New file
            Write-Host "NEW FILE: $fileKey"
            $currentFile.status = "green"
            $currentFile.hasChanges = $true
            $hasChanges = $true
            
            # Mark parent folder as green
            $parentPath = Split-Path $fileKey -Parent
            if ($parentPath -and $currentState.folders.ContainsKey($parentPath)) {
                if ($currentState.folders[$parentPath].status -ne "green") {
                    $currentState.folders[$parentPath].status = "green"
                    Write-Host "  Marking parent green: $parentPath"
                }
            }
        }
    }
    
    # Check for deleted items
    foreach ($folderKey in $baselineState.folders.Keys) {
        if (-not $currentState.folders.ContainsKey($folderKey)) {
            Write-Host "DELETED FOLDER: $folderKey"
            $currentState.folders[$folderKey] = $baselineState.folders[$folderKey].Clone()
            $currentState.folders[$folderKey].status = "missing"
            $hasChanges = $true
        }
    }
    
    foreach ($fileKey in $baselineState.files.Keys) {
        if (-not $currentState.files.ContainsKey($fileKey)) {
            Write-Host "DELETED FILE: $fileKey"
            $currentState.files[$fileKey] = $baselineState.files[$fileKey].Clone()
            $currentState.files[$fileKey].status = "missing"
            $hasChanges = $true
        }
    }
    
    Write-Host "Comparison complete. Has changes: $hasChanges"
    return $hasChanges
}

function Save-StateToJson {
    param(
        [hashtable]$state,
        [string]$jsonPath
    )
    
    try {
        # Ensure directory exists
        $directory = Split-Path $jsonPath -Parent
        if (-not (Test-Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
        
        $json = $state | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($jsonPath, $json, [System.Text.Encoding]::UTF8)
        Write-Host "State saved to: $jsonPath"
        return $true
    } catch {
        Write-Host "Error saving state: $_"
        return $false
    }
}

function Load-StateFromJson {
    param([string]$jsonPath)
    
    Write-Host "Loading state from: $jsonPath"
    
    if (-not (Test-Path $jsonPath)) {
        Write-Host "JSON file not found: $jsonPath"
        return $null
    }
    
    try {
        $jsonContent = Get-Content $jsonPath -Raw -Encoding UTF8
        $state = $jsonContent | ConvertFrom-Json
        
        # Convert to consistent hashtables
        $convertedState = ConvertTo-ConsistentHashtable $state
        
        Write-Host "State loaded successfully"
        Write-Host "Folders: $($convertedState.folders.Count)"
        Write-Host "Files: $($convertedState.files.Count)"
        
        return $convertedState
    } catch {
        Write-Host "Error loading JSON: $_"
        return $null
    }
}

function Build-TreeView {
    param([hashtable]$state)
    
    Write-Host "Building tree view..."
    $TreeView.Nodes.Clear()
    
    if (-not $state -or -not $state.folders) {
        Write-Host "No state or folders to display"
        return
    }
    
    # Create root node
    $rootNode = New-Object System.Windows.Forms.TreeNode
    $rootNode.Name = "root"
    $rootNode.Text = if ($state.rootPath) { Split-Path $state.rootPath -Leaf } else { "Root" }
    
    # Set root color based on status
    if ($state.folders.ContainsKey('.')) {
        $rootStatus = $state.folders['.'].status
        $rootNode.ForeColor = switch ($rootStatus) {
            "green" { 'Green' }
            "missing" { 'Gray' }
            default { 'Red' }
        }
        if ($rootStatus -eq "missing") {
            $rootNode.NodeFont = New-Object System.Drawing.Font($TreeView.Font, [System.Drawing.FontStyle]::Strikeout)
        }
    }
    
    $TreeView.Nodes.Add($rootNode)
    
    # Create lookup for nodes
    $nodeMap = @{ '.' = $rootNode }
    
    # Add folders
    $sortedFolders = $state.folders.Keys | Where-Object { $_ -ne '.' } | Sort-Object
    foreach ($folderKey in $sortedFolders) {
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
                
                # Set color and style based on status
                $nodeStatus = if ($i -eq $parts.Length - 1) { $folder.status } else { "red" }
                $node.ForeColor = switch ($nodeStatus) {
                    "green" { 'Green' }
                    "missing" { 'Gray' }
                    default { 'Red' }
                }
                
                if ($nodeStatus -eq "missing") {
                    $node.NodeFont = New-Object System.Drawing.Font($TreeView.Font, [System.Drawing.FontStyle]::Strikeout)
                }
                
                # Add timestamp for leaf nodes
                if ($i -eq $parts.Length - 1) {
                    try {
                        $lastMod = [DateTime]::Parse($folder.last_modified).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
                        $node.Text = "$part (Modified: $lastMod)"
                    } catch {
                        $node.Text = $part
                    }
                } else {
                    $node.Text = $part
                }
                
                $parent.Nodes.Add($node)
                $nodeMap[$currentPath] = $node
            }
            
            $parent = $nodeMap[$currentPath]
        }
    }
    
    # Add files
    $sortedFiles = $state.files.Keys | Sort-Object
    foreach ($fileKey in $sortedFiles) {
        $file = $state.files[$fileKey]
        $folderPath = Split-Path $file.path -Parent
        $fileName = Split-Path $file.path -Leaf
        
        $parentNode = if ($folderPath -and $nodeMap.ContainsKey($folderPath)) { 
            $nodeMap[$folderPath] 
        } else { 
            $rootNode 
        }
        
        $fileNode = New-Object System.Windows.Forms.TreeNode
        $fileNode.Name = $fileName
        
        # Set color and style based on status
        $fileNode.ForeColor = switch ($file.status) {
            "green" { 'Green' }
            "missing" { 'Gray' }
            default { 'Red' }
        }
        
        if ($file.status -eq "missing") {
            $fileNode.NodeFont = New-Object System.Drawing.Font($TreeView.Font, [System.Drawing.FontStyle]::Strikeout)
        }
        
        # Add file details
        try {
            $lastMod = [DateTime]::Parse($file.last_modified).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
            $size = [math]::Round($file.size / 1KB, 2)
            $fileNode.Text = "$fileName (Modified: $lastMod, Size: $size KB)"
        } catch {
            $fileNode.Text = $fileName
        }
        
        $parentNode.Nodes.Add($fileNode)
    }
    
    $rootNode.Expand()
    Write-Host "Tree view built successfully"
}

function UpdateFolderPath {
    $script:basePath = "G:\Shared drives\TriMech Solutions\"
    $company = $CompanyTextBox.Text.Trim()
    $issue = $IssueTextBox.Text.Trim()
    
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

# Event Handlers
$CompanyTextBox.Add_TextChanged({ UpdateFolderPath })
$IssueTextBox.Add_TextChanged({ UpdateFolderPath })

$PickJsonButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.InitialDirectory = $PWD
    $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $JsonTextBox.Text = [System.IO.Path]::GetFileName($dialog.FileName)
    }
})

$InitButton.Add_Click({
    $company = $CompanyTextBox.Text.Trim()
    if (-not $company) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a company name.")
        return
    }
    
    $folderPath = $FolderTextBox.Text
    $jsonPath = Join-Path $PWD "$company.json"
    $JsonTextBox.Text = "$company.json"
    
    try {
        # Check if folder exists
        if (-not (Test-Path $folderPath)) {
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
        
        # Get current state and save as baseline
        $state = Get-CurrentFileSystemState -folderPath $folderPath
        $Global:BaselineState = $state
        $Global:CurrentState = $state
        
        # Save to JSON
        if (Save-StateToJson -state $state -jsonPath $jsonPath) {
            Build-TreeView -state $state
            [System.Windows.Forms.MessageBox]::Show(
                "Initialized JSON at: $jsonPath`nTracking $($state.folders.Count) folders and $($state.files.Count) files",
                "Initialization Complete"
            )
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error saving JSON file.")
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_")
    }
})

$LoadJsonButton.Add_Click({
    $jsonFileName = [System.IO.Path]::GetFileName($JsonTextBox.Text)
    if (-not $jsonFileName) {
        [System.Windows.Forms.MessageBox]::Show("Please specify a JSON file name.")
        return
    }
    
    $jsonPath = Join-Path $PWD $jsonFileName
    $state = Load-StateFromJson -jsonPath $jsonPath
    
    if ($state) {
        $Global:BaselineState = $state
        
        # Get current filesystem state
        $folderPath = $FolderTextBox.Text
        if (Test-Path $folderPath) {
            $currentState = Get-CurrentFileSystemState -folderPath $folderPath
            
            # Compare and update colors
            Compare-States -baselineState $Global:BaselineState -currentState $currentState
            $Global:CurrentState = $currentState
            
            Build-TreeView -state $currentState
        } else {
            # Just show the baseline state
            $Global:CurrentState = $state
            Build-TreeView -state $state
        }
        
        [System.Windows.Forms.MessageBox]::Show(
            "Loaded JSON successfully`nFolders: $($state.folders.Count)`nFiles: $($state.files.Count)",
            "Load Complete"
        )
    } else {
        [System.Windows.Forms.MessageBox]::Show("Error loading JSON file.")
    }
})

$RefreshButton.Add_Click({
    if (-not $Global:BaselineState) {
        [System.Windows.Forms.MessageBox]::Show("No baseline state loaded. Please load a JSON file first.")
        return
    }
    
    $folderPath = $FolderTextBox.Text
    if (-not (Test-Path $folderPath)) {
        [System.Windows.Forms.MessageBox]::Show("Folder path does not exist: $folderPath")
        return
    }
    
    try {
        # Get current filesystem state
        $currentState = Get-CurrentFileSystemState -folderPath $folderPath
        
        # Compare with baseline
        $hasChanges = Compare-States -baselineState $Global:BaselineState -currentState $currentState
        
        # Update global state and tree view
        $Global:CurrentState = $currentState
        Build-TreeView -state $currentState
        
        $message = if ($hasChanges) {
            "Refresh complete. Changes detected and highlighted in green."
        } else {
            "Refresh complete. No changes detected since baseline."
        }
        
        [System.Windows.Forms.MessageBox]::Show($message, "Refresh Complete")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error during refresh: $_")
    }
})

# Initialize
UpdateFolderPath

$Form.ShowDialog()