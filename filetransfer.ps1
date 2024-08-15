Add-Type -AssemblyName System.Windows.Forms

# Get path oft this script and the logfile
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = [System.IO.Path]::GetDirectoryName($scriptPath)
$logFilePath = Join-Path -Path $scriptDirectory -ChildPath "logfile.txt"
$serverListFilePath = Join-Path -Path $scriptDirectory -ChildPath "serverlist.csv"

function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFilePath -Value $logMessage
}

function Get-FileDialog {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    #$OpenFileDialog.Filter = "XML files (*.xml)|*.xml"
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
}

# if no serverlist.csv, a new on with default value is created 
function Create-DefaultServerList {
    $defaultContent = @"
"OS","Name","Path"
"@

    # Write the default values to the new CSV file
    $defaultContent | Set-Content -Path $serverListFilePath
    Write-Log "Server list file created with default values."
}

function Load-ServerList {
    if (-not (Test-Path $serverListFilePath)) {
        Create-DefaultServerList
    }

    $dataTable = New-Object System.Data.DataTable
    $dataTable.Columns.Add("OS")
    $dataTable.Columns.Add("Name")
    $dataTable.Columns.Add("Path")

    $serverList = Import-Csv -Path $serverListFilePath
    foreach ($server in $serverList) {
        $row = $dataTable.NewRow()
        $row.OS = $server.OS
        $row.Name = $server.Name
        $row.Path = $server.Path
        $dataTable.Rows.Add($row)
    }

    $dataGridView.DataSource = $dataTable
}

function Save-ServerList {
    $serverListFile = $serverListFilePath
    $dataTable = $dataGridView.DataSource
    $serverList = @()

    foreach ($row in $dataTable.Rows) {
        $serverList += [PSCustomObject]@{
            OS = $row.OS
            Name = $row.Name
            Path         = $row.Path
        }
    }

    $serverList | Export-Csv -Path $serverListFile -NoTypeInformation
}

function Test-WindowsCredentials {
    param (
        [string]$serverName,
        [pscredential]$credential
    )

    try {
        $session = New-PSSession -ComputerName $serverName -Credential $credential
        Remove-PSSession -Session $session
        return $true
    }
    catch {
        return $false
    }
}

function Copy-LicenseToWindows {
    param (
        [string]$licensePath,
        [string]$serverName,
        [string]$destinationPath,
        [pscredential]$credential
    )

    try {
        Write-Log "Connecting to $serverName $destinationPath"
        $session = New-PSSession -ComputerName $serverName -Credential $credential
        Copy-Item -Path $licensePath -Destination $destinationPath -ToSession $session -Force
        Invoke-Command -Session $session -ScriptBlock {
            param (
                $serverName,
                $destinationPath
            )
            Write-Log "License file successfully copied to $serverName $destinationPath"
        } -ArgumentList $destinationPath
        Remove-PSSession -Session $session
    }
    catch {
        Write-Log "ERROR. Failed to copy license file to $serverName $destinationPath :  $_"
        $global:errorCount++
    }
}

function Copy-LicenseToLinux {
    param (
        [string]$licensePath,
        [string]$serverName,
        [string]$destinationPath,
        [string]$username,
        [string]$password
    )
    
    $destination = "$username@$($server.Name):$($server.Path)"

    Write-Log "Connecting to $serverName $destinationPath"

    pscp -pw $password $licensePath $destination

    if ($LASTEXITCODE -eq 1) {
        $global:errorCount++
        Write-Log "ERROR. Failed to copy license file to $serverName $destinationPath :  $_"
    }
    else {
        Write-Log "License file successfully copied to $serverName $destinationPath"
    }    
}

function Copy-Loop {
    if (-not (Test-Path $serverListFilePath)) {
        Create-DefaultServerList
    }

    Write-Log "Process started"
    $licensePath = $pathTextBox.Text

    if ([string]::IsNullOrEmpty($licensePath)) {
        Write-Log "No file was selected. Process ended`n"
        exit
    }
    Show_Progress_Screen

    $credentialValid = $false
    while (-not $credentialValid) {
        $credential = $Host.ui.PromptForCredential("Authorization required", "Please sign in with Administrator rights.", "", "NetBiosUserName")
        if ($null -eq $credential) {
            Write-Log "Credential prompt canceled by user. Process ended`n"
            Show_Start_Screen
            return
        }
        $credentialValid = Test-WindowsCredentials -serverName $servers[0].Name -credential $credential
        if (-not $credentialValid) {
            [System.Windows.Forms.MessageBox]::Show("Invalid credentials. Please try again.")
        }
    }

    $progressBar.Visible = $true
    $index = 0
    $serverList = Import-Csv -Path $serverListFilePath
    $allServers = $serverList.Count
    $global:errorCount = 0 # 'global' is needed to get the error count from the functions
    
    foreach ($server in $serverList) {

        $label.Text = "Copying files to " + $server.Name + $server.Path + "`n($index / $allServers)"
        Start-Sleep -Milliseconds 200 # Delay to prevent the GUI to get stuck

        if ($server.OS -eq "Windows") {
            Copy-LicenseToWindows  -licensePath $licensePath -serverName $server.Name -destinationPath $server.Path -credential $credential
        }
        elseif ($server.OS -eq "Linux") {
                $password = "root"
            }

            Copy-LicenseToLinux -licensePath $licensePath -serverName $server.Name -destinationPath $server.Path -username "quser" -password $password
        }
        $index++
        $progressPercent = ($index / $allServers) * 100
        $progressBar.Value = $progressPercent
        $roundedPercentage = [math]::Round($progressPercent)
        $progressLabel.Text = "$roundedPercentage% Complete"
    }
    Write-Log "Finished copy process. Process ended`n"
    Show_End_Screen
}

#Functions to display different screens
function Show_End_Screen {
    $label.Visible = $true
    $label.Text = "Finished copy process with $global:errorCount error(s). Check the logfile for detailed information."

    $buttonStart.Visible = $false 
    $buttonCancel.Visible = $true
    $buttonCancel.Text = "Close"
    $buttonShowLog.Visible = $true
    $buttonSelect.Visible = $false
    $pathTextBox.Visible = $false
    $progressBar.Visible = $false
    $progressLabel.Visible = $false
    $buttonCancel.Location = "250,90"
}

function Show_Progress_Screen {
    $label.Visible = $true
    $label.Text = "Waiting for credentials..."
    $buttonSelect.Visible = $false
    $pathTextBox.Visible = $false
    $buttonStart.Visible = $false
    $buttonCancel.Visible = $false
    $buttonServers.Visible = $false
    $progressLabel.Visible = $true
}

function Show_Start_Screen {
    $label.Text = "Welcome to the automatic Installer for the COBOL-IT-licenses.`nPlease select the license you want to distribute to all servers and then press start."
    $buttonStart.Visible = $true
    $buttonCancel.Visible = $true
    $buttonShowLog.Visible = $false
    $buttonSelect.Visible = $true
    $buttonServers.Visible = $true
    $pathTextBox.Visible = $true
    $progressBar.Visible = $false
    $progressLabel.Visible = $false
    $pathTextBox.Text = ""
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "COBOL-IT License Installer"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false;
$form.Width = 500
$form.Height = 200

$label = New-Object System.Windows.Forms.Label
$label.Text = "Welcome to the automatic Installer for the COBOL-IT-licenses.`nPlease select the license you want to distribute to all servers and then press start."
$label.AutoSize = $true
$label.Top = 20
$label.Left = 20
$form.Controls.Add($label)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Size = '300,23'
$pathTextBox.Location = '20,64'
$pathTextBox.ReadOnly = $true
$form.Controls.Add($pathTextBox)

$buttonSelect = New-Object System.Windows.Forms.Button
$buttonSelect.Text = "Select"
$buttonSelect.Location = '330,60'
$buttonSelect.Size = "80,30"
$buttonSelect.Add_Click({ 
    $pathTextBox.Text = Get-FileDialog
})
$form.Controls.Add($buttonSelect)

$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Start"
$buttonStart.Location = "30,110"
$buttonStart.Size = "80,30"
$buttonStart.Enabled = $false
$buttonStart.Add_Click({ 
    $label.Visible = $false
    Copy-Loop
})
$form.Controls.Add($buttonStart)

$buttonServers = New-Object System.Windows.Forms.Button
$buttonServers.Text = "Edit Servers"
$buttonServers.Location = "130,110"
$buttonServers.Size = "80,30"
$buttonServers.Add_Click({
    Load-ServerList
    $settingsForm.ShowDialog()
})
$form.Controls.Add($buttonServers)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Text = "Abort"
$buttonCancel.Location = "330,110"
$buttonCancel.Size = "80,30"
$buttonCancel.Add_Click({ 
    $terminateScript = $true 
    $form.Close()
    return $terminateScript 
})
$form.Controls.Add($buttonCancel)

$buttonShowLog = New-Object System.Windows.Forms.Button
$buttonShowLog.Text = "Show logfile"
$buttonShowLog.Location = "50,90"
$buttonShowLog.Size = "80,30"
$buttonShowLog.Add_Click({ Start-Process $logFilePath })
$buttonShowLog.Visible = $false
$form.Controls.Add($buttonShowLog)

#Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(35, 60)
$progressBar.Size = New-Object System.Drawing.Size(280, 20)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Location = New-Object System.Drawing.Point(325, 60)
$progressLabel.AutoSize = $true
$progressLabel.Visible = $false
$form.Controls.Add($progressLabel)

# serverliste bearbeiten
$settingsForm = New-Object System.Windows.Forms.Form
$settingsForm.Text = "Server Settings"
$settingsForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$settingsForm.MaximizeBox = $false;
$settingsForm.Width = 580
$settingsForm.Height = 580

$panelServerList = New-Object System.Windows.Forms.Panel
$panelServerList.Size = New-Object System.Drawing.Size(560, 460)
$panelServerList.Location = New-Object System.Drawing.Point(10, 20)
$settingsForm.Controls.Add($panelServerList)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Size = New-Object System.Drawing.Size(530, 450)
$dataGridView.Location = New-Object System.Drawing.Point(10, 10)
$dataGridView.AutoGenerateColumns = $true
$dataGridView.AutoSizeColumnsMode = 'AllCells'
$dataGridView.AllowUserToAddRows = $true
$dataGridView.AllowUserToDeleteRows = $true
$panelServerList.Controls.Add($dataGridView)

$buttonRemoveServer = New-Object System.Windows.Forms.Button
$buttonRemoveServer.Text = "Remove Selected"
$buttonRemoveServer.Size = New-Object System.Drawing.Size(120, 30)
$buttonRemoveServer.Location = New-Object System.Drawing.Point(250, 495)
$buttonRemoveServer.Add_Click({
    if ($dataGridView.SelectedRows.Count -gt 0) {
        $dataGridView.Rows.Remove($dataGridView.SelectedRows[0])
    }
})

$settingsForm.Controls.Add($buttonRemoveServer)
$buttonSaveServerList = New-Object System.Windows.Forms.Button
$buttonSaveServerList.Text = "Save"
$buttonSaveServerList.Size = New-Object System.Drawing.Size(80, 30)
$buttonSaveServerList.Location = New-Object System.Drawing.Point(380, 495)
$buttonSaveServerList.Add_Click({
    Save-ServerList
})
$settingsForm.Controls.Add($buttonSaveServerList)

$buttonExitServerlist = New-Object System.Windows.Forms.Button
$buttonExitServerlist.Text = "Exit"
$buttonExitServerlist.Size = New-Object System.Drawing.Size(80, 30)
$buttonExitServerlist.Location = New-Object System.Drawing.Point(470, 495)
$buttonExitServerlist.Add_Click({
    $settingsForm.Close()
})
$settingsForm.Controls.Add($buttonExitServerlist)


# Sets the starting position of the form at run time.
$CenterScreen = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.StartPosition = $CenterScreen
$settingsForm.StartPosition = $CenterScreen

# Checks if there is a path
$pathTextBox.Add_TextChanged({
    if ($pathTextBox.TextLength -eq 0) {
        $buttonStart.Enabled = $false
    } else {
        $buttonStart.Enabled = $true
    }
})

$form.ShowDialog()
