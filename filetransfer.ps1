Add-Type -AssemblyName System.Windows.Forms

$servers = @( # Name <=> IP-adress
    [PSCustomObject]@{ Os = "Windows"; Name = "1.2.3.4"; Path = "C:\destination\" },
    [PSCustomObject]@{ Os = "Linux"; Name = "5.6.7.8"; Path = "/destination/" }
)


# Get path oft this script and the logfile
$scriptPath = $MyInvocation.MyCommand.Path
$scriptDirectory = [System.IO.Path]::GetDirectoryName($scriptPath)
$logFilePath = Join-Path -Path $scriptDirectory -ChildPath "logfile.txt"

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
    # $OpenFileDialog.Filter = "Textfiles (*.txt)|*.txt"
    $OpenFileDialog.ShowDialog() | Out-Null
    return $OpenFileDialog.FileName
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
            Write-Log "License file copied to $serverName $destinationPath"
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
        Write-Log "License file copied to $serverName $destinationPath"
    }    
}

function Copy-Loop {
    Write-Log "Process started"
    $licensePath = $pathTextBox.Text

    if ([string]::IsNullOrEmpty($licensePath)) {
        Write-Log "No file was selected. Process ended`n"
        exit
    }
    Show_Progress_Screen
    $credential = $Host.ui.PromptForCredential("Authorization required", "Please sign in with Administrator rights.", "", "NetBiosUserName")

    # $credentialValid = $false
    # while (-not $credentialValid) {
    #     $credential = $Host.ui.PromptForCredential("Authorization required", "Please sign in with Administrator rights.", "", "NetBiosUserName")
    #     if ($null -eq $credential) {
    #         Write-Log "Credential prompt canceled by user. Process ended`n"
    #         Show_Start_Screen
    #         return
    #     }
    #     $credentialValid = Test-WindowsCredentials -serverName $servers[0].Name -credential $credential
    #     if (-not $credentialValid) {
    #         [System.Windows.Forms.MessageBox]::Show("Invalid credentials. Please try again.")
    #     }
    # }

    $progressBar.Visible = $true
    $index = 0
    $allServers = $servers.Count
    $global:errorCount = 0 # 'global' is needed to get the error count from the functions
    
    foreach ($server in $servers) {

        $label.Text = "Copying files to " + $server.Name + $server.Path + "`n($index / $allServers)"
        Start-Sleep -Milliseconds 200 # Delay to prevent the GUI to get stuck

        if ($server.Os -eq "Windows") {
            Copy-LicenseToWindows  -licensePath $licensePath -serverName $server.Name -destinationPath $server.Path -credential $credential
        }
        elseif ($server.Os -eq "Linux") {
            if ($server.Name -eq "server22") {
                $password = "abc"
            }
            else {
                $password = "xyz"
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
}

function Show_Progress_Screen {
    $label.Visible = $true
    $label.Text = "Waiting for credentials..."
    $buttonSelect.Visible = $false
    $pathTextBox.Visible = $false
    $buttonStart.Visible = $false
    $buttonCancel.Visible = $false
    $progressLabel.Visible = $true
}

function Show_Start_Screen {
    $label.Text = "Welcome to the automatic Installer for the COBOL-IT-licenses.`nPlease select the license you want to distribute to all servers and then press start."
    $buttonStart.Visible = $true
    $buttonCancel.Visible = $true
    $buttonShowLog.Visible = $false
    $buttonSelect.Visible = $true
    $pathTextBox.Visible = $true
    $progressBar.Visible = $false
    $progressLabel.Visible = $false
    $pathTextBox.Text = ""
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "COBOL-IT License Installer"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
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
$pathTextBox.Location = '20,60'
$pathTextBox.ReadOnly = $true
$form.Controls.Add($pathTextBox)

$buttonSelect = New-Object System.Windows.Forms.Button
$buttonSelect.Text = "Select"
$buttonSelect.Location = '330,60'
$buttonSelect.Add_Click({ 
        $pathTextBox.Text = Get-FileDialog
    })
$form.Controls.Add($buttonSelect)

$buttonStart = New-Object System.Windows.Forms.Button
$buttonStart.Text = "Start"
$buttonStart.Top = 90
$buttonStart.Left = 50
$buttonStart.Enabled = $false
$buttonStart.Add_Click({ 
        $label.Visible = $false
        Copy-Loop
    })
$form.Controls.Add($buttonStart)

$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Text = "Abort"
$buttonCancel.Top = 90
$buttonCancel.Left = 150
$buttonCancel.Add_Click({ 
        $terminateScript = $true 
        $form.Close()
        return $terminateScript 
    })
$form.Controls.Add($buttonCancel)

$buttonShowLog = New-Object System.Windows.Forms.Button
$buttonShowLog.Text = "Show logfile"
$buttonShowLog.Top = 90
$buttonShowLog.Left = 50
$buttonShowLog.Add_Click({ Start-Process C:\Users\me\Repositories\Powershell\cobolinstaller\logfile.txt })
$buttonShowLog.Visible = $false
$form.Controls.Add($buttonShowLog)

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

#Sets the starting position of the form at run time.
$CenterScreen = [System.Windows.Forms.FormStartPosition]::CenterScreen;
$form.StartPosition = $CenterScreen;

#Checks if there is a path
$pathTextBox.Add_TextChanged({
        if ($pathTextBox.TextLength -eq 0) {
            $buttonStart.Enabled = $false
        }
        else {
            $buttonStart.Enabled = $true
        }
    })

$form.ShowDialog()
