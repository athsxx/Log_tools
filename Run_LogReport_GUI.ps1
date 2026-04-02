<#
.SYNOPSIS
Native Windows Forms GUI for Log Report Generator
.DESCRIPTION
This PowerShell script provides a native Windows UI that visually mimics the Python
Tkinter application but executes the headless 'cli.exe' (or 'python cli.py') in the background.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

# Check for backend
$backendExe = ".\LogReportGenerator_PowerShell_Interface.exe"
$backendPy = "python"
$backendScript = "cli.py"

$useExe = $false
if (Test-Path $backendExe) {
    $useExe = $true
}

# --- Plugins Hardcoded List ---
$plugins = @(
    @{ Key = "catia_license"; Name = "CATIA License Server" },
    @{ Key = "catia_usage_stats"; Name = "CATIA Usage Stats" },
    @{ Key = "catia_token"; Name = "CATIA Token Usage" },
    @{ Key = "ansys"; Name = "Ansys License Manager" },
    @{ Key = "ansys_peak"; Name = "Ansys Peak Usage" },
    @{ Key = "cortona"; Name = "Cortona RLM" },
    @{ Key = "cortona_admin"; Name = "Cortona Admin Server" },
    @{ Key = "creo"; Name = "Creo (PTC)" },
    @{ Key = "matlab"; Name = "MATLAB" },
    @{ Key = "nx"; Name = "NX Siemens (FlexNet/FlexLM)" }
)

# --- Forms & Layout ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Log Report Generator - PowerShell Edition"
$form.Size = New-Object System.Drawing.Size(900, 700)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Header
$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Text = "Log Report Generator"
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblHeader.ForeColor = [System.Drawing.Color]::FromArgb(183, 28, 28) # B71C1C
$lblHeader.Location = New-Object System.Drawing.Point(20, 15)
$lblHeader.AutoSize = $true
$form.Controls.Add($lblHeader)

# Tabs
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 60)
$tabControl.Size = New-Object System.Drawing.Size(840, 250)

$tabManual = New-Object System.Windows.Forms.TabPage
$tabManual.Text = "Manual Selection"
$tabManual.BackColor = [System.Drawing.Color]::White

$tabAuto = New-Object System.Windows.Forms.TabPage
$tabAuto.Text = "Auto-Scan Folder"
$tabAuto.BackColor = [System.Drawing.Color]::White

$tabControl.Controls.Add($tabManual)
$tabControl.Controls.Add($tabAuto)
$form.Controls.Add($tabControl)

# --- Manual Tab Controls ---
$lblManualSoft = New-Object System.Windows.Forms.Label
$lblManualSoft.Text = "1. Select Software Module:"
$lblManualSoft.Location = New-Object System.Drawing.Point(15, 20)
$lblManualSoft.AutoSize = $true
$tabManual.Controls.Add($lblManualSoft)

$comboPlugin = New-Object System.Windows.Forms.ComboBox
$comboPlugin.Location = New-Object System.Drawing.Point(15, 45)
$comboPlugin.Size = New-Object System.Drawing.Size(300, 30)
$comboPlugin.DropDownStyle = "DropDownList"
foreach ($p in $plugins) {
    [void]$comboPlugin.Items.Add($p.Name)
}
$comboPlugin.SelectedIndex = 0
$tabManual.Controls.Add($comboPlugin)

$btnBrowseManual = New-Object System.Windows.Forms.Button
$btnBrowseManual.Text = "2. Browse Log Files..."
$btnBrowseManual.Location = New-Object System.Drawing.Point(15, 90)
$btnBrowseManual.Size = New-Object System.Drawing.Size(300, 35)
$btnBrowseManual.BackColor = [System.Drawing.Color]::White
$tabManual.Controls.Add($btnBrowseManual)

$listManualFiles = New-Object System.Windows.Forms.ListBox
$listManualFiles.Location = New-Object System.Drawing.Point(340, 20)
$listManualFiles.Size = New-Object System.Drawing.Size(470, 180)
$listManualFiles.HorizontalScrollbar = $true
$tabManual.Controls.Add($listManualFiles)

$selectedManualFilesPaths = @()

$btnBrowseManual.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Title = "Select Log Files"
    if ($dlg.ShowDialog() -eq "OK") {
        $script:selectedManualFilesPaths = $dlg.FileNames
        $listManualFiles.Items.Clear()
        foreach ($f in $script:selectedManualFilesPaths) {
            [void]$listManualFiles.Items.Add($f)
        }
    }
})

# --- Auto-Scan Tab Controls ---
$btnBrowseAuto = New-Object System.Windows.Forms.Button
$btnBrowseAuto.Text = "1. Select Root Folder..."
$btnBrowseAuto.Location = New-Object System.Drawing.Point(15, 20)
$btnBrowseAuto.Size = New-Object System.Drawing.Size(300, 35)
$btnBrowseAuto.BackColor = [System.Drawing.Color]::White
$tabAuto.Controls.Add($btnBrowseAuto)

$lblAutoFolder = New-Object System.Windows.Forms.Label
$lblAutoFolder.Text = "No folder selected."
$lblAutoFolder.Location = New-Object System.Drawing.Point(15, 70)
$lblAutoFolder.Size = New-Object System.Drawing.Size(300, 40)
$tabAuto.Controls.Add($lblAutoFolder)

$lblAutoSoft = New-Object System.Windows.Forms.Label
$lblAutoSoft.Text = "2. Select Plugins to Scan:"
$lblAutoSoft.Location = New-Object System.Drawing.Point(340, 20)
$lblAutoSoft.AutoSize = $true
$tabAuto.Controls.Add($lblAutoSoft)

$checkPlugins = New-Object System.Windows.Forms.CheckedListBox
$checkPlugins.Location = New-Object System.Drawing.Point(340, 45)
$checkPlugins.Size = New-Object System.Drawing.Size(470, 150)
foreach ($p in $plugins) {
    [void]$checkPlugins.Items.Add($p.Name, $true) # Check all by default
}
$tabAuto.Controls.Add($checkPlugins)

$selectedAutoFolder = ""

$btnBrowseAuto.Add_Click({
    # Use FolderBrowserDialog via forms
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select root folder for auto-scan"
    if ($dlg.ShowDialog() -eq "OK") {
        $script:selectedAutoFolder = $dlg.SelectedPath
        $lblAutoFolder.Text = $script:selectedAutoFolder
    }
})

# --- Output Controls ---
$btnGenerate = New-Object System.Windows.Forms.Button
$btnGenerate.Text = "Generate Report"
$btnGenerate.Location = New-Object System.Drawing.Point(20, 320)
$btnGenerate.Size = New-Object System.Drawing.Size(840, 40)
$btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(183, 28, 28)
$btnGenerate.ForeColor = [System.Drawing.Color]::White
$btnGenerate.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$btnGenerate.FlatStyle = "Flat"
$form.Controls.Add($btnGenerate)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 370)
$progressBar.Size = New-Object System.Drawing.Size(840, 25)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

$txtConsole = New-Object System.Windows.Forms.TextBox
$txtConsole.Location = New-Object System.Drawing.Point(20, 405)
$txtConsole.Size = New-Object System.Drawing.Size(840, 240)
$txtConsole.Multiline = $true
$txtConsole.ReadOnly = $true
$txtConsole.ScrollBars = "Vertical"
$txtConsole.BackColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
$txtConsole.ForeColor = [System.Drawing.Color]::LightGray
$txtConsole.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($txtConsole)

# Helper to Log text
function Log-Text {
    param($msg)
    if ($txtConsole.InvokeRequired) {
        $txtConsole.Invoke([action]{
            $txtConsole.AppendText("$msg`r`n")
        })
    } else {
        $txtConsole.AppendText("$msg`r`n")
    }
}

function Set-Progress {
    param($val)
    if ($progressBar.InvokeRequired) {
        $progressBar.Invoke([action]{
            $progressBar.Value = $val
        })
    } else {
        $progressBar.Value = $val
    }
}

# --- Background Worker Logic ---
$worker = New-Object System.ComponentModel.BackgroundWorker
$worker.WorkerReportsProgress = $true

$worker.add_DoWork({
    param($sender, $e)
    
    $argsArray = $e.Argument
    
    $procInfo = New-Object System.Diagnostics.ProcessStartInfo
    if ($useExe) {
        $procInfo.FileName = $backendExe
        $procInfo.Arguments = $argsArray
    } else {
        $procInfo.FileName = $backendPy
        $procInfo.Arguments = "`"$backendScript`" $argsArray"
    }
    
    $procInfo.RedirectStandardOutput = $true
    $procInfo.RedirectStandardError = $true
    $procInfo.UseShellExecute = $false
    $procInfo.CreateNoWindow = $true
    $procInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $procInfo
    
    $proc.start() | Out-Null
    
    while (-not $proc.StandardOutput.EndOfStream) {
        $line = $proc.StandardOutput.ReadLine()
        if ($line -match "^B_LOG\|([^\|]*)\|(.*)") {
            $cat = $matches[1]
            $msg = $matches[2]
            $worker.ReportProgress(0, "[$cat] $msg")
        } elseif ($line -match "^B_PROG\|(\d+)") {
            $worker.ReportProgress([int]$matches[1])
        } else {
            $worker.ReportProgress(0, $line)
        }
    }
    
    # Check stderr for crashes
    $err = $proc.StandardError.ReadToEnd()
    if (![string]::IsNullOrWhiteSpace($err)) {
        $worker.ReportProgress(0, "[FATAL] $err")
    }
    
    $proc.WaitForExit()
})

$worker.add_ProgressChanged({
    param($sender, $e)
    
    if ($e.UserState -ne $null) {
        Log-Text $e.UserState
    }
    if ($e.ProgressPercentage -gt 0) {
        Set-Progress $e.ProgressPercentage
    }
})

$worker.add_RunWorkerCompleted({
    $btnGenerate.Enabled = $true
    $btnGenerate.Text = "Generate Report"
    $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(183, 28, 28)
})

# --- Execution ---
$btnGenerate.Add_Click({
    $txtConsole.Clear()
    Set-Progress 0
    
    $cmdArgs = ""
    
    if ($tabControl.SelectedTab -eq $tabManual) {
        if ($selectedManualFilesPaths.Count -eq 0) {
            Log-Text "[ERROR] Please select at least one log file."
            return
        }
        $pluginName = $comboPlugin.SelectedItem
        $pluginKey = ($plugins | Where-Object { $_.Name -eq $pluginName }).Key
        $filesStr = $selectedManualFilesPaths -join "|"
        
        $cmdArgs = "--mode manual --types `"$pluginKey`" --source `"$filesStr`""
    } else {
        if ([string]::IsNullOrWhiteSpace($selectedAutoFolder)) {
            Log-Text "[ERROR] Please select a root folder to scan."
            return
        }
        $selectedKeys = @()
        foreach ($item in $checkPlugins.CheckedItems) {
            $key = ($plugins | Where-Object { $_.Name -eq $item }).Key
            $selectedKeys += $key
        }
        if ($selectedKeys.Count -eq 0) {
            Log-Text "[ERROR] Please check at least one software type."
            return
        }
        $typesStr = $selectedKeys -join ","
        $cmdArgs = "--mode auto --types `"$typesStr`" --source `"$selectedAutoFolder`""
    }

    $btnGenerate.Enabled = $false
    $btnGenerate.BackColor = [System.Drawing.Color]::Gray
    $btnGenerate.Text = "Processing..."
    
    Log-Text "[INFO] Starting generator..."
    $worker.RunWorkerAsync($cmdArgs)
})

# Display form
$form.ShowDialog() | Out-Null
