<# ----- About: ----
    # New-ScheduledTaskFromScript - GUI Tool for Creating Scheduled Tasks from PowerShell Scripts
    # Revision v4.0 - 2026-03-11
    # Author: Eric Harless, Head Backup Nerd - N-able
    # GitHub: https://github.com/BackupNerd/Backup-Scripts
#>

<# ----- Legal: ----
    # Sample scripts are not supported under any N-able support program or service.
    # The sample scripts are provided AS IS without warranty of any kind.
    # N-able expressly disclaims all implied warranties including, warranties
    # of merchantability or of fitness for a particular purpose. 
    # In no event shall N-able or any other party be liable for damages arising
    # out of the use of or inability to use the sample scripts.
#>

<# ----- Compatibility: ----
    # For use with PowerShell 5.1 and PowerShell 7.x
    # Requires Windows with Task Scheduler
    # Requires Administrator privileges to create scheduled tasks
#>

<# ----- Behavior: ----
    # GUI application that:
    # 1. Allows file browser selection of PS1 script
    # 2. Automatically detects PowerShell version requirements (#Requires -Version)
    # 3. Parses script parameters and renders appropriate controls:
    #    - [switch] / [bool] / [int] 0/1 -> CheckBox (pre-checked from default value)
    #    - [ValidateSet] -> ComboBox picklist
    #    - $PSScriptRoot defaults -> gray placeholder (omitted from command; script uses runtime fallback)
    #    - All others -> TextBox
    #    - Shows parameter count: Script Parameters (exposed/total) with bold warning if mismatch
    # 4. Provides intuitive UI for:
    #    - Task name (auto-populated from script filename)
    #    - User account with password, or Run as SYSTEM
    #    - Run whether user is logged on or not
    #    - Run with highest privileges
    #    - Run hidden (no console window)
    #    - Schedule: Once, Daily, Weekly, At Startup, At Logon, Every X Minutes/Hours/Days
    # 5. Preview window shows exact PowerShell command that will be registered
    # 6. Test Run executes the command interactively with:
    #    - Header banner showing PowerShell version and executable path
    #    - Footer banner confirming test run complete
    #    - Window stays open (-NoExit) for review
    # 7. Creates Windows Scheduled Task with specified settings
    # 8. Always uses -Command mode (not -File) for correct [switch]/[bool] param handling
    # 9. Always prepends Set-Location to script directory so $PSScriptRoot resolves correctly
#>

#region ----- Script Location Change ----
    # CRITICAL: Change to script directory IMMEDIATELY to avoid auto-loading scripts from other directories
    # This must happen before any other code to prevent PowerShell from loading scripts in the current directory
    if ($PSScriptRoot) {
        Set-Location $PSScriptRoot
    }
#endregion

#region ----- Environment, Variables, Names and Paths ----
    # Ensure TLS 1.2 for any web operations
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Load Windows Forms assemblies
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Create global tooltip control for form elements
    $script:toolTip = New-Object System.Windows.Forms.ToolTip
    $script:toolTip.AutoPopDelay = 15000  # 15 seconds display time
    $script:toolTip.InitialDelay = 500    # 0.5 second before showing
    $script:toolTip.ReshowDelay = 200     # 0.2 second between tooltips
    $script:toolTip.ShowAlways = $true
    $script:toolTip.IsBalloon = $false
#endregion

#region ----- Functions ----

Function Get-ScriptRequirements {
    <#
    .SYNOPSIS
    Parses a PowerShell script to extract version requirements and parameters
    
    .PARAMETER ScriptPath
    Full path to the PowerShell script
    
    .OUTPUTS
    Hashtable with RequiredVersion, Parameters, and other metadata
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptPath
    )
    
    if (-not (Test-Path $ScriptPath)) {
        throw "Script not found: $ScriptPath"
    }
    
    $result = @{
        RequiredVersion = $null
        PowerShellEdition = 'Any'
        Parameters = @()
        HasCmdletBinding = $false
    }
    
    try {
        # Read script content
        $content = Get-Content -Path $ScriptPath -Raw
        
        # Check for #Requires statements
        if ($content -match '#Requires\s+-Version\s+(\d+(\.\d+)?)') {
            $result.RequiredVersion = $matches[1]
        }
        
        if ($content -match '#Requires\s+-PSEdition\s+(Core|Desktop)') {
            $result.PowerShellEdition = $matches[1]
        }
        
        # Check for CmdletBinding
        if ($content -match '\[CmdletBinding\(\)\]') {
            $result.HasCmdletBinding = $true
        }
        
        # Parse parameters using AST
        # Use ParseInput instead of ParseFile to avoid PS5 parameter transformation errors
        # MUST use -Encoding UTF8: PS5 defaults to system ANSI encoding, which misreads
        # Unicode characters (em dash, box-drawing chars, etc.) in string literals,
        # breaking the AST parser and silently truncating the parameter list.
        try {
            $content = Get-Content -Path $ScriptPath -Raw -Encoding UTF8 -ErrorAction Stop
            $tokens = $null
            $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($content, [ref]$tokens, [ref]$errors)
            
            # If there are parse errors, log them but continue
            if ($errors -and $errors.Count -gt 0) {
                Write-Host "Warning: Script has syntax issues. Parameter detection may be incomplete." -ForegroundColor Yellow
                foreach ($err in $errors | Select-Object -First 3) {
                    Write-Host "  $($err.Message)" -ForegroundColor Gray
                }
            }
        }
        catch {
            Write-Host "Warning: Failed to read script. Some parameters may not be detected." -ForegroundColor Yellow
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Gray
            $ast = $null
        }
        
        # Find ONLY the top-level script parameter block (not function parameters)
        # Get the script block AST (the root)
        if ($ast) {
            $scriptBlockAst = $ast.Find({
                $args[0] -is [System.Management.Automation.Language.ScriptBlockAst]
            }, $false)
        } else {
            $scriptBlockAst = $null
        }
        
        # Get only the parameter block directly under the script block (not nested in functions)
        if ($scriptBlockAst -and $scriptBlockAst.ParamBlock) {
            $paramBlock = $scriptBlockAst.ParamBlock
            
            foreach ($param in $paramBlock.Parameters) {
                # Skip AST error nodes — PS5 parser emits these with names like '__error__'
                # when it can't fully resolve a parameter declaration
                $paramName = $param.Name.VariablePath.UserPath
                if ([string]::IsNullOrWhiteSpace($paramName) -or $paramName -like '__*') { continue }

                $paramInfo = @{
                    Name = $paramName
                    Type = if ($param.StaticType) { $param.StaticType.Name } else { 'Object' }
                    Mandatory = $false
                    DefaultValue = $null
                    ValidateSetValues = $null
                    HelpMessage = $null
                }
                
                # Check for Mandatory, HelpMessage, and ValidateSet attributes
                foreach ($attribute in $param.Attributes) {
                    if ($attribute.TypeName.Name -eq 'Parameter') {
                        foreach ($arg in $attribute.NamedArguments) {
                            if ($arg.ArgumentName -eq 'Mandatory' -and $arg.Argument.VariablePath.UserPath -eq 'true') {
                                $paramInfo.Mandatory = $true
                            }
                            if ($arg.ArgumentName -eq 'HelpMessage') {
                                $paramInfo.HelpMessage = $arg.Argument.Value
                            }
                        }
                    }
                    if ($attribute.TypeName.Name -eq 'ValidateSet') {
                        $paramInfo.ValidateSetValues = @($attribute.PositionalArguments | ForEach-Object { $_.Value })
                    }
                }
                
                # Check for default value
                if ($param.DefaultValue) {
                    $paramInfo.DefaultValue = $param.DefaultValue.Extent.Text
                }
                
                $result.Parameters += $paramInfo
            }
        }
        
    } catch {
        Write-Warning "Failed to parse script: $($_.Exception.Message)"
    }
    
    return $result
}

Function Get-PowerShellExecutable {
    <#
    .SYNOPSIS
    Determines the appropriate PowerShell executable path based on requirements
    
    .PARAMETER RequiredVersion
    Minimum required PowerShell version
    
    .PARAMETER PowerShellEdition
    Required PowerShell edition (Core, Desktop, or Any)
    
    .OUTPUTS
    Full path to PowerShell executable
    #>
    param(
        [string]$RequiredVersion,
        [string]$PowerShellEdition = 'Any'
    )
    
    # PowerShell 7 locations
    $ps7Paths = @(
        "${env:ProgramFiles}\PowerShell\7\pwsh.exe",
        "${env:ProgramFiles(x86)}\PowerShell\7\pwsh.exe"
    )
    
    # PowerShell 5.1 (Windows PowerShell)
    $ps5Path = "${env:SystemRoot}\System32\WindowsPowerShell\v1.0\powershell.exe"
    
    # If Core edition required, use PowerShell 7
    if ($PowerShellEdition -eq 'Core') {
        foreach ($path in $ps7Paths) {
            if (Test-Path $path) {
                return $path
            }
        }
        throw "PowerShell Core (7.x) is required but not found"
    }
    
    # If Desktop edition required, use PowerShell 5.1
    if ($PowerShellEdition -eq 'Desktop') {
        if (Test-Path $ps5Path) {
            return $ps5Path
        }
        throw "PowerShell Desktop (5.1) is required but not found"
    }
    
    # If version 7+ required, use PowerShell 7
    if ($RequiredVersion -and [version]$RequiredVersion -ge [version]'7.0') {
        foreach ($path in $ps7Paths) {
            if (Test-Path $path) {
                return $path
            }
        }
        throw "PowerShell 7.x is required but not found"
    }
    
    # Default: prefer PowerShell 7 if available, otherwise use 5.1
    foreach ($path in $ps7Paths) {
        if (Test-Path $path) {
            return $path
        }
    }
    
    if (Test-Path $ps5Path) {
        return $ps5Path
    }
    
    throw "No PowerShell installation found"
}

Function New-ScheduledTaskGUI {
    <#
    .SYNOPSIS
    Creates and displays the main GUI for scheduled task creation
    #>
    
    # Create main form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Create Scheduled Task from PowerShell Script'
    $form.Size = New-Object System.Drawing.Size(780, 960)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.HelpButton = $true
    
    # Add help button click handler
    $form.Add_HelpButtonClicked({
        param($sender, $e)
        $aboutMsg = @"
Scheduled Task Creator for PowerShell Scripts
Version 2.1 - Updated January 2026

================================================================

Author: Eric Harless, Head Backup Nerd
Company: N-able
GitHub: @Backup_Nerd

This tool creates Windows scheduled tasks from
PowerShell scripts with automatic parameter
detection and PowerShell version compatibility.

Features:
- Auto-detect script parameters
- PowerShell 5.1 and 7.x support
- Run as SYSTEM or User account
- Flexible scheduling options
- Run hidden or interactive
- Test run with emulation warnings
- Tasks stored in \N-able\NerdScripts\

v2.1 New Features:
- Parameter tooltips from comment-based help
- Hover over parameters to see descriptions
- Requires .PARAMETER help in source scripts

v2.0 Features:
- SYSTEM account support (NT AUTHORITY\SYSTEM)
- Automatic user/password field toggling
- Enhanced preview with execution context
- Test run warnings for SYSTEM emulation
- DPAPI credential path display

================================================================

Sample scripts are not supported under any
N-able support program or service.
Use at your own risk.
"@
        [System.Windows.Forms.MessageBox]::Show($aboutMsg, 'About Scheduled Task Creator', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $e.Cancel = $true
    })
    
    # Current Y position for controls
    $yPos = 20
    
    #region Script Selection
    
    $lblScript = New-Object System.Windows.Forms.Label
    $lblScript.Location = New-Object System.Drawing.Point(20, $yPos)
    $lblScript.Size = New-Object System.Drawing.Size(100, 20)
    $lblScript.Text = 'Script Path:'
    $form.Controls.Add($lblScript)
    
    $txtScript = New-Object System.Windows.Forms.TextBox
    $txtScript.Location = New-Object System.Drawing.Point(130, $yPos)
    $txtScript.Size = New-Object System.Drawing.Size(510, 20)
    $txtScript.ReadOnly = $true
    $form.Controls.Add($txtScript)
    
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowseY = $yPos - 2
    $btnBrowse.Location = New-Object System.Drawing.Point(660, $btnBrowseY)
    $btnBrowse.Size = New-Object System.Drawing.Size(80, 24)
    $btnBrowse.Text = 'Browse...'
    $form.Controls.Add($btnBrowse)
    $script:toolTip.SetToolTip($btnBrowse, "Select a PowerShell (.ps1) script to schedule.`nScript parameters will be automatically detected and displayed below.")
    
    $yPos += 35
    
    # PowerShell Version Radio Buttons
    $grpPSVersion = New-Object System.Windows.Forms.GroupBox
    $grpPSVersion.Location = New-Object System.Drawing.Point(130, $yPos)
    $grpPSVersion.Size = New-Object System.Drawing.Size(620, 50)
    $grpPSVersion.Text = 'PowerShell Version'
    $form.Controls.Add($grpPSVersion)
    $script:toolTip.SetToolTip($grpPSVersion, "Choose which PowerShell version will execute this script in Task Scheduler.`nAuto-Detect: Uses script's #Requires directive or defaults to best available.`nForce PS 5.1: Windows PowerShell (built-in, more compatible with older modules).`nForce PS 7+: PowerShell Core (modern, cross-platform, better performance).")
    
    $radioAny = New-Object System.Windows.Forms.RadioButton
    $radioAny.Location = New-Object System.Drawing.Point(10, 20)
    $radioAny.Size = New-Object System.Drawing.Size(100, 20)
    $radioAny.Text = 'Auto-Detect'
    $radioAny.Checked = $true
    $grpPSVersion.Controls.Add($radioAny)
    $script:toolTip.SetToolTip($radioAny, "Automatically detect PowerShell version based on script requirements.`nRecommended for most scripts.")
    
    $radioPS5 = New-Object System.Windows.Forms.RadioButton
    $radioPS5.Location = New-Object System.Drawing.Point(120, 20)
    $radioPS5.Size = New-Object System.Drawing.Size(150, 20)
    $radioPS5.Text = 'Force PS 5.1'
    $grpPSVersion.Controls.Add($radioPS5)
    $script:toolTip.SetToolTip($radioPS5, "Force Windows PowerShell 5.1 (C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe).`nUse for: Legacy scripts, modules requiring PowerShell Desktop edition.`nNOTE: Will be disabled if script requires PowerShell 7+.")
    
    $radioPS7 = New-Object System.Windows.Forms.RadioButton
    $radioPS7.Location = New-Object System.Drawing.Point(280, 20)
    $radioPS7.Size = New-Object System.Drawing.Size(150, 20)
    $radioPS7.Text = 'Force PS 7+'
    $grpPSVersion.Controls.Add($radioPS7)
    $script:toolTip.SetToolTip($radioPS7, "Force PowerShell 7+ Core edition (C:\Program Files\PowerShell\7\pwsh.exe).`nUse for: Modern scripts, better performance, cross-platform compatibility.`nRequires PowerShell 7+ to be installed.")
    
    $yPos += 60
    
    # PowerShell Version Detection Label
    $lblPSVersion = New-Object System.Windows.Forms.Label
    $lblPSVersion.Location = New-Object System.Drawing.Point(130, $yPos)
    $lblPSVersion.Size = New-Object System.Drawing.Size(620, 20)
    $lblPSVersion.Text = 'No script selected'
    $lblPSVersion.ForeColor = [System.Drawing.Color]::Gray
    $form.Controls.Add($lblPSVersion)
    
    $yPos += 30
    
    #endregion
    
    #region Parameters Section
    
    $grpParams = New-Object System.Windows.Forms.GroupBox
    $grpParams.Location = New-Object System.Drawing.Point(20, $yPos)
    $grpParams.Size = New-Object System.Drawing.Size(730, 375)
    $grpParams.Text = 'Script Parameters'
    $form.Controls.Add($grpParams)
    
    # Scrollable panel for parameters
    $panelParams = New-Object System.Windows.Forms.Panel
    $panelParams.Location = New-Object System.Drawing.Point(10, 25)
    $panelParams.Size = New-Object System.Drawing.Size(710, 345)
    $panelParams.AutoScroll = $true
    $grpParams.Controls.Add($panelParams)
    
    $yPos += 385
    
    #endregion
    
    #region Task Settings
    
    $grpTask = New-Object System.Windows.Forms.GroupBox
    $grpTask.Location = New-Object System.Drawing.Point(20, $yPos)
    $grpTask.Size = New-Object System.Drawing.Size(730, 320)
    $grpTask.Text = 'Scheduled Task Settings'
    $form.Controls.Add($grpTask)
    
    $taskYPos = 25
    
    # Task Name
    $lblTaskName = New-Object System.Windows.Forms.Label
    $lblTaskName.Location = New-Object System.Drawing.Point(20, $taskYPos)
    $lblTaskName.Size = New-Object System.Drawing.Size(100, 20)
    $lblTaskName.Text = 'Task Name:'
    $grpTask.Controls.Add($lblTaskName)
    
    $txtTaskName = New-Object System.Windows.Forms.TextBox
    $txtTaskName.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $txtTaskName.Size = New-Object System.Drawing.Size(570, 20)
    $grpTask.Controls.Add($txtTaskName)
    $script:toolTip.SetToolTip($txtTaskName, "Friendly name for this scheduled task.`nTask will be stored in: \N-able\NerdScripts\[TaskName]`nExample: 'Daily Backup Report' or 'Sync Tickets Every Hour'")
    
    $taskYPos += 28
    
    # Run as SYSTEM checkbox
    $chkRunAsSystem = New-Object System.Windows.Forms.CheckBox
    $chkRunAsSystem.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $chkRunAsSystem.Size = New-Object System.Drawing.Size(570, 20)
    $chkRunAsSystem.Text = 'Run as SYSTEM (NT AUTHORITY\SYSTEM - for scripts requiring elevated file access)'
    $chkRunAsSystem.Checked = $true
    $grpTask.Controls.Add($chkRunAsSystem)
    $script:toolTip.SetToolTip($chkRunAsSystem, "Run as SYSTEM (NT AUTHORITY\SYSTEM) - Windows' highest privilege account.`n`nUSE WHEN:`n - Script needs full system file access (C:\Program Files, etc.)`n - Accessing services or system configuration`n - No interactive user session needed`n`nNO PASSWORD REQUIRED - SYSTEM is a built-in account.`n`nIMPORTANT: If your script uses DPAPI-encrypted credential files,`nthey must be created BY the SYSTEM account to be accessible.`n`nCAUTION: SYSTEM has unrestricted access to all system resources.")
    
    $taskYPos += 25
    
    # User Account
    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Location = New-Object System.Drawing.Point(20, $taskYPos)
    $lblUser.Size = New-Object System.Drawing.Size(100, 20)
    $lblUser.Text = 'Run As User:'
    $grpTask.Controls.Add($lblUser)
    
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $txtUser.Size = New-Object System.Drawing.Size(570, 20)
    $txtUser.Text = "$env:USERDOMAIN\$env:USERNAME"
    $grpTask.Controls.Add($txtUser)
    $script:toolTip.SetToolTip($txtUser, "User account that will run the scheduled task.`n`nFORMAT: DOMAIN\Username or COMPUTERNAME\Username`nEXAMPLE: CORP\eric.harless or WORKSTATION\Administrator`n`nNOTE:`n - Account must have appropriate permissions for script operations`n - For local accounts use: COMPUTERNAME\Username`n - If script uses DPAPI credential files, they must be created by this user")
    
    $taskYPos += 28
    
    # Password
    $lblPassword = New-Object System.Windows.Forms.Label
    $lblPassword.Location = New-Object System.Drawing.Point(20, $taskYPos)
    $lblPassword.Size = New-Object System.Drawing.Size(100, 20)
    $lblPassword.Text = 'Password:'
    $grpTask.Controls.Add($lblPassword)
    
    $txtPassword = New-Object System.Windows.Forms.TextBox
    $txtPassword.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $txtPassword.Size = New-Object System.Drawing.Size(570, 20)
    $txtPassword.PasswordChar = '*'
    $grpTask.Controls.Add($txtPassword)
    $script:toolTip.SetToolTip($txtPassword, "Password for the user account.`n`nREQUIRED WHEN:`n - 'Run whether user is logged on or not' is checked`n - 'Run with highest privileges' is checked`n`nNOT REQUIRED:`n - Running as SYSTEM (no password needed)`n - Only when user is logged on (interactive mode)`n`nSECURITY:`nPassword is encrypted by Task Scheduler and stored in Windows Credential Manager.`nThis is separate from any credentials your script might use.")
    
    # Add event handler for Run as SYSTEM checkbox
    $chkRunAsSystem.Add_CheckedChanged({
        if ($chkRunAsSystem.Checked) {
            # Disable user/password fields when SYSTEM is selected
            $txtUser.Enabled = $false
            $txtPassword.Enabled = $false
            $chkRunLoggedOff.Enabled = $false
            $txtUser.Text = 'NT AUTHORITY\SYSTEM'
            $txtPassword.Text = ''
            $chkRunLoggedOff.Checked = $true  # SYSTEM always runs whether user is logged on or not
        } else {
            # Re-enable user/password fields
            $txtUser.Enabled = $true
            $txtPassword.Enabled = $true
            $chkRunLoggedOff.Enabled = $true
            $txtUser.Text = "$env:USERDOMAIN\$env:USERNAME"
        }
    })
    
    $taskYPos += 28
    
    # Run whether user is logged on
    $chkRunLoggedOff = New-Object System.Windows.Forms.CheckBox
    $chkRunLoggedOff.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $chkRunLoggedOff.Size = New-Object System.Drawing.Size(570, 20)
    $chkRunLoggedOff.Text = 'Run whether user is logged on or not (requires password)'
    $chkRunLoggedOff.Checked = $false
    $grpTask.Controls.Add($chkRunLoggedOff)
    $script:toolTip.SetToolTip($chkRunLoggedOff, "Task runs in background even when user is NOT logged in.`n`nCHECKED (Background Mode):`n - Task runs whether user is logged on or not`n - Runs in non-interactive session (no desktop access)`n - REQUIRES password`n - Ideal for server automation and unattended scripts`n - Task Scheduler LogonType: Password`n`nUNCHECKED (Interactive Mode):`n - Task only runs when user is logged in`n - Can display UI and interact with desktop`n - Password optional`n - Task Scheduler LogonType: Interactive Only`n`nREQUIRES: Password must be provided for background mode.")
    
    $taskYPos += 25
    
    # Run with highest privileges
    $chkHighestPrivileges = New-Object System.Windows.Forms.CheckBox
    $chkHighestPrivileges.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $chkHighestPrivileges.Size = New-Object System.Drawing.Size(570, 20)
    $chkHighestPrivileges.Text = 'Run with highest privileges'
    $chkHighestPrivileges.Checked = $false
    $grpTask.Controls.Add($chkHighestPrivileges)
    $script:toolTip.SetToolTip($chkHighestPrivileges, "Run with Administrator privileges (bypasses UAC).`n`nCHECKED (Elevated/Administrator):`n - Script runs with full administrative rights`n - Bypasses User Account Control (UAC)`n - Can modify system files, registry, services`n - Task Scheduler RunLevel: Highest`n - REQUIRES password for user accounts`n`nUNCHECKED (Standard User):`n - Script runs with limited user permissions`n - Subject to UAC prompts`n - Cannot modify protected system resources`n - Task Scheduler RunLevel: Limited`n`nRECOMMENDED: Keep checked for scripts that modify system configuration.`nNOTE: SYSTEM account always runs with highest privileges.")
    
    $taskYPos += 25
    
    # Run hidden (no window)
    $chkRunHidden = New-Object System.Windows.Forms.CheckBox
    $chkRunHidden.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $chkRunHidden.Size = New-Object System.Drawing.Size(570, 20)
    $chkRunHidden.Text = 'Run hidden (no PowerShell window - prevents focus stealing)'
    $chkRunHidden.Checked = $false
    $grpTask.Controls.Add($chkRunHidden)
    $script:toolTip.SetToolTip($chkRunHidden, "Hide PowerShell console window during execution.`n`nCHECKED (Hidden):`n - PowerShell runs with -WindowStyle Hidden`n - No console window appears`n - Prevents focus stealing from active applications`n - Ideal for background automation`n - Script output logged to Task Scheduler History`n`nUNCHECKED (Visible):`n - PowerShell console window appears`n - Can see script output in real-time`n - Useful for debugging or user-facing scripts`n - Window may steal focus from other applications`n`nRECOMMENDED: Keep checked for unattended automation.`nNOTE: Test Run button ALWAYS shows window (ignores this setting).")

    # Apply initial UI state — all controls now exist so we can safely set dependent state.
    # CheckedChanged does not fire when .Checked is set in code before the handler is attached,
    # so we apply the SYSTEM-checked defaults explicitly here.
    if ($chkRunAsSystem.Checked) {
        $txtUser.Text             = 'NT AUTHORITY\SYSTEM'
        $txtUser.Enabled          = $false
        $txtPassword.Enabled      = $false
        $chkRunLoggedOff.Checked  = $true
        $chkRunLoggedOff.Enabled  = $false
    }
    
    $taskYPos += 28
    
    # Schedule Type
    $lblSchedule = New-Object System.Windows.Forms.Label
    $lblSchedule.Location = New-Object System.Drawing.Point(20, $taskYPos)
    $lblSchedule.Size = New-Object System.Drawing.Size(100, 20)
    $lblSchedule.Text = 'Schedule:'
    $grpTask.Controls.Add($lblSchedule)
    
    $cmbSchedule = New-Object System.Windows.Forms.ComboBox
    $cmbSchedule.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $cmbSchedule.Size = New-Object System.Drawing.Size(150, 20)
    $cmbSchedule.DropDownStyle = 'DropDownList'
    @('Once', 'Daily', 'Hourly', 'Weekly', 'Every X Minutes', 'Every X Hours', 'Every X Days', 'At Startup', 'At Logon') | ForEach-Object { $cmbSchedule.Items.Add($_) } | Out-Null
    $cmbSchedule.SelectedIndex = 1
    $grpTask.Controls.Add($cmbSchedule)
    $script:toolTip.SetToolTip($cmbSchedule, "Trigger schedule for task execution.`n`nOPTIONS:`n - Once: Run one time at specified date/time`n - Daily: Run every day at specified time`n - Hourly: Run every hour starting at specified time`n - Weekly: Run on selected days of week`n - Every X Minutes: Run at custom minute interval`n - Every X Hours: Run at custom hour interval`n - Every X Days: Run at custom day interval`n - At Startup: Run when Windows starts (before login)`n - At Logon: Run when user logs in`n`nNOTE: Multiple triggers can be added in Task Scheduler after creation.")
    
    $taskYPos += 28
    
    # Start Date/Time
    $lblStartTime = New-Object System.Windows.Forms.Label
    $lblStartTime.Location = New-Object System.Drawing.Point(20, $taskYPos)
    $lblStartTime.Size = New-Object System.Drawing.Size(100, 20)
    $lblStartTime.Text = 'Start Date/Time:'
    $grpTask.Controls.Add($lblStartTime)
    
    $dtpStartDate = New-Object System.Windows.Forms.DateTimePicker
    $dtpStartDate.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $dtpStartDate.Size = New-Object System.Drawing.Size(150, 20)
    $dtpStartDate.Format = 'Short'
    $grpTask.Controls.Add($dtpStartDate)
    $script:toolTip.SetToolTip($dtpStartDate, "Start date for the scheduled task.`n`nFor recurring tasks (Daily, Weekly, etc.), this is the FIRST occurrence.`nFor 'Once' tasks, this is the ONLY execution date.`n`nNOTE: Task will not run before this date/time.")
    
    $dtpStartTime = New-Object System.Windows.Forms.DateTimePicker
    $dtpStartTime.Location = New-Object System.Drawing.Point(290, $taskYPos)
    $dtpStartTime.Size = New-Object System.Drawing.Size(100, 20)
    $dtpStartTime.Format = 'Time'
    $dtpStartTime.ShowUpDown = $true
    $grpTask.Controls.Add($dtpStartTime)
    $script:toolTip.SetToolTip($dtpStartTime, "Start time for the scheduled task.`n`nFor recurring tasks, this is the daily execution time.`nFor interval-based tasks (Every X Minutes/Hours), this is when the interval starts.`n`nEXAMPLE:`n - Daily at 2:00 AM: Runs every day at 2:00 AM`n - Every 30 Minutes starting at 8:00 AM: First run at 8:00 AM, then 8:30 AM, 9:00 AM, etc.")
    
    $taskYPos += 28
    
    # Interval controls (initially hidden)
    $lblInterval = New-Object System.Windows.Forms.Label
    $lblInterval.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $lblInterval.Size = New-Object System.Drawing.Size(80, 20)
    $lblInterval.Text = 'Every:'
    $lblInterval.Visible = $false
    $grpTask.Controls.Add($lblInterval)
    
    $numInterval = New-Object System.Windows.Forms.NumericUpDown
    $numInterval.Location = New-Object System.Drawing.Point(210, $taskYPos)
    $numInterval.Size = New-Object System.Drawing.Size(60, 20)
    $numInterval.Minimum = 1
    $numInterval.Maximum = 999
    $numInterval.Value = 1
    $numInterval.Visible = $false
    $grpTask.Controls.Add($numInterval)
    $script:toolTip.SetToolTip($numInterval, "Custom interval value for repetition.`n`nEXAMPLES:`n - Every 15 Minutes: Enter 15`n - Every 4 Hours: Enter 4`n - Every 7 Days: Enter 7`n`nMINIMUM: 1`nMAXIMUM: Varies by unit (1440 min, 168 hrs, 365 days)`n`nNOTE: Repetition starts at the specified Start Date/Time.")
    
    $lblIntervalUnit = New-Object System.Windows.Forms.Label
    $intervalUnitY = $taskYPos + 3
    $lblIntervalUnit.Location = New-Object System.Drawing.Point(275, $intervalUnitY)
    $lblIntervalUnit.Size = New-Object System.Drawing.Size(100, 20)
    $lblIntervalUnit.Text = 'minutes'
    $lblIntervalUnit.Visible = $false
    $grpTask.Controls.Add($lblIntervalUnit)
    
    $taskYPos += 28
    
    # Weekly options (initially hidden)
    $lblWeekdays = New-Object System.Windows.Forms.Label
    $lblWeekdays.Location = New-Object System.Drawing.Point(130, $taskYPos)
    $lblWeekdays.Size = New-Object System.Drawing.Size(100, 20)
    $lblWeekdays.Text = 'Days of Week:'
    $lblWeekdays.Visible = $false
    $grpTask.Controls.Add($lblWeekdays)
    
    $clbWeekdays = New-Object System.Windows.Forms.CheckedListBox
    $weekdaysY = $taskYPos + 20
    $clbWeekdays.Location = New-Object System.Drawing.Point(130, $weekdaysY)
    $clbWeekdays.Size = New-Object System.Drawing.Size(570, 70)
    @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday') | ForEach-Object { $clbWeekdays.Items.Add($_) } | Out-Null
    $clbWeekdays.Visible = $false
    $grpTask.Controls.Add($clbWeekdays)
    $script:toolTip.SetToolTip($clbWeekdays, "Select which days of the week the task should run.`n`nCheck one or more days for weekly execution.`n`nEXAMPLES:`n - Business Days: Mon, Tue, Wed, Thu, Fri`n - Weekends: Sat, Sun`n - Custom: Any combination`n`nNOTE: Task runs at the specified Start Time on selected days.")
    
    #endregion
    
    #region Buttons
    
    $yPos += 330
    
    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Location = New-Object System.Drawing.Point(440, $yPos)
    $btnPreview.Size = New-Object System.Drawing.Size(90, 35)
    $btnPreview.Text = 'Preview'
    $btnPreview.Enabled = $false
    $form.Controls.Add($btnPreview)
    $script:toolTip.SetToolTip($btnPreview, "Preview the scheduled task configuration before creating it.`n`nShows:`n - Task name and path`n - Execution context (user/SYSTEM)`n - PowerShell executable and arguments`n - Full command line`n - Schedule and trigger settings`n`nACTIONS:`n - Test Run: Execute script NOW as current user (for testing)`n - Copy: Copy command line to clipboard`n`nNOTE: Test Run runs as current user, NOT as scheduled task user.")
    
    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Location = New-Object System.Drawing.Point(540, $yPos)
    $btnCreate.Size = New-Object System.Drawing.Size(90, 35)
    $btnCreate.Text = 'Create Task'
    $btnCreate.Enabled = $false
    $form.Controls.Add($btnCreate)
    $script:toolTip.SetToolTip($btnCreate, "Create the scheduled task in Windows Task Scheduler.`n`nTASK LOCATION: \N-able\NerdScripts\[TaskName]`n`nVALIDATIONS:`n - Task name required`n - Password required for elevated/background tasks`n - Mandatory script parameters must be filled`n - User credentials verified before creation`n`nAFTER CREATION:`n - Task registered in Task Scheduler`n - Can be managed via taskschd.msc`n - History logged in Event Viewer`n`nNOTE: Requires Administrator privileges to create tasks.")
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(640, $yPos)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 35)
    $btnCancel.Text = 'Cancel'
    $btnCancel.DialogResult = 'Cancel'
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel
    $script:toolTip.SetToolTip($btnCancel, "Cancel and close without creating a scheduled task.")
    
    # Ensure form height accommodates all controls with bottom padding
    $requiredHeight = $yPos + 35 + 50  # Button height + bottom padding
    if ($form.Height -lt $requiredHeight) {
        $form.Height = $requiredHeight
    }
    
    #endregion
    
    #region Event Handlers
    
    # Script variable to hold parsed script info
    $script:scriptInfo = $null
    $script:paramControls = @{}
    $script:paramTypes    = @{}   # maps param name -> 'SwitchParameter' | 'Boolean' | etc.
    
    # Helper function to update PowerShell version display
    Function Update-PSVersionDisplay {
        if (-not $script:scriptInfo) { return }
        
        try {
            # Check if script requires PS7+ (disable PS5 option if incompatible)
            $requiresPS7 = $false
            if ($script:scriptInfo.RequiredVersion -and [version]$script:scriptInfo.RequiredVersion -ge [version]'7.0') {
                $requiresPS7 = $true
            }
            if ($script:scriptInfo.PowerShellEdition -eq 'Core') {
                $requiresPS7 = $true
            }
            
            # Enable/disable PS5 radio button based on compatibility
            if ($requiresPS7) {
                $radioPS5.Enabled = $false
                $radioPS5.ForeColor = [System.Drawing.Color]::Gray
                # If PS5 was selected, auto-switch to PS7
                if ($radioPS5.Checked) {
                    $radioPS7.Checked = $true
                }
            } else {
                $radioPS5.Enabled = $true
                $radioPS5.ForeColor = [System.Drawing.Color]::Black
            }
            
            $psExe = $null
            $versionText = ""
            
            if ($radioPS5.Checked) {
                $psExe = Get-PowerShellExecutable -RequiredVersion '5.1' -PowerShellEdition 'Desktop'
                $versionText = "PowerShell 5.1 (Forced) - Using: $psExe"
            } elseif ($radioPS7.Checked) {
                $psExe = Get-PowerShellExecutable -RequiredVersion '7.0' -PowerShellEdition 'Core'
                $versionText = "PowerShell 7+ (Forced) - Using: $psExe"
            } else {
                # Auto-detect
                $versionText = "PowerShell "
                if ($script:scriptInfo.RequiredVersion) {
                    $versionText += "$($script:scriptInfo.RequiredVersion)+ "
                } else {
                    $versionText += "Any Version "
                }
                
                if ($script:scriptInfo.PowerShellEdition -ne 'Any') {
                    $versionText += "($($script:scriptInfo.PowerShellEdition) Edition)"
                }
                
                $psExe = Get-PowerShellExecutable -RequiredVersion $script:scriptInfo.RequiredVersion -PowerShellEdition $script:scriptInfo.PowerShellEdition
                $versionText += " - Using: $psExe"
            }
            
            $lblPSVersion.Text = $versionText
            $lblPSVersion.ForeColor = [System.Drawing.Color]::Green
        } catch {
            $lblPSVersion.Text = "ERROR: $_"
            $lblPSVersion.ForeColor = [System.Drawing.Color]::Red
        }
    }
    
    # Radio button event handlers
    $radioAny.Add_CheckedChanged({ if ($radioAny.Checked) { Update-PSVersionDisplay } })
    $radioPS5.Add_CheckedChanged({ if ($radioPS5.Checked) { Update-PSVersionDisplay } })
    $radioPS7.Add_CheckedChanged({ if ($radioPS7.Checked) { Update-PSVersionDisplay } })
    
    # Browse button click
    $btnBrowse.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = 'PowerShell Scripts (*.ps1)|*.ps1|All Files (*.*)|*.*'
        $openFileDialog.Title = 'Select PowerShell Script'
        $openFileDialog.InitialDirectory = $PSScriptRoot
        
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            $txtScript.Text = $openFileDialog.FileName
            
            try {
                # Parse script
                $script:scriptInfo = Get-ScriptRequirements -ScriptPath $openFileDialog.FileName
                
                # Update version display using helper function
                Update-PSVersionDisplay
                
                # Clear existing parameter controls
                $panelParams.Controls.Clear()
                $grpParams.Text = 'Script Parameters'
                $script:paramControls = @{}
                $script:paramTypes    = @{}
                
                # Create tooltip control for parameter help
                $toolTip = New-Object System.Windows.Forms.ToolTip
                $toolTip.AutoPopDelay = 10000  # 10 seconds
                $toolTip.InitialDelay = 500    # 0.5 second
                $toolTip.ReshowDelay = 200     # 0.2 second
                $toolTip.ShowAlways = $true
                
                # Try to get help information from the script
                $scriptHelp = $null
                try {
                    $scriptHelp = Get-Help -Name $openFileDialog.FileName -ErrorAction SilentlyContinue
                } catch {
                    # Script doesn't have comment-based help, tooltips won't be available
                }
                
                # Create parameter controls
                $paramY = 10
                foreach ($param in $script:scriptInfo.Parameters) {
                    # Skip parameters without a name
                    if ([string]::IsNullOrWhiteSpace($param.Name)) {
                        continue
                    }
                    
                    # Parameter label with type inline
                    $lblParam = New-Object System.Windows.Forms.Label
                    $lblParam.Location = New-Object System.Drawing.Point(10, $paramY)
                    $lblParam.Size = New-Object System.Drawing.Size(150, 20)
                    $paramText = "-$($param.Name)"
                    if ($param.Mandatory) {
                        $paramText += ' *'
                        $lblParam.ForeColor = [System.Drawing.Color]::Red
                    }
                    $lblParam.Text = $paramText
                    $panelParams.Controls.Add($lblParam)
                    
                    # Type label (inline, before control)
                    $lblType = New-Object System.Windows.Forms.Label
                    $lblType.Location = New-Object System.Drawing.Point(170, $paramY)
                    $lblType.Size = New-Object System.Drawing.Size(80, 20)
                    # Shorten type names to prevent wrapping
                    $typeDisplay = $param.Type
                    if ($typeDisplay -eq 'SwitchParameter') { $typeDisplay = 'Switch' }
                    elseif ($typeDisplay -eq 'String') { $typeDisplay = 'Str' }
                    elseif ($typeDisplay -eq 'Int32') { $typeDisplay = 'Int' }
                    elseif ($typeDisplay -eq 'Boolean') { $typeDisplay = 'Bool' }
                    $lblType.Text = "[$typeDisplay]"
                    $lblType.ForeColor = [System.Drawing.Color]::Gray
                    $lblType.Font = New-Object System.Drawing.Font($lblType.Font.FontFamily, 7.5)
                    $panelParams.Controls.Add($lblType)
                    
                    # Parameter value control
                    if ($param.ValidateSetValues) {
                        # ValidateSet → render as a ComboBox picklist
                        $cmbParam = New-Object System.Windows.Forms.ComboBox
                        $cmbParam.Location = New-Object System.Drawing.Point(260, $paramY)
                        $cmbParam.Size = New-Object System.Drawing.Size(350, 20)
                        $cmbParam.DropDownStyle = 'DropDownList'
                        foreach ($setVal in $param.ValidateSetValues) {
                            $cmbParam.Items.Add($setVal) | Out-Null
                        }
                        # Pre-select default value if present
                        if ($param.DefaultValue) {
                            $defaultClean = $param.DefaultValue.Trim('"', "'")
                            $idx = $cmbParam.Items.IndexOf($defaultClean)
                            if ($idx -ge 0) { $cmbParam.SelectedIndex = $idx }
                            elseif ($cmbParam.Items.Count -gt 0) { $cmbParam.SelectedIndex = 0 }
                        } elseif ($cmbParam.Items.Count -gt 0) {
                            $cmbParam.SelectedIndex = 0
                        }
                        $panelParams.Controls.Add($cmbParam)
                        $script:paramControls[$param.Name] = $cmbParam

                        # Attach tooltip: prefer comment-based help (.PARAMETER), fall back to HelpMessage attribute
                        $helpText = $null
                        if ($scriptHelp -and $scriptHelp.parameters.parameter) {
                            $paramHelp = $scriptHelp.parameters.parameter | Where-Object { $_.name -eq $param.Name }
                            if ($paramHelp -and $paramHelp.description.Text) {
                                $helpText = ($paramHelp.description.Text -join " ").Trim()
                            }
                        }
                        if ([string]::IsNullOrWhiteSpace($helpText) -and -not [string]::IsNullOrWhiteSpace($param.HelpMessage)) {
                            $helpText = $param.HelpMessage
                        }
                        if (-not [string]::IsNullOrWhiteSpace($helpText)) {
                            $toolTip.SetToolTip($cmbParam, $helpText)
                        }
                    } elseif ($param.Type -eq 'Boolean' -or $param.Type -eq 'SwitchParameter' -or
                              ($param.Type -eq 'Int32' -and ($param.DefaultValue -eq '0' -or $param.DefaultValue -eq '1'))) {
                        # [bool], [switch], or [int] 0/1 flag — render as checkbox
                        $chkParam = New-Object System.Windows.Forms.CheckBox
                        $chkParam.Location = New-Object System.Drawing.Point(260, $paramY)
                        $chkParam.Size = New-Object System.Drawing.Size(350, 20)
                        # Pre-check: [bool]/$true default, or [int] with 1 default ([switch] has no default -> unchecked)
                        $chkParam.Checked = ($param.DefaultValue -eq '$true' -or $param.DefaultValue -eq '1')
                        $panelParams.Controls.Add($chkParam)
                        $script:paramControls[$param.Name] = $chkParam
                        $script:paramTypes[$param.Name]    = $param.Type   # 'Boolean', 'SwitchParameter', or 'Int32'

                        # Attach tooltip: prefer comment-based help (.PARAMETER), fall back to HelpMessage attribute
                        $helpText = $null
                        if ($scriptHelp -and $scriptHelp.parameters.parameter) {
                            $paramHelp = $scriptHelp.parameters.parameter | Where-Object { $_.name -eq $param.Name }
                            if ($paramHelp -and $paramHelp.description.Text) {
                                $helpText = ($paramHelp.description.Text -join " ").Trim()
                            }
                        }
                        if ([string]::IsNullOrWhiteSpace($helpText) -and -not [string]::IsNullOrWhiteSpace($param.HelpMessage)) {
                            $helpText = $param.HelpMessage
                        }
                        if (-not [string]::IsNullOrWhiteSpace($helpText)) {
                            $toolTip.SetToolTip($chkParam, $helpText)
                        }
                    } else {
                        $txtParam = New-Object System.Windows.Forms.TextBox
                        $txtParam.Location = New-Object System.Drawing.Point(260, $paramY)
                        $txtParam.Size = New-Object System.Drawing.Size(350, 20)
                        if ($param.DefaultValue) {
                            # Check if default is $PSScriptRoot - if so, leave empty (script will use its own default)
                            if ($param.DefaultValue -match '\$PSScriptRoot') {
                                # PS5-compatible placeholder (PlaceholderText is PS6+/WinForms .NET5+)
                                # Use gray text + Tag + Enter/Leave events to simulate it
                                $placeholderStr = '(uses $PSScriptRoot from script)'
                                $txtParam.Text      = $placeholderStr
                                $txtParam.ForeColor = [System.Drawing.SystemColors]::GrayText
                                $txtParam.Tag       = 'placeholder'
                                $txtParam.Add_Enter({
                                    if ($this.Tag -eq 'placeholder') {
                                        $this.Text      = ''
                                        $this.ForeColor = [System.Drawing.SystemColors]::WindowText
                                        $this.Tag       = $null
                                    }
                                })
                                $txtParam.Add_Leave({
                                    if ([string]::IsNullOrWhiteSpace($this.Text)) {
                                        $this.Text      = '(uses $PSScriptRoot from script)'
                                        $this.ForeColor = [System.Drawing.SystemColors]::GrayText
                                        $this.Tag       = 'placeholder'
                                    }
                                })
                            } else {
                                $txtParam.Text = $param.DefaultValue.Trim('"', "'", '$')
                            }
                        }
                        $panelParams.Controls.Add($txtParam)
                        $script:paramControls[$param.Name] = $txtParam
                        
                        # Attach tooltip: prefer comment-based help (.PARAMETER), fall back to HelpMessage attribute
                        $helpText = $null
                        if ($scriptHelp -and $scriptHelp.parameters.parameter) {
                            $paramHelp = $scriptHelp.parameters.parameter | Where-Object { $_.name -eq $param.Name }
                            if ($paramHelp -and $paramHelp.description.Text) {
                                $helpText = ($paramHelp.description.Text -join " ").Trim()
                            }
                        }
                        if ([string]::IsNullOrWhiteSpace($helpText) -and -not [string]::IsNullOrWhiteSpace($param.HelpMessage)) {
                            $helpText = $param.HelpMessage
                        }
                        if (-not [string]::IsNullOrWhiteSpace($helpText)) {
                            $toolTip.SetToolTip($txtParam, $helpText)
                        }
                    }
                    
                    $paramY += 25
                }

                # Update GroupBox title with exposed/total param count
                $totalDetected = ($script:scriptInfo.Parameters | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) }).Count
                $totalExposed  = $script:paramControls.Count
                $grpParams.Text = "Script Parameters  ($totalExposed/$totalDetected)"

                # Bold warning if any params were detected but not rendered
                if ($totalExposed -ne $totalDetected) {
                    $lblParamMismatch = New-Object System.Windows.Forms.Label
                    $lblParamMismatch.Location = New-Object System.Drawing.Point(10, $paramY)
                    $lblParamMismatch.Size = New-Object System.Drawing.Size(680, 20)
                    $lblParamMismatch.Text = [char]0x26A0 + "  $($totalDetected - $totalExposed) parameter(s) detected in script but not shown - unsupported type or parse issue."
                    $lblParamMismatch.ForeColor = [System.Drawing.Color]::DarkOrange
                    $lblParamMismatch.Font = New-Object System.Drawing.Font($lblParamMismatch.Font, [System.Drawing.FontStyle]::Bold)
                    $panelParams.Controls.Add($lblParamMismatch)
                }

                if ($script:scriptInfo.Parameters.Count -eq 0) {
                    $lblNoParams = New-Object System.Windows.Forms.Label
                    $lblNoParams.Location = New-Object System.Drawing.Point(10, 10)
                    $lblNoParams.Size = New-Object System.Drawing.Size(600, 20)
                    $lblNoParams.Text = 'This script has no parameters'
                    $lblNoParams.ForeColor = [System.Drawing.Color]::Gray
                    $panelParams.Controls.Add($lblNoParams)
                }
                
                # Auto-populate task name
                if ([string]::IsNullOrWhiteSpace($txtTaskName.Text)) {
                    $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($openFileDialog.FileName)
                    $txtTaskName.Text = "Run $scriptName"
                }
                
                # Enable create and preview buttons
                $btnCreate.Enabled = $true
                $btnPreview.Enabled = $true
                
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error parsing script: $($_.Exception.Message)", "Error", 'OK', 'Error')
                $lblPSVersion.Text = "Error: $($_.Exception.Message)"
                $lblPSVersion.ForeColor = [System.Drawing.Color]::Red
            }
        }
    })
    
    # Schedule type change
    $cmbSchedule.Add_SelectedIndexChanged({
        $isWeekly = $cmbSchedule.SelectedItem -eq 'Weekly'
        $isInterval = $cmbSchedule.SelectedItem -in @('Every X Minutes', 'Every X Hours', 'Every X Days')
        $showDateTime = $cmbSchedule.SelectedItem -notin @('At Startup', 'At Logon')
        
        # Update interval unit label
        if ($cmbSchedule.SelectedItem -eq 'Every X Minutes') {
            $lblIntervalUnit.Text = 'minutes'
            $numInterval.Maximum = 1440  # Max 1440 minutes (24 hours)
        } elseif ($cmbSchedule.SelectedItem -eq 'Every X Hours') {
            $lblIntervalUnit.Text = 'hours'
            $numInterval.Maximum = 168  # Max 168 hours (7 days)
        } elseif ($cmbSchedule.SelectedItem -eq 'Every X Days') {
            $lblIntervalUnit.Text = 'days'
            $numInterval.Maximum = 365  # Max 365 days
        }
        
        $lblWeekdays.Visible = $isWeekly
        $clbWeekdays.Visible = $isWeekly
        $lblInterval.Visible = $isInterval
        $numInterval.Visible = $isInterval
        $lblIntervalUnit.Visible = $isInterval
        $lblStartTime.Visible = $showDateTime
        $dtpStartDate.Visible = $showDateTime
        $dtpStartTime.Visible = $showDateTime
    })
    
    # ---- Param Compatibility Crib Notes (for scripts registered via this GUI) ----
    # This GUI ALWAYS uses -Command mode, which evaluates PS expressions. Rules below
    # reflect how this GUI generates commands AND how target scripts should be authored.
    #
    # [switch]  PREFERRED. Detected by param.Type='SwitchParameter'. Rendered as checkbox.
    #           Default: always unchecked (no default value in param block).
    #           Emit: checked -> "-Name"  |  unchecked -> omit entirely (NEVER pass -Name $false).
    #
    # [bool]    Detected by param.Type='Boolean'. Rendered as checkbox.
    #           Default: '$true' -> checked, '$false' or absent -> unchecked.
    #           Emit: ([bool]$true) / ([bool]$false)  (explicit cast required for -Command mode).
    #           WARN: [bool] breaks in -File mode. Prefer [switch] in new scripts.
    #
    # [int] 0/1 Detected if param.Type='Int32' AND default='0' or '1'. Rendered as checkbox.
    #           Default: '1' -> checked, '0' -> unchecked.
    #           Emit: 1 / 0  (integer literal, no quotes).
    #
    # [string] + $PSScriptRoot default: rendered as gray placeholder, omitted from command.
    #           Target script must have runtime fallback: if (-not $Var) { $Var = $PSScriptRoot }
    #
    # Set-Location is ALWAYS prepended to command so $PSScriptRoot resolves in target script.
    # Full reference: .github\instructions\15-param-compatibility-ps5-ps7.md
    # ---------------------------------------------------------------------------------

    # Function to build command arguments (used by both preview and create)
    Function Build-CommandArguments {
        param([bool]$IncludeWindowStyle = $true)
        
        # Build using -Command with proper variable casting for bool parameters
        # AND Set-Location to the script directory so $PSScriptRoot works correctly
        $scriptPath = $txtScript.Text
        $scriptDirectory = Split-Path -Parent $scriptPath
        
        # Start building the command script block
        # CRITICAL: Set location to script directory BEFORE running script so $PSScriptRoot-based paths work
        $commandScript = "Set-Location '$scriptDirectory'; & '$scriptPath'"
        
        # Add parameters with proper casting
        foreach ($paramName in $script:paramControls.Keys) {
            $control = $script:paramControls[$paramName]
            
            if ($control -is [System.Windows.Forms.CheckBox]) {
                if ($script:paramTypes[$paramName] -eq 'SwitchParameter') {
                    # [switch] params: emit -Name when checked, omit entirely when unchecked
                    # (passing -Name ([bool]$false) would set the switch $true because it's present)
                    if ($control.Checked) {
                        $commandScript += " -${paramName}"
                    }
                } elseif ($script:paramTypes[$paramName] -eq 'Int32') {
                    # [int] 0/1 flag: pass 1 (checked) or 0 (unchecked)
                    # In -Command mode these are integers, not strings, so PS5 coercion works
                    $intValue = if ($control.Checked) { '1' } else { '0' }
                    $commandScript += " -${paramName} $intValue"
                } else {
                    # [bool] params: always emit with explicit cast so task scheduler passes $true/$false correctly
                    $boolValue = if ($control.Checked) { 'true' } else { 'false' }
                    $commandScript += " -${paramName} ([bool]`$$boolValue)"
                }
            } elseif ($control -is [System.Windows.Forms.ComboBox]) {
                # ValidateSet picklist — pass the selected item as a quoted string
                if ($null -ne $control.SelectedItem) {
                    $value = $control.SelectedItem.ToString() -replace "'", "''"
                    $commandScript += " -${paramName} '$value'"
                }
            } elseif ($control -is [System.Windows.Forms.TextBox]) {
                # Skip if still showing gray placeholder (Tag='placeholder') — let script use its own default
                if ($control.Tag -eq 'placeholder') { continue }
                if (-not [string]::IsNullOrWhiteSpace($control.Text)) {
                    $value = $control.Text
                    # Skip parameters that are still showing the default placeholder values
                    # This allows the script to use its own defaults (like $PSScriptRoot)
                    if ($value -eq 'PSScriptRoot' -or $value -eq '$PSScriptRoot') {
                        # Don't pass this parameter - let script use its default
                        continue
                    }
                    # Escape single quotes in the value
                    $value = $value -replace "'", "''"
                    $commandScript += " -${paramName} '$value'"
                }
            }
        }
        
        # Build the full argument string
        $arguments = "-NoProfile -ExecutionPolicy Bypass"
        
        if ($IncludeWindowStyle -and $chkRunHidden.Checked) {
            $arguments += " -WindowStyle Hidden"
        }
        
        $arguments += " -Command `"$commandScript`""
        
        return $arguments
    }
    
    # Preview button click
    $btnPreview.Add_Click({
        if (-not $script:scriptInfo) {
            [System.Windows.Forms.MessageBox]::Show("Please select a PowerShell script first", "No Script Selected", 'OK', 'Warning')
            return
        }
        
        # Validate password requirement for elevated/background tasks
        if (-not $chkRunAsSystem.Checked) {
            if (($chkHighestPrivileges.Checked -or $chkRunLoggedOff.Checked) -and [string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                $warningMsg = "A password is required when using these settings:`n`n"
                if ($chkHighestPrivileges.Checked) {
                    $warningMsg += "- Run with highest privileges`n"
                }
                if ($chkRunLoggedOff.Checked) {
                    $warningMsg += "- Run whether user is logged on or not`n"
                }
                $warningMsg += "`nPlease enter a password for user: $($txtUser.Text)"
                
                [System.Windows.Forms.MessageBox]::Show($warningMsg, "Password Required", 'OK', 'Warning')
                return
            }
        }
        
        # Build PowerShell command based on radio button selection
        if ($radioPS5.Checked) {
            $psExe = Get-PowerShellExecutable -RequiredVersion '5.1' -PowerShellEdition 'Desktop'
        } elseif ($radioPS7.Checked) {
            $psExe = Get-PowerShellExecutable -RequiredVersion '7.0' -PowerShellEdition 'Core'
        } else {
            $psExe = Get-PowerShellExecutable -RequiredVersion $script:scriptInfo.RequiredVersion -PowerShellEdition $script:scriptInfo.PowerShellEdition
        }
        
        # Build arguments using shared function
        $arguments = Build-CommandArguments -IncludeWindowStyle $true
        
        # Get script directory for working directory
        $scriptDirectory = Split-Path -Path $txtScript.Text -Parent
        
        # Build execution context message
        $accountContext = if ($chkRunAsSystem.Checked) {
            "SYSTEM Account (NT AUTHORITY\SYSTEM)
  - No password required
  - Full system privileges
  - Access to system files (C:\Program Files, etc.)
  - DPAPI credentials: C:\ProgramData\MXB\SYSTEM_API_Credentials.Secure.xml"
        } else {
            "User Account: $($txtUser.Text)
  - Password required: $(if ([string]::IsNullOrWhiteSpace($txtPassword.Text)) { 'NOT SET' } else { 'SET' })
  - DPAPI credentials: C:\ProgramData\MXB\{computername}_{username}_API_Credentials.Secure.xml"
        }
        
        # Build preview message
        $previewMsg = @"
TASK SCHEDULER CONFIGURATION PREVIEW
================================================================

Task Name: $($txtTaskName.Text)
Task Path: \N-able\NerdScripts\

Execution Context:
$accountContext

PowerShell Executable:
$psExe

Command Arguments:
$arguments

Working Directory:
$scriptDirectory

Full Command Line:
`"$psExe`" $arguments

================================================================

Schedule: $($cmbSchedule.SelectedItem)
Run Level: $(if ($chkHighestPrivileges.Checked) { 'Highest Privileges' } else { 'Limited User' })
Execution Mode: $(if ($chkRunLoggedOff.Checked) { 'Whether user is logged on or not' } else { 'Only when user is logged on (Interactive)' })
Run Hidden: $(if ($chkRunHidden.Checked) { 'Yes (window will close automatically)' } else { 'No (window will stay open)' })

================================================================
NOTE: Test Run executes as current user, NOT as scheduled task user.
      Actual scheduled task will run with the account shown above.
================================================================
"@
        
        # Create preview form
        $previewForm = New-Object System.Windows.Forms.Form
        $previewForm.Text = 'Scheduled Task Preview'
        $previewForm.Size = New-Object System.Drawing.Size(800, 600)
        $previewForm.StartPosition = 'CenterParent'
        $previewForm.FormBorderStyle = 'Sizable'
        $previewForm.MinimizeBox = $false
        $previewForm.MaximizeBox = $true
        
        $txtPreview = New-Object System.Windows.Forms.TextBox
        $txtPreview.Location = New-Object System.Drawing.Point(10, 10)
        $txtPreview.Size = New-Object System.Drawing.Size(760, 500)
        $txtPreview.Multiline = $true
        $txtPreview.ScrollBars = 'Both'
        $txtPreview.WordWrap = $false
        $txtPreview.Font = New-Object System.Drawing.Font('Consolas', 9)
        $txtPreview.Text = $previewMsg
        $txtPreview.ReadOnly = $true
        $previewForm.Controls.Add($txtPreview)
        
        $btnTest = New-Object System.Windows.Forms.Button
        $btnTest.Location = New-Object System.Drawing.Point(470, 520)
        $btnTest.Size = New-Object System.Drawing.Size(90, 30)
        $btnTest.Text = 'Test Run'
        $btnTest.Add_Click({
            # NOTE: Test Run intentionally ignores "Run hidden" checkbox to allow user to see output/errors.
            # The window will always be visible during testing and will stay open (-NoExit) after completion.
            # The actual scheduled task WILL honor the "Run hidden" setting.
            
            # Warn if SYSTEM is selected (cannot emulate SYSTEM context from Test Run)
            if ($chkRunAsSystem.Checked) {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "WARNING: Test Run cannot execute as SYSTEM account.`n`n" +
                    "Test will run as current user: $env:USERDOMAIN\$env:USERNAME`n`n" +
                    "Actual scheduled task will run as: NT AUTHORITY\SYSTEM`n`n" +
                    "Differences:`n" +
                    "  - File access permissions (SYSTEM has more access)`n" +
                    "  - DPAPI credential file location`n" +
                    "  - Environment variables`n`n" +
                    "Continue with test as current user?",
                    "SYSTEM Emulation Not Possible",
                    'YesNo',
                    'Warning'
                )
                if ($result -eq 'No') {
                    $previewForm.Close()
                    return
                }
            }
            
            # Run test - ALWAYS visible window (ignore Run Hidden checkbox for testing)
            # Build test arguments from base arguments, removing -WindowStyle Hidden if present
            $testArgs = $arguments -replace '-WindowStyle Hidden', ''
            
            # Extract the command portion to append post-execution message
            if ($testArgs -match '-Command "(.+)"$') {
                $commandContent = $matches[1]
                # Append informational message after script execution
                $postMessage = "; Write-Host ''; Write-Host '========================================' -ForegroundColor Cyan; Write-Host 'TEST RUN COMPLETE' -ForegroundColor Cyan; Write-Host '========================================' -ForegroundColor Cyan; Write-Host 'NOTE: Test Run ignores Run Hidden setting to show output.' -ForegroundColor Yellow; Write-Host '      Actual scheduled task will honor Run Hidden setting.' -ForegroundColor Yellow; Write-Host '      Window kept open for review (close manually).' -ForegroundColor Yellow; Write-Host '========================================' -ForegroundColor Cyan"
                $preMessage = "Write-Host '========================================' -ForegroundColor Cyan; Write-Host 'TEST RUN  |  PowerShell ' + `$PSVersionTable.PSVersion.ToString() + '  (' + (Get-Process -Id `$PID).Path + ')' -ForegroundColor Cyan; Write-Host '========================================' -ForegroundColor Cyan; Write-Host ''; "
                $testArgs = $testArgs -replace [regex]::Escape("-Command `"$commandContent`""), "-NoExit -Command `"$preMessage$commandContent$postMessage"
            } else {
                # Fallback: just add -NoExit if pattern doesn't match
                $testArgs = "-NoExit $testArgs"
            }
            
            # Start elevated process if "Run with highest privileges" is checked
            if ($chkHighestPrivileges.Checked) {
                Start-Process $psExe -ArgumentList $testArgs -Verb RunAs -WorkingDirectory $scriptDirectory
            } else {
                Start-Process $psExe -ArgumentList $testArgs -WorkingDirectory $scriptDirectory
            }
        })
        $previewForm.Controls.Add($btnTest)
        
        $btnCopy = New-Object System.Windows.Forms.Button
        $btnCopy.Location = New-Object System.Drawing.Point(570, 520)
        $btnCopy.Size = New-Object System.Drawing.Size(90, 30)
        $btnCopy.Text = 'Copy'
        $btnCopy.Add_Click({
            [System.Windows.Forms.Clipboard]::SetText("`"$psExe`" $arguments")
            [System.Windows.Forms.MessageBox]::Show('Command copied to clipboard!', 'Copied', 'OK', 'Information')
        })
        $previewForm.Controls.Add($btnCopy)
        
        $btnClosePreview = New-Object System.Windows.Forms.Button
        $btnClosePreview.Location = New-Object System.Drawing.Point(670, 520)
        $btnClosePreview.Size = New-Object System.Drawing.Size(90, 30)
        $btnClosePreview.Text = 'Close'
        $btnClosePreview.DialogResult = 'OK'
        $previewForm.Controls.Add($btnClosePreview)
        $previewForm.AcceptButton = $btnClosePreview
        
        $null = $previewForm.ShowDialog()
    })
    
    # Create task button click
    $btnCreate.Add_Click({
        try {
            # Validation
            if ([string]::IsNullOrWhiteSpace($txtTaskName.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a task name", "Validation Error", 'OK', 'Warning')
                return
            }
            
            # Validate password requirement for elevated/background tasks (skip for SYSTEM)
            if (-not $chkRunAsSystem.Checked) {
                if (($chkHighestPrivileges.Checked -or $chkRunLoggedOff.Checked) -and [string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                    $warningMsg = "A password is required when using these settings:`n`n"
                    if ($chkHighestPrivileges.Checked) {
                        $warningMsg += "- Run with highest privileges`n"
                    }
                    if ($chkRunLoggedOff.Checked) {
                        $warningMsg += "- Run whether user is logged on or not`n"
                    }
                    $warningMsg += "`nPlease enter a password for user: $($txtUser.Text)"
                    
                    [System.Windows.Forms.MessageBox]::Show($warningMsg, "Password Required", 'OK', 'Warning')
                    return
                }
            }
            
            # Validate user account and credentials (skip for SYSTEM)
            if (-not $chkRunAsSystem.Checked) {
                if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
                    [System.Windows.Forms.MessageBox]::Show("Please enter a user account", "Validation Error", 'OK', 'Warning')
                    return
                }
                
                if ($chkRunLoggedOff.Checked) {
                    if ([string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                        [System.Windows.Forms.MessageBox]::Show("Password is required when running whether user is logged on or not", "Validation Error", 'OK', 'Warning')
                        return
                    }
                    
                    # Validate credentials before creating task
                    try {
                        $secPass = ConvertTo-SecureString $txtPassword.Text -AsPlainText -Force
                        $testCred = New-Object System.Management.Automation.PSCredential($txtUser.Text, $secPass)
                        # Test credential by attempting to use it (this will fail if invalid)
                        $null = Start-Process cmd.exe -ArgumentList '/c echo test' -Credential $testCred -WindowStyle Hidden -PassThru -ErrorAction Stop
                        Start-Sleep -Milliseconds 100
                    } catch {
                        $result = [System.Windows.Forms.MessageBox]::Show("Unable to validate credentials. The password may be incorrect.`n`nDo you want to continue anyway?", "Credential Warning", 'YesNo', 'Warning')
                        if ($result -eq 'No') {
                            return
                        }
                    }
                }
            }
            
            # Check for mandatory parameters
            foreach ($param in $script:scriptInfo.Parameters) {
                if ($param.Mandatory) {
                    $control = $script:paramControls[$param.Name]
                    if ($control -is [System.Windows.Forms.TextBox] -and [string]::IsNullOrWhiteSpace($control.Text)) {
                        [System.Windows.Forms.MessageBox]::Show("Parameter -$($param.Name) is mandatory", "Validation Error", 'OK', 'Warning')
                        return
                    }
                }
            }
            
            # Build PowerShell command based on radio button selection
            if ($radioPS5.Checked) {
                # Force PowerShell 5.1
                $psExe = Get-PowerShellExecutable -RequiredVersion '5.1' -PowerShellEdition 'Desktop'
            } elseif ($radioPS7.Checked) {
                # Force PowerShell 7+
                $psExe = Get-PowerShellExecutable -RequiredVersion '7.0' -PowerShellEdition 'Core'
            } else {
                # Auto-detect based on script requirements
                $psExe = Get-PowerShellExecutable -RequiredVersion $script:scriptInfo.RequiredVersion -PowerShellEdition $script:scriptInfo.PowerShellEdition
            }
            
            # Get script's root directory for WorkingDirectory
            $scriptDirectory = Split-Path -Path $txtScript.Text -Parent
            
            # Build arguments using shared function
            $arguments = Build-CommandArguments -IncludeWindowStyle $true
            
            # Build scheduled task action
            $action = New-ScheduledTaskAction -Execute $psExe -Argument $arguments -WorkingDirectory $scriptDirectory
            
            # Debug: Show the command that will be executed
            $debugCommand = "`"$psExe`" $arguments"
            Write-Host "`n========================================" -ForegroundColor Cyan
            Write-Host "Command to be executed:" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host $debugCommand -ForegroundColor White
            Write-Host "========================================`n" -ForegroundColor Cyan
            
            # Build trigger based on schedule type
            $trigger = switch ($cmbSchedule.SelectedItem) {
                'Once' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    New-ScheduledTaskTrigger -Once -At $startTime
                }
                'Daily' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    New-ScheduledTaskTrigger -Daily -At $startTime
                }
                'Hourly' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Hours 1)
                    $trigger
                }
                'Every X Minutes' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    $interval = [int]$numInterval.Value
                    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Minutes $interval)
                    $trigger
                }
                'Every X Hours' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    $interval = [int]$numInterval.Value
                    $trigger = New-ScheduledTaskTrigger -Once -At $startTime -RepetitionInterval (New-TimeSpan -Hours $interval)
                    $trigger
                }
                'Every X Days' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    $interval = [int]$numInterval.Value
                    New-ScheduledTaskTrigger -Daily -DaysInterval $interval -At $startTime
                }
                'Weekly' {
                    $startTime = Get-Date -Date $dtpStartDate.Value.Date -Hour $dtpStartTime.Value.Hour -Minute $dtpStartTime.Value.Minute -Second 0
                    $days = @()
                    for ($i = 0; $i -lt $clbWeekdays.Items.Count; $i++) {
                        if ($clbWeekdays.GetItemChecked($i)) {
                            $days += $clbWeekdays.Items[$i]
                        }
                    }
                    if ($days.Count -eq 0) {
                        [System.Windows.Forms.MessageBox]::Show("Please select at least one day of the week", "Validation Error", 'OK', 'Warning')
                        return
                    }
                    New-ScheduledTaskTrigger -Weekly -At $startTime -DaysOfWeek $days
                }
                'At Startup' {
                    New-ScheduledTaskTrigger -AtStartup
                }
                'At Logon' {
                    New-ScheduledTaskTrigger -AtLogon
                }
            }
            
            # Build principal (user context for task execution)
            if ($chkRunAsSystem.Checked) {
                # SYSTEM account - simple principal, always highest privileges
                $principal = New-ScheduledTaskPrincipal -UserId 'NT AUTHORITY\SYSTEM' -RunLevel 'Highest'
            } else {
                # User account - standard principal
                $principalArgs = @{
                    UserId = $txtUser.Text
                    RunLevel = if ($chkHighestPrivileges.Checked) { 'Highest' } else { 'Limited' }
                }
                
                if ($chkRunLoggedOff.Checked) {
                    $principalArgs.LogonType = 'Password'
                } else {
                    $principalArgs.LogonType = 'Interactive'
                }
                
                $principal = New-ScheduledTaskPrincipal @principalArgs
            }
            
            # Build settings
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
            
            # Register task
            $registerArgs = @{
                TaskName = $txtTaskName.Text
                TaskPath = '\N-able\NerdScripts\'
                Action = $action
                Trigger = $trigger
                Principal = $principal
                Settings = $settings
            }
            
            if ($chkRunAsSystem.Checked) {
                # SYSTEM account - Principal contains everything, no password needed
                # Task Scheduler automatically handles SYSTEM account authentication
            } elseif ($chkRunLoggedOff.Checked -and -not [string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                # User account with password authentication
                # REMOVE Principal and use User/Password instead
                $registerArgs.Remove('Principal')
                $registerArgs.User = $txtUser.Text
                $registerArgs.Password = $txtPassword.Text
                # RunLevel must be specified separately when using User/Password
                if ($chkHighestPrivileges.Checked) {
                    $registerArgs.RunLevel = 'Highest'
                }
            }
            # When NOT using password (interactive mode), the Principal object already contains everything we need
            
            try {
                $taskCreated = Register-ScheduledTask @registerArgs -Force -ErrorAction Stop
                
                # For password-based tasks (NOT SYSTEM), re-apply credentials using schtasks for reliability
                if (-not $chkRunAsSystem.Checked -and $chkRunLoggedOff.Checked -and -not [string]::IsNullOrWhiteSpace($txtPassword.Text)) {
                    Start-Sleep -Milliseconds 500
                    $taskFullPath = "\N-able\NerdScripts\$($txtTaskName.Text)"
                    
                    # Use schtasks to update credentials (more reliable than Register-ScheduledTask for passwords)
                    $schtasksArgs = "/Change /TN `"$taskFullPath`" /RU `"$($txtUser.Text)`" /RP `"$($txtPassword.Text)`""
                    $schtasksResult = & schtasks.exe $schtasksArgs.Split(' ') 2>&1
                    
                    if ($LASTEXITCODE -ne 0) {
                        Write-Warning "schtasks credential update returned code: $LASTEXITCODE"
                    }
                }
            } catch {
                $errorMsg = $_.Exception.Message
                $errorCode = $_.Exception.HResult
                
                # Common error codes
                $errorDetails = switch ($errorCode) {
                    -2147023570 { "ERROR: Invalid username or password (0x8007052E)`n`nPlease verify:`n- Username format is correct (DOMAIN\User or Computer\User)`n- Password is correct`n- Account has 'Log on as batch job' rights" }
                    -2147024891 { "ERROR: Access denied`n`nPlease run this script as Administrator" }
                    default { "ERROR: $errorMsg`n`nError Code: $errorCode" }
                }
                
                [System.Windows.Forms.MessageBox]::Show($errorDetails, "Task Registration Failed", 'OK', 'Error')
                return
            }
            
            # Verify task was created successfully
            Start-Sleep -Milliseconds 500
            $verifyTask = Get-ScheduledTask -TaskPath '\N-able\NerdScripts\' -TaskName $txtTaskName.Text -ErrorAction SilentlyContinue
            
            if (-not $verifyTask) {
                [System.Windows.Forms.MessageBox]::Show("Task registration completed but task was not found in Task Scheduler.`n`nPlease check Task Scheduler manually.", "Warning", 'OK', 'Warning')
                return
            }
            
            # Display task details for verification
            $taskDetails = @"
Task Created Successfully!

Task Name: $($txtTaskName.Text)
Task Path: \N-able\NerdScripts\
Run Level: $($verifyTask.Principal.RunLevel)
Logon Type: $($verifyTask.Principal.LogonType)
User: $($verifyTask.Principal.UserId)
State: $($verifyTask.State)
"@
            
            [System.Windows.Forms.MessageBox]::Show($taskDetails, "Task Created Successfully", 'OK', 'Information')
            $form.DialogResult = 'OK'
            $form.Close()
            
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating scheduled task:`n`n$($_.Exception.Message)", "Error", 'OK', 'Error')
        }
    })
    
    #endregion
    
    # Show form
    $result = $form.ShowDialog()
    
    return $result
}

#endregion

#region ----- Main Script ----

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "`n*** WARNING: This script requires Administrator privileges to create scheduled tasks" -ForegroundColor Yellow
    Write-Host "Please run PowerShell as Administrator and try again.`n" -ForegroundColor Yellow
    
    $response = Read-Host "Would you like to restart this script as Administrator? (Y/N)"
    if ($response -eq 'Y' -or $response -eq 'y') {
        # Use the same PowerShell version that launched the script
        $psExe = if ($PSVersionTable.PSEdition -eq 'Core') { 
            (Get-Process -Id $PID).Path 
        } else { 
            'powershell.exe' 
        }
        Start-Process $psExe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    } else {
        exit 1
    }
}

Write-Host "`n=======================================" -ForegroundColor Cyan
Write-Host "Scheduled Task Creator" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host "Opening GUI...`n" -ForegroundColor White

# Launch GUI
$result = New-ScheduledTaskGUI

if ($result -eq 'OK') {
    Write-Host "`nTask created successfully" -ForegroundColor Green
} else {
    Write-Host "`nOperation cancelled" -ForegroundColor Yellow
}

# Exit behavior: Close external windows but preserve VS Code integrated terminals
# Check if running in an external console window (not VS Code integrated terminal)
if ($Host.Name -eq 'ConsoleHost' -and [Environment]::UserInteractive) {
    # External PowerShell window - force close
    [Environment]::Exit(0)
} else {
    # VS Code integrated terminal or ISE - graceful exit
    exit 0
}

#endregion
