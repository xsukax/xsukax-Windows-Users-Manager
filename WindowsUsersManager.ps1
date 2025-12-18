#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    xsukax Windows Users Manager
.DESCRIPTION
    Professional PowerShell GUI for managing local Windows user accounts
.AUTHOR
    xsukax
#>

# ============================================================================
# ERROR HANDLING SETUP
# ============================================================================

$ErrorActionPreference = "Stop"
$script:ErrorLogPath = "$env:TEMP\xsukax_WindowsUsersManager_Error.log"

function Write-ErrorLog {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Add-Content -Path $script:ErrorLogPath -Value $logEntry -ErrorAction SilentlyContinue
}

# Wrap entire script in try-catch
try {

# ============================================================================
# ASSEMBLY LOADING
# ============================================================================

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    $msg = "Failed to load Windows Forms assemblies: $($_.Exception.Message)"
    Write-ErrorLog $msg
    [System.Windows.Forms.MessageBox]::Show($msg, "Fatal Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# ============================================================================
# GLOBAL VARIABLES
# ============================================================================

$script:MainForm = $null
$script:DataGrid = $null
$script:StatusLabel = $null

# ============================================================================
# LOGGING AND UI HELPER FUNCTIONS
# ============================================================================

function Write-StatusMessage {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Error", "Warning")]
        [string]$Type = "Info"
    )
    
    try {
        if ($script:StatusLabel) {
            $timestamp = Get-Date -Format "HH:mm:ss"
            $script:StatusLabel.Text = "[$timestamp] $Message"
            
            switch ($Type) {
                "Success" { $script:StatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen }
                "Error"   { $script:StatusLabel.ForeColor = [System.Drawing.Color]::Red }
                "Warning" { $script:StatusLabel.ForeColor = [System.Drawing.Color]::DarkOrange }
                default   { $script:StatusLabel.ForeColor = [System.Drawing.Color]::Black }
            }
            
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        Write-ErrorLog "Write-StatusMessage failed: $($_.Exception.Message)"
    }
}

function Show-Dialog {
    param(
        [string]$Message,
        [string]$Title,
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    try {
        return [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
    }
    catch {
        Write-ErrorLog "Show-Dialog failed: $($_.Exception.Message)"
        return [System.Windows.Forms.DialogResult]::None
    }
}

# ============================================================================
# DATA RETRIEVAL FUNCTIONS
# ============================================================================

function Get-AllLocalUsers {
    try {
        $users = @(Get-LocalUser -ErrorAction Stop | Select-Object Name, Enabled, Description, FullName,
            LastLogon, PasswordLastSet, PasswordExpires, UserMayChangePassword,
            PasswordRequired, PasswordNeverExpires, AccountExpires, SID, LockedOut)
        return $users
    }
    catch {
        Write-StatusMessage "Error retrieving users: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Get-AllLocalUsers failed: $($_.Exception.Message)"
        return @()
    }
}

function Get-UserGroupMembership {
    param([string]$Username)
    
    try {
        $user = Get-LocalUser -Name $Username -ErrorAction Stop
        $groups = @()
        
        $allGroups = Get-LocalGroup -ErrorAction Stop
        foreach ($group in $allGroups) {
            try {
                $members = Get-LocalGroupMember -Group $group.Name -ErrorAction SilentlyContinue
                if ($members.SID -contains $user.SID) {
                    $groups += $group.Name
                }
            }
            catch {
                # Silently skip groups we can't read
            }
        }
        
        return $groups
    }
    catch {
        Write-ErrorLog "Get-UserGroupMembership failed for '$Username': $($_.Exception.Message)"
        return @()
    }
}

function Test-PasswordStrength {
    param([string]$Password)
    
    $issues = @()
    
    if ($Password.Length -lt 8) {
        $issues += "Must be at least 8 characters"
    }
    if ($Password -notmatch '[A-Z]') {
        $issues += "Must contain uppercase letter"
    }
    if ($Password -notmatch '[a-z]') {
        $issues += "Must contain lowercase letter"
    }
    if ($Password -notmatch '\d') {
        $issues += "Must contain number"
    }
    if ($Password -notmatch '[^a-zA-Z0-9]') {
        $issues += "Must contain special character"
    }
    
    return $issues
}

# ============================================================================
# UI UPDATE FUNCTIONS
# ============================================================================

function Update-UserGrid {
    if (-not $script:DataGrid) { 
        Write-ErrorLog "DataGrid is null in Update-UserGrid"
        return 
    }
    
    try {
        Write-StatusMessage "Refreshing user list..." -Type "Info"
        
        $script:DataGrid.Rows.Clear()
        $users = Get-AllLocalUsers
        
        if ($users.Count -eq 0) {
            Write-StatusMessage "No users found or unable to retrieve users" -Type "Warning"
            return
        }
        
        foreach ($user in $users) {
            try {
                $groups = Get-UserGroupMembership -Username $user.Name
                $isAdmin = $groups -contains "Administrators"
                $groupList = $groups -join ", "
                
                $lastLogon = if ($user.LastLogon) { 
                    $user.LastLogon.ToString("yyyy-MM-dd HH:mm") 
                } else { 
                    "Never" 
                }
                
                $passwordSet = if ($user.PasswordLastSet) { 
                    $user.PasswordLastSet.ToString("yyyy-MM-dd") 
                } else { 
                    "N/A" 
                }
                
                $rowIndex = $script:DataGrid.Rows.Add(
                    $user.Name,
                    $user.Enabled,
                    $isAdmin,
                    $groupList,
                    $user.Description,
                    $lastLogon,
                    $passwordSet
                )
                
                # Store complete user object in row tag
                $script:DataGrid.Rows[$rowIndex].Tag = $user
            }
            catch {
                Write-ErrorLog "Error adding user $($user.Name) to grid: $($_.Exception.Message)"
            }
        }
        
        Write-StatusMessage "Loaded $($users.Count) users" -Type "Success"
    }
    catch {
        Write-StatusMessage "Error refreshing grid: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Update-UserGrid failed: $($_.Exception.Message)"
    }
}

# ============================================================================
# USER MANAGEMENT FUNCTIONS
# ============================================================================

function Show-AddUserDialog {
    try {
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = "Add New User"
        $dialog.Size = New-Object System.Drawing.Size(470, 480)
        $dialog.StartPosition = "CenterParent"
        $dialog.FormBorderStyle = "FixedDialog"
        $dialog.MaximizeBox = $false
        
        $y = 20
        
        # Username
        $lblUser = New-Object System.Windows.Forms.Label
        $lblUser.Location = New-Object System.Drawing.Point(20, $y)
        $lblUser.Size = New-Object System.Drawing.Size(130, 20)
        $lblUser.Text = "Username:"
        $dialog.Controls.Add($lblUser)
        
        $txtUser = New-Object System.Windows.Forms.TextBox
        $txtUser.Location = New-Object System.Drawing.Point(160, $y)
        $txtUser.Size = New-Object System.Drawing.Size(280, 20)
        $dialog.Controls.Add($txtUser)
        $y += 35
        
        # Full Name
        $lblFull = New-Object System.Windows.Forms.Label
        $lblFull.Location = New-Object System.Drawing.Point(20, $y)
        $lblFull.Size = New-Object System.Drawing.Size(130, 20)
        $lblFull.Text = "Full Name:"
        $dialog.Controls.Add($lblFull)
        
        $txtFull = New-Object System.Windows.Forms.TextBox
        $txtFull.Location = New-Object System.Drawing.Point(160, $y)
        $txtFull.Size = New-Object System.Drawing.Size(280, 20)
        $dialog.Controls.Add($txtFull)
        $y += 35
        
        # Description
        $lblDesc = New-Object System.Windows.Forms.Label
        $lblDesc.Location = New-Object System.Drawing.Point(20, $y)
        $lblDesc.Size = New-Object System.Drawing.Size(130, 20)
        $lblDesc.Text = "Description:"
        $dialog.Controls.Add($lblDesc)
        
        $txtDesc = New-Object System.Windows.Forms.TextBox
        $txtDesc.Location = New-Object System.Drawing.Point(160, $y)
        $txtDesc.Size = New-Object System.Drawing.Size(280, 20)
        $dialog.Controls.Add($txtDesc)
        $y += 35
        
        # Password
        $lblPass = New-Object System.Windows.Forms.Label
        $lblPass.Location = New-Object System.Drawing.Point(20, $y)
        $lblPass.Size = New-Object System.Drawing.Size(130, 20)
        $lblPass.Text = "Password:"
        $dialog.Controls.Add($lblPass)
        
        $txtPass = New-Object System.Windows.Forms.TextBox
        $txtPass.Location = New-Object System.Drawing.Point(160, $y)
        $txtPass.Size = New-Object System.Drawing.Size(280, 20)
        $txtPass.PasswordChar = '*'
        $dialog.Controls.Add($txtPass)
        $y += 35
        
        # Confirm Password
        $lblConfirm = New-Object System.Windows.Forms.Label
        $lblConfirm.Location = New-Object System.Drawing.Point(20, $y)
        $lblConfirm.Size = New-Object System.Drawing.Size(130, 20)
        $lblConfirm.Text = "Confirm Password:"
        $dialog.Controls.Add($lblConfirm)
        
        $txtConfirm = New-Object System.Windows.Forms.TextBox
        $txtConfirm.Location = New-Object System.Drawing.Point(160, $y)
        $txtConfirm.Size = New-Object System.Drawing.Size(280, 20)
        $txtConfirm.PasswordChar = '*'
        $dialog.Controls.Add($txtConfirm)
        $y += 45
        
        # Options
        $chkMustChange = New-Object System.Windows.Forms.CheckBox
        $chkMustChange.Location = New-Object System.Drawing.Point(20, $y)
        $chkMustChange.Size = New-Object System.Drawing.Size(400, 20)
        $chkMustChange.Text = "User must change password at next logon"
        $dialog.Controls.Add($chkMustChange)
        $y += 30
        
        $chkNeverExpires = New-Object System.Windows.Forms.CheckBox
        $chkNeverExpires.Location = New-Object System.Drawing.Point(20, $y)
        $chkNeverExpires.Size = New-Object System.Drawing.Size(400, 20)
        $chkNeverExpires.Text = "Password never expires"
        $dialog.Controls.Add($chkNeverExpires)
        $y += 30
        
        $chkDisabled = New-Object System.Windows.Forms.CheckBox
        $chkDisabled.Location = New-Object System.Drawing.Point(20, $y)
        $chkDisabled.Size = New-Object System.Drawing.Size(400, 20)
        $chkDisabled.Text = "Account is disabled"
        $dialog.Controls.Add($chkDisabled)
        $y += 30
        
        $chkAdmin = New-Object System.Windows.Forms.CheckBox
        $chkAdmin.Location = New-Object System.Drawing.Point(20, $y)
        $chkAdmin.Size = New-Object System.Drawing.Size(400, 20)
        $chkAdmin.Text = "Add to Administrators group"
        $chkAdmin.ForeColor = [System.Drawing.Color]::Red
        $dialog.Controls.Add($chkAdmin)
        $y += 50
        
        # Mutual exclusion
        $chkMustChange.Add_CheckedChanged({
            if ($chkMustChange.Checked) { $chkNeverExpires.Checked = $false }
        })
        $chkNeverExpires.Add_CheckedChanged({
            if ($chkNeverExpires.Checked) { $chkMustChange.Checked = $false }
        })
        
        # Buttons
        $btnCreate = New-Object System.Windows.Forms.Button
        $btnCreate.Location = New-Object System.Drawing.Point(240, $y)
        $btnCreate.Size = New-Object System.Drawing.Size(90, 30)
        $btnCreate.Text = "Create"
        $btnCreate.Add_Click({
            try {
                # Validate
                if ([string]::IsNullOrWhiteSpace($txtUser.Text)) {
                    Show-Dialog "Username is required" "Validation" -Icon Warning
                    return
                }
                
                if ([string]::IsNullOrWhiteSpace($txtPass.Text)) {
                    Show-Dialog "Password is required" "Validation" -Icon Warning
                    return
                }
                
                if ($txtPass.Text -ne $txtConfirm.Text) {
                    Show-Dialog "Passwords do not match" "Validation" -Icon Warning
                    return
                }
                
                $issues = Test-PasswordStrength -Password $txtPass.Text
                if ($issues.Count -gt 0) {
                    Show-Dialog ("Password requirements:`n`n" + ($issues -join "`n")) "Password Strength" -Icon Warning
                    return
                }
                
                $secPass = ConvertTo-SecureString $txtPass.Text -AsPlainText -Force
                
                $params = @{
                    Name = $txtUser.Text
                    Password = $secPass
                    PasswordNeverExpires = $chkNeverExpires.Checked
                }
                
                if ($txtFull.Text) { $params['FullName'] = $txtFull.Text }
                if ($txtDesc.Text) { $params['Description'] = $txtDesc.Text }
                
                New-LocalUser @params -ErrorAction Stop | Out-Null
                
                if ($chkMustChange.Checked) {
                    & net user $txtUser.Text /logonpasswordchg:yes 2>&1 | Out-Null
                }
                
                if ($chkDisabled.Checked) {
                    Disable-LocalUser -Name $txtUser.Text -ErrorAction Stop
                }
                
                if ($chkAdmin.Checked) {
                    Add-LocalGroupMember -Group "Administrators" -Member $txtUser.Text -ErrorAction Stop
                }
                
                Write-StatusMessage "User '$($txtUser.Text)' created successfully" -Type "Success"
                Show-Dialog "User created successfully" "Success" -Icon Information
                
                $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $dialog.Close()
            }
            catch {
                Write-StatusMessage "Error creating user: $($_.Exception.Message)" -Type "Error"
                Write-ErrorLog "Create user failed: $($_.Exception.Message)"
                Show-Dialog "Failed to create user:`n`n$($_.Exception.Message)" "Error" -Icon Error
            }
        })
        $dialog.Controls.Add($btnCreate)
        $dialog.AcceptButton = $btnCreate
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(340, $y)
        $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Controls.Add($btnCancel)
        $dialog.CancelButton = $btnCancel
        
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Update-UserGrid
        }
    }
    catch {
        Write-ErrorLog "Show-AddUserDialog failed: $($_.Exception.Message)"
        Show-Dialog "Error opening Add User dialog: $($_.Exception.Message)" "Error" -Icon Error
    }
}

function Show-UserPropertiesDialog {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) {
            Show-Dialog "Please select a user first" "No Selection" -Icon Information
            return
        }
        
        $selectedRow = $script:DataGrid.SelectedRows[0]
        $username = $selectedRow.Cells[0].Value
        $userObj = $selectedRow.Tag
        
        if (-not $userObj) {
            Show-Dialog "Unable to load user data" "Error" -Icon Error
            return
        }
        
        # Create Properties Dialog
        $propDlg = New-Object System.Windows.Forms.Form
        $propDlg.Text = "User Properties - $username"
        $propDlg.Size = New-Object System.Drawing.Size(620, 580)
        $propDlg.StartPosition = "CenterParent"
        $propDlg.FormBorderStyle = "FixedDialog"
        $propDlg.MaximizeBox = $false
        
        # Create TabControl
        $tabs = New-Object System.Windows.Forms.TabControl
        $tabs.Location = New-Object System.Drawing.Point(10, 10)
        $tabs.Size = New-Object System.Drawing.Size(585, 470)
        $propDlg.Controls.Add($tabs)
        
        # ========================================
        # GENERAL TAB
        # ========================================
        $tabGeneral = New-Object System.Windows.Forms.TabPage
        $tabGeneral.Text = "General"
        $tabs.TabPages.Add($tabGeneral)
        
        $y = 20
        
        # Username (readonly)
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Username:"
        $tabGeneral.Controls.Add($lbl)
        
        $txtUsername = New-Object System.Windows.Forms.TextBox
        $txtUsername.Location = New-Object System.Drawing.Point(150, $y)
        $txtUsername.Size = New-Object System.Drawing.Size(400, 20)
        $txtUsername.Text = $username
        $txtUsername.ReadOnly = $true
        $txtUsername.BackColor = [System.Drawing.Color]::WhiteSmoke
        $tabGeneral.Controls.Add($txtUsername)
        $y += 35
        
        # Full Name
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Full Name:"
        $tabGeneral.Controls.Add($lbl)
        
        $txtFullName = New-Object System.Windows.Forms.TextBox
        $txtFullName.Location = New-Object System.Drawing.Point(150, $y)
        $txtFullName.Size = New-Object System.Drawing.Size(400, 20)
        $txtFullName.Text = $userObj.FullName
        $tabGeneral.Controls.Add($txtFullName)
        $y += 35
        
        # Description
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "Description:"
        $tabGeneral.Controls.Add($lbl)
        
        $txtDescription = New-Object System.Windows.Forms.TextBox
        $txtDescription.Location = New-Object System.Drawing.Point(150, $y)
        $txtDescription.Size = New-Object System.Drawing.Size(400, 60)
        $txtDescription.Multiline = $true
        $txtDescription.ScrollBars = "Vertical"
        $txtDescription.Text = $userObj.Description
        $tabGeneral.Controls.Add($txtDescription)
        $y += 75
        
        # Account Status
        $grp = New-Object System.Windows.Forms.GroupBox
        $grp.Location = New-Object System.Drawing.Point(20, $y)
        $grp.Size = New-Object System.Drawing.Size(530, 90)
        $grp.Text = "Account Status"
        $tabGeneral.Controls.Add($grp)
        
        $chkEnabled = New-Object System.Windows.Forms.CheckBox
        $chkEnabled.Location = New-Object System.Drawing.Point(15, 25)
        $chkEnabled.Size = New-Object System.Drawing.Size(200, 20)
        $chkEnabled.Text = "Account is enabled"
        $chkEnabled.Checked = $userObj.Enabled
        $grp.Controls.Add($chkEnabled)
        
        $chkLocked = New-Object System.Windows.Forms.CheckBox
        $chkLocked.Location = New-Object System.Drawing.Point(15, 50)
        $chkLocked.Size = New-Object System.Drawing.Size(200, 20)
        $chkLocked.Text = "Account is locked"
        $chkLocked.Checked = $userObj.LockedOut
        $chkLocked.Enabled = $false
        $grp.Controls.Add($chkLocked)
        $y += 105
        
        # SID
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(120, 20)
        $lbl.Text = "SID:"
        $tabGeneral.Controls.Add($lbl)
        
        $txtSID = New-Object System.Windows.Forms.TextBox
        $txtSID.Location = New-Object System.Drawing.Point(150, $y)
        $txtSID.Size = New-Object System.Drawing.Size(400, 20)
        $txtSID.Text = $userObj.SID
        $txtSID.ReadOnly = $true
        $txtSID.BackColor = [System.Drawing.Color]::WhiteSmoke
        $tabGeneral.Controls.Add($txtSID)
        
        # ========================================
        # ACCOUNT TAB
        # ========================================
        $tabAccount = New-Object System.Windows.Forms.TabPage
        $tabAccount.Text = "Account"
        $tabs.TabPages.Add($tabAccount)
        
        $y = 20
        
        # Password Options
        $grp = New-Object System.Windows.Forms.GroupBox
        $grp.Location = New-Object System.Drawing.Point(20, $y)
        $grp.Size = New-Object System.Drawing.Size(530, 140)
        $grp.Text = "Password Options"
        $tabAccount.Controls.Add($grp)
        
        $chkPassRequired = New-Object System.Windows.Forms.CheckBox
        $chkPassRequired.Location = New-Object System.Drawing.Point(15, 25)
        $chkPassRequired.Size = New-Object System.Drawing.Size(500, 20)
        $chkPassRequired.Text = "Password is required"
        $chkPassRequired.Checked = $userObj.PasswordRequired
        $grp.Controls.Add($chkPassRequired)
        
        $chkUserCanChange = New-Object System.Windows.Forms.CheckBox
        $chkUserCanChange.Location = New-Object System.Drawing.Point(15, 50)
        $chkUserCanChange.Size = New-Object System.Drawing.Size(500, 20)
        $chkUserCanChange.Text = "User can change password"
        $chkUserCanChange.Checked = $userObj.UserMayChangePassword
        $grp.Controls.Add($chkUserCanChange)
        
        $chkPassNeverExpires = New-Object System.Windows.Forms.CheckBox
        $chkPassNeverExpires.Location = New-Object System.Drawing.Point(15, 75)
        $chkPassNeverExpires.Size = New-Object System.Drawing.Size(500, 20)
        $chkPassNeverExpires.Text = "Password never expires"
        $chkPassNeverExpires.Checked = $userObj.PasswordNeverExpires
        $grp.Controls.Add($chkPassNeverExpires)
        
        $lblPassAge = New-Object System.Windows.Forms.Label
        $lblPassAge.Location = New-Object System.Drawing.Point(15, 105)
        $lblPassAge.Size = New-Object System.Drawing.Size(500, 25)
        if ($userObj.PasswordLastSet) {
            $days = [math]::Round(((Get-Date) - $userObj.PasswordLastSet).TotalDays)
            $lblPassAge.Text = "Password last set: $($userObj.PasswordLastSet.ToString('yyyy-MM-dd HH:mm')) ($days days ago)"
        } else {
            $lblPassAge.Text = "Password last set: Never"
        }
        $grp.Controls.Add($lblPassAge)
        $y += 155
        
        # Account Expiration
        $grp = New-Object System.Windows.Forms.GroupBox
        $grp.Location = New-Object System.Drawing.Point(20, $y)
        $grp.Size = New-Object System.Drawing.Size(530, 100)
        $grp.Text = "Account Expiration"
        $tabAccount.Controls.Add($grp)
        
        $chkAcctNeverExpires = New-Object System.Windows.Forms.CheckBox
        $chkAcctNeverExpires.Location = New-Object System.Drawing.Point(15, 25)
        $chkAcctNeverExpires.Size = New-Object System.Drawing.Size(500, 20)
        $chkAcctNeverExpires.Text = "Account never expires"
        $chkAcctNeverExpires.Checked = ($userObj.AccountExpires -eq $null)
        $grp.Controls.Add($chkAcctNeverExpires)
        
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(15, 53)
        $lbl.Size = New-Object System.Drawing.Size(100, 20)
        $lbl.Text = "Expires on:"
        $grp.Controls.Add($lbl)
        
        $dtpExpires = New-Object System.Windows.Forms.DateTimePicker
        $dtpExpires.Location = New-Object System.Drawing.Point(120, 50)
        $dtpExpires.Size = New-Object System.Drawing.Size(200, 20)
        if ($userObj.AccountExpires) {
            $dtpExpires.Value = $userObj.AccountExpires
        }
        $dtpExpires.Enabled = -not $chkAcctNeverExpires.Checked
        $grp.Controls.Add($dtpExpires)
        
        $chkAcctNeverExpires.Add_CheckedChanged({
            $dtpExpires.Enabled = -not $chkAcctNeverExpires.Checked
        })
        $y += 115
        
        # Logon Info
        $grp = New-Object System.Windows.Forms.GroupBox
        $grp.Location = New-Object System.Drawing.Point(20, $y)
        $grp.Size = New-Object System.Drawing.Size(530, 70)
        $grp.Text = "Logon Information"
        $tabAccount.Controls.Add($grp)
        
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(15, 25)
        $lbl.Size = New-Object System.Drawing.Size(500, 35)
        if ($userObj.LastLogon) {
            $lbl.Text = "Last logon: $($userObj.LastLogon.ToString('yyyy-MM-dd HH:mm:ss'))"
        } else {
            $lbl.Text = "Last logon: Never"
        }
        $grp.Controls.Add($lbl)
        
        # ========================================
        # MEMBER OF TAB
        # ========================================
        $tabGroups = New-Object System.Windows.Forms.TabPage
        $tabGroups.Text = "Member Of"
        $tabs.TabPages.Add($tabGroups)
        
        $lstGroups = New-Object System.Windows.Forms.ListBox
        $lstGroups.Location = New-Object System.Drawing.Point(20, 20)
        $lstGroups.Size = New-Object System.Drawing.Size(530, 400)
        $tabGroups.Controls.Add($lstGroups)
        
        $groups = Get-UserGroupMembership -Username $username
        foreach ($grp in $groups) {
            $lstGroups.Items.Add($grp) | Out-Null
        }
        
        # ========================================
        # BUTTONS
        # ========================================
        $btnApply = New-Object System.Windows.Forms.Button
        $btnApply.Location = New-Object System.Drawing.Point(300, 495)
        $btnApply.Size = New-Object System.Drawing.Size(90, 30)
        $btnApply.Text = "Apply"
        $btnApply.Add_Click({
            try {
                # Update Full Name
                if ($txtFullName.Text -ne $userObj.FullName) {
                    Set-LocalUser -Name $username -FullName $txtFullName.Text -ErrorAction Stop
                }
                
                # Update Description
                if ($txtDescription.Text -ne $userObj.Description) {
                    Set-LocalUser -Name $username -Description $txtDescription.Text -ErrorAction Stop
                }
                
                # Update Enabled
                if ($chkEnabled.Checked -ne $userObj.Enabled) {
                    if ($chkEnabled.Checked) {
                        Enable-LocalUser -Name $username -ErrorAction Stop
                    } else {
                        Disable-LocalUser -Name $username -ErrorAction Stop
                    }
                }
                
                # Update Password Options
                Set-LocalUser -Name $username `
                    -PasswordNeverExpires $chkPassNeverExpires.Checked `
                    -UserMayChangePassword $chkUserCanChange.Checked `
                    -ErrorAction Stop
                
                # Update Account Expiration
                if ($chkAcctNeverExpires.Checked) {
                    Set-LocalUser -Name $username -AccountNeverExpires -ErrorAction Stop
                } else {
                    Set-LocalUser -Name $username -AccountExpires $dtpExpires.Value -ErrorAction Stop
                }
                
                Write-StatusMessage "Properties updated for '$username'" -Type "Success"
                Show-Dialog "Properties updated successfully" "Success" -Icon Information
                
                Update-UserGrid
            }
            catch {
                Write-StatusMessage "Error updating properties: $($_.Exception.Message)" -Type "Error"
                Write-ErrorLog "Update properties failed: $($_.Exception.Message)"
                Show-Dialog "Failed to update properties:`n`n$($_.Exception.Message)" "Error" -Icon Error
            }
        })
        $propDlg.Controls.Add($btnApply)
        
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Location = New-Object System.Drawing.Point(400, 495)
        $btnOK.Size = New-Object System.Drawing.Size(90, 30)
        $btnOK.Text = "OK"
        $btnOK.Add_Click({
            $btnApply.PerformClick()
            $propDlg.Close()
        })
        $propDlg.Controls.Add($btnOK)
        $propDlg.AcceptButton = $btnOK
        
        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Location = New-Object System.Drawing.Point(500, 495)
        $btnClose.Size = New-Object System.Drawing.Size(90, 30)
        $btnClose.Text = "Cancel"
        $btnClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $propDlg.Controls.Add($btnClose)
        $propDlg.CancelButton = $btnClose
        
        $propDlg.ShowDialog() | Out-Null
    }
    catch {
        Write-ErrorLog "Show-UserPropertiesDialog failed: $($_.Exception.Message)"
        Show-Dialog "Error opening Properties: $($_.Exception.Message)" "Error" -Icon Error
    }
}

function Remove-SelectedUser {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) { return }
        
        $username = $script:DataGrid.SelectedRows[0].Cells[0].Value
        
        $result = Show-Dialog "Delete user '$username'?`n`nThis cannot be undone." "Confirm Delete" -Buttons YesNo -Icon Warning
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Remove-LocalUser -Name $username -ErrorAction Stop
            Write-StatusMessage "User '$username' deleted" -Type "Success"
            Update-UserGrid
        }
    }
    catch {
        Write-StatusMessage "Error deleting user: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Remove user failed: $($_.Exception.Message)"
        Show-Dialog "Failed to delete user:`n`n$($_.Exception.Message)" "Error" -Icon Error
    }
}

function Enable-SelectedUser {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) { return }
        
        $username = $script:DataGrid.SelectedRows[0].Cells[0].Value
        
        Enable-LocalUser -Name $username -ErrorAction Stop
        Write-StatusMessage "User '$username' enabled" -Type "Success"
        Update-UserGrid
    }
    catch {
        Write-StatusMessage "Error: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Enable user failed: $($_.Exception.Message)"
        Show-Dialog "Failed to enable user:`n`n$($_.Exception.Message)" "Error" -Icon Error
    }
}

function Disable-SelectedUser {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) { return }
        
        $username = $script:DataGrid.SelectedRows[0].Cells[0].Value
        
        $result = Show-Dialog "Disable user '$username'?" "Confirm" -Buttons YesNo -Icon Warning
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Disable-LocalUser -Name $username -ErrorAction Stop
            Write-StatusMessage "User '$username' disabled" -Type "Success"
            Update-UserGrid
        }
    }
    catch {
        Write-StatusMessage "Error: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Disable user failed: $($_.Exception.Message)"
        Show-Dialog "Failed to disable user:`n`n$($_.Exception.Message)" "Error" -Icon Error
    }
}

function Show-ChangePasswordDialog {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) { return }
        
        $username = $script:DataGrid.SelectedRows[0].Cells[0].Value
        
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = "Change Password - $username"
        $dialog.Size = New-Object System.Drawing.Size(420, 270)
        $dialog.StartPosition = "CenterParent"
        $dialog.FormBorderStyle = "FixedDialog"
        $dialog.MaximizeBox = $false
        
        $y = 20
        
        # New Password
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(130, 20)
        $lbl.Text = "New Password:"
        $dialog.Controls.Add($lbl)
        
        $txtNew = New-Object System.Windows.Forms.TextBox
        $txtNew.Location = New-Object System.Drawing.Point(160, $y)
        $txtNew.Size = New-Object System.Drawing.Size(230, 20)
        $txtNew.PasswordChar = '*'
        $dialog.Controls.Add($txtNew)
        $y += 35
        
        # Confirm
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, $y)
        $lbl.Size = New-Object System.Drawing.Size(130, 20)
        $lbl.Text = "Confirm Password:"
        $dialog.Controls.Add($lbl)
        
        $txtConfirm = New-Object System.Windows.Forms.TextBox
        $txtConfirm.Location = New-Object System.Drawing.Point(160, $y)
        $txtConfirm.Size = New-Object System.Drawing.Size(230, 20)
        $txtConfirm.PasswordChar = '*'
        $dialog.Controls.Add($txtConfirm)
        $y += 45
        
        # Options
        $chkMustChange = New-Object System.Windows.Forms.CheckBox
        $chkMustChange.Location = New-Object System.Drawing.Point(20, $y)
        $chkMustChange.Size = New-Object System.Drawing.Size(370, 20)
        $chkMustChange.Text = "User must change password at next logon"
        $dialog.Controls.Add($chkMustChange)
        $y += 30
        
        $chkNeverExpires = New-Object System.Windows.Forms.CheckBox
        $chkNeverExpires.Location = New-Object System.Drawing.Point(20, $y)
        $chkNeverExpires.Size = New-Object System.Drawing.Size(370, 20)
        $chkNeverExpires.Text = "Password never expires"
        $dialog.Controls.Add($chkNeverExpires)
        $y += 45
        
        # Mutual exclusion
        $chkMustChange.Add_CheckedChanged({
            if ($chkMustChange.Checked) { $chkNeverExpires.Checked = $false }
        })
        $chkNeverExpires.Add_CheckedChanged({
            if ($chkNeverExpires.Checked) { $chkMustChange.Checked = $false }
        })
        
        # Buttons
        $btnChange = New-Object System.Windows.Forms.Button
        $btnChange.Location = New-Object System.Drawing.Point(200, $y)
        $btnChange.Size = New-Object System.Drawing.Size(90, 30)
        $btnChange.Text = "Change"
        $btnChange.Add_Click({
            try {
                if ([string]::IsNullOrWhiteSpace($txtNew.Text)) {
                    Show-Dialog "Password is required" "Validation" -Icon Warning
                    return
                }
                
                if ($txtNew.Text -ne $txtConfirm.Text) {
                    Show-Dialog "Passwords do not match" "Validation" -Icon Warning
                    return
                }
                
                $issues = Test-PasswordStrength -Password $txtNew.Text
                if ($issues.Count -gt 0) {
                    Show-Dialog ("Password requirements:`n`n" + ($issues -join "`n")) "Password Strength" -Icon Warning
                    return
                }
                
                $secPass = ConvertTo-SecureString $txtNew.Text -AsPlainText -Force
                Set-LocalUser -Name $username -Password $secPass -ErrorAction Stop
                
                if ($chkNeverExpires.Checked) {
                    Set-LocalUser -Name $username -PasswordNeverExpires $true -ErrorAction Stop
                } else {
                    Set-LocalUser -Name $username -PasswordNeverExpires $false -ErrorAction Stop
                }
                
                if ($chkMustChange.Checked) {
                    & net user $username /logonpasswordchg:yes 2>&1 | Out-Null
                }
                
                Write-StatusMessage "Password changed for '$username'" -Type "Success"
                Show-Dialog "Password changed successfully" "Success" -Icon Information
                
                $dialog.Close()
            }
            catch {
                Write-StatusMessage "Error: $($_.Exception.Message)" -Type "Error"
                Write-ErrorLog "Change password failed: $($_.Exception.Message)"
                Show-Dialog "Failed to change password:`n`n$($_.Exception.Message)" "Error" -Icon Error
            }
        })
        $dialog.Controls.Add($btnChange)
        $dialog.AcceptButton = $btnChange
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(300, $y)
        $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dialog.Controls.Add($btnCancel)
        $dialog.CancelButton = $btnCancel
        
        $dialog.ShowDialog() | Out-Null
    }
    catch {
        Write-ErrorLog "Show-ChangePasswordDialog failed: $($_.Exception.Message)"
        Show-Dialog "Error opening dialog: $($_.Exception.Message)" "Error" -Icon Error
    }
}

function Show-ManageGroupsDialog {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) { return }
        
        $username = $script:DataGrid.SelectedRows[0].Cells[0].Value
        
        $dialog = New-Object System.Windows.Forms.Form
        $dialog.Text = "Manage Groups - $username"
        $dialog.Size = New-Object System.Drawing.Size(520, 470)
        $dialog.StartPosition = "CenterParent"
        $dialog.FormBorderStyle = "FixedDialog"
        $dialog.MaximizeBox = $false
        
        # Available Groups
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(20, 20)
        $lbl.Size = New-Object System.Drawing.Size(180, 20)
        $lbl.Text = "Available Groups:"
        $dialog.Controls.Add($lbl)
        
        $lstAvailable = New-Object System.Windows.Forms.ListBox
        $lstAvailable.Location = New-Object System.Drawing.Point(20, 45)
        $lstAvailable.Size = New-Object System.Drawing.Size(180, 320)
        $lstAvailable.SelectionMode = "MultiExtended"
        $dialog.Controls.Add($lstAvailable)
        
        # Member Of
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Location = New-Object System.Drawing.Point(310, 20)
        $lbl.Size = New-Object System.Drawing.Size(180, 20)
        $lbl.Text = "Member Of:"
        $dialog.Controls.Add($lbl)
        
        $lstMember = New-Object System.Windows.Forms.ListBox
        $lstMember.Location = New-Object System.Drawing.Point(310, 45)
        $lstMember.Size = New-Object System.Drawing.Size(180, 320)
        $lstMember.SelectionMode = "MultiExtended"
        $dialog.Controls.Add($lstMember)
        
        # Add Button
        $btnAdd = New-Object System.Windows.Forms.Button
        $btnAdd.Location = New-Object System.Drawing.Point(220, 140)
        $btnAdd.Size = New-Object System.Drawing.Size(70, 30)
        $btnAdd.Text = "Add >>"
        $btnAdd.Add_Click({
            try {
                foreach ($item in $lstAvailable.SelectedItems) {
                    try {
                        Add-LocalGroupMember -Group $item -Member $username -ErrorAction Stop
                    }
                    catch {
                        # Ignore if already member
                    }
                }
                
                # Refresh
                $userGroups = Get-UserGroupMembership -Username $username
                $lstMember.Items.Clear()
                $lstAvailable.Items.Clear()
                
                $allGroups = Get-LocalGroup -ErrorAction Stop
                foreach ($grp in $allGroups) {
                    if ($userGroups -contains $grp.Name) {
                        $lstMember.Items.Add($grp.Name) | Out-Null
                    } else {
                        $lstAvailable.Items.Add($grp.Name) | Out-Null
                    }
                }
                
                Write-StatusMessage "Groups updated for '$username'" -Type "Success"
            }
            catch {
                Write-ErrorLog "Add to group failed: $($_.Exception.Message)"
            }
        })
        $dialog.Controls.Add($btnAdd)
        
        # Remove Button
        $btnRemove = New-Object System.Windows.Forms.Button
        $btnRemove.Location = New-Object System.Drawing.Point(220, 180)
        $btnRemove.Size = New-Object System.Drawing.Size(70, 30)
        $btnRemove.Text = "<< Remove"
        $btnRemove.Add_Click({
            try {
                foreach ($item in $lstMember.SelectedItems) {
                    try {
                        Remove-LocalGroupMember -Group $item -Member $username -ErrorAction Stop
                    }
                    catch {
                        # Ignore if not member
                    }
                }
                
                # Refresh
                $userGroups = Get-UserGroupMembership -Username $username
                $lstMember.Items.Clear()
                $lstAvailable.Items.Clear()
                
                $allGroups = Get-LocalGroup -ErrorAction Stop
                foreach ($grp in $allGroups) {
                    if ($userGroups -contains $grp.Name) {
                        $lstMember.Items.Add($grp.Name) | Out-Null
                    } else {
                        $lstAvailable.Items.Add($grp.Name) | Out-Null
                    }
                }
                
                Write-StatusMessage "Groups updated for '$username'" -Type "Success"
            }
            catch {
                Write-ErrorLog "Remove from group failed: $($_.Exception.Message)"
            }
        })
        $dialog.Controls.Add($btnRemove)
        
        # Populate lists
        $userGroups = Get-UserGroupMembership -Username $username
        $allGroups = Get-LocalGroup -ErrorAction Stop
        
        foreach ($grp in $allGroups) {
            if ($userGroups -contains $grp.Name) {
                $lstMember.Items.Add($grp.Name) | Out-Null
            } else {
                $lstAvailable.Items.Add($grp.Name) | Out-Null
            }
        }
        
        # Close Button
        $btnClose = New-Object System.Windows.Forms.Button
        $btnClose.Location = New-Object System.Drawing.Point(400, 385)
        $btnClose.Size = New-Object System.Drawing.Size(90, 30)
        $btnClose.Text = "Close"
        $btnClose.Add_Click({
            $dialog.Close()
            Update-UserGrid
        })
        $dialog.Controls.Add($btnClose)
        $dialog.AcceptButton = $btnClose
        
        $dialog.ShowDialog() | Out-Null
    }
    catch {
        Write-ErrorLog "Show-ManageGroupsDialog failed: $($_.Exception.Message)"
        Show-Dialog "Error opening dialog: $($_.Exception.Message)" "Error" -Icon Error
    }
}

function Toggle-AdminPrivilege {
    try {
        if (-not $script:DataGrid -or $script:DataGrid.SelectedRows.Count -eq 0) { return }
        
        $username = $script:DataGrid.SelectedRows[0].Cells[0].Value
        $isAdmin = $script:DataGrid.SelectedRows[0].Cells[2].Value
        
        if ($isAdmin) {
            $result = Show-Dialog "Remove administrator privileges from '$username'?" "Confirm" -Buttons YesNo -Icon Warning
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Remove-LocalGroupMember -Group "Administrators" -Member $username -ErrorAction Stop
                Write-StatusMessage "Removed '$username' from Administrators" -Type "Success"
            }
        } else {
            $result = Show-Dialog "Grant administrator privileges to '$username'?`n`nThis gives full system control." "Confirm" -Buttons YesNo -Icon Warning
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Add-LocalGroupMember -Group "Administrators" -Member $username -ErrorAction Stop
                Write-StatusMessage "Added '$username' to Administrators" -Type "Success"
            }
        }
        
        Update-UserGrid
    }
    catch {
        Write-StatusMessage "Error: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Toggle admin failed: $($_.Exception.Message)"
        Show-Dialog "Failed to modify privileges:`n`n$($_.Exception.Message)" "Error" -Icon Error
    }
}

function Toggle-BuiltInAdmin {
    try {
        $admin = Get-LocalUser -Name "Administrator" -ErrorAction Stop
        
        if ($admin.Enabled) {
            $result = Show-Dialog "Disable the built-in Administrator account?`n`nIt will be hidden from login screen." "Confirm" -Buttons YesNo -Icon Warning
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Disable-LocalUser -Name "Administrator" -ErrorAction Stop
                Write-StatusMessage "Built-in Administrator disabled" -Type "Success"
                Show-Dialog "Built-in Administrator account is now disabled" "Success" -Icon Information
            }
        } else {
            $result = Show-Dialog "Enable the built-in Administrator account?`n`nIt will appear on login screen." "Confirm" -Buttons YesNo -Icon Warning
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Enable-LocalUser -Name "Administrator" -ErrorAction Stop
                Write-StatusMessage "Built-in Administrator enabled" -Type "Success"
                Show-Dialog "Built-in Administrator account is now enabled" "Success" -Icon Information
            }
        }
        
        Update-UserGrid
    }
    catch {
        Write-StatusMessage "Error: $($_.Exception.Message)" -Type "Error"
        Write-ErrorLog "Toggle built-in admin failed: $($_.Exception.Message)"
        Show-Dialog "Failed to modify Administrator account:`n`n$($_.Exception.Message)" "Error" -Icon Error
    }
}

# ============================================================================
# CREATE MAIN FORM
# ============================================================================

function New-MainForm {
    try {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "xsukax Windows Users Manager v2.1"
        $form.Size = New-Object System.Drawing.Size(1220, 720)
        $form.StartPosition = "CenterScreen"
        $form.MinimumSize = New-Object System.Drawing.Size(1000, 600)
        
        # Toolbar Panel
        $toolbar = New-Object System.Windows.Forms.Panel
        $toolbar.Dock = "Top"
        $toolbar.Height = 50
        $toolbar.BackColor = [System.Drawing.Color]::WhiteSmoke
        $form.Controls.Add($toolbar)
        
        $x = 10
        $y = 10
        
        # Add User
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(95, 30)
        $btn.Text = "Add User"
        $btn.Add_Click({ Show-AddUserDialog })
        $toolbar.Controls.Add($btn)
        $x += 105
        
        # Properties
        $btnProps = New-Object System.Windows.Forms.Button
        $btnProps.Location = New-Object System.Drawing.Point($x, $y)
        $btnProps.Size = New-Object System.Drawing.Size(95, 30)
        $btnProps.Text = "Properties"
        $btnProps.Add_Click({ Show-UserPropertiesDialog })
        $toolbar.Controls.Add($btnProps)
        $x += 105
        
        # Delete
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(95, 30)
        $btn.Text = "Delete User"
        $btn.Add_Click({ Remove-SelectedUser })
        $toolbar.Controls.Add($btn)
        $x += 115
        
        # Enable
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(80, 30)
        $btn.Text = "Enable"
        $btn.Add_Click({ Enable-SelectedUser })
        $toolbar.Controls.Add($btn)
        $x += 90
        
        # Disable
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(80, 30)
        $btn.Text = "Disable"
        $btn.Add_Click({ Disable-SelectedUser })
        $toolbar.Controls.Add($btn)
        $x += 100
        
        # Change Password
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(130, 30)
        $btn.Text = "Change Password"
        $btn.Add_Click({ Show-ChangePasswordDialog })
        $toolbar.Controls.Add($btn)
        $x += 140
        
        # Manage Groups
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(120, 30)
        $btn.Text = "Manage Groups"
        $btn.Add_Click({ Show-ManageGroupsDialog })
        $toolbar.Controls.Add($btn)
        $x += 130
        
        # Toggle Admin
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(110, 30)
        $btn.Text = "Toggle Admin"
        $btn.BackColor = [System.Drawing.Color]::LightCoral
        $btn.Add_Click({ Toggle-AdminPrivilege })
        $toolbar.Controls.Add($btn)
        $x += 120
        
        # Built-in Admin
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($x, $y)
        $btn.Size = New-Object System.Drawing.Size(120, 30)
        $btn.Text = "Built-in Admin"
        $btn.BackColor = [System.Drawing.Color]::LightYellow
        $btn.Add_Click({ Toggle-BuiltInAdmin })
        $toolbar.Controls.Add($btn)
        
        # DataGridView
        $dgv = New-Object System.Windows.Forms.DataGridView
        $dgv.Location = New-Object System.Drawing.Point(10, 60)
        $dgv.Size = New-Object System.Drawing.Size(1185, 570)
        $dgv.Anchor = "Top,Bottom,Left,Right"
        $dgv.AllowUserToAddRows = $false
        $dgv.AllowUserToDeleteRows = $false
        $dgv.ReadOnly = $true
        $dgv.SelectionMode = "FullRowSelect"
        $dgv.MultiSelect = $false
        $dgv.AutoSizeColumnsMode = "Fill"
        $dgv.BackgroundColor = [System.Drawing.Color]::White
        
        # Columns
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = "Username"
        $col.HeaderText = "Username"
        $col.Width = 150
        $dgv.Columns.Add($col) | Out-Null
        
        $col = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $col.Name = "Enabled"
        $col.HeaderText = "Enabled"
        $col.Width = 80
        $dgv.Columns.Add($col) | Out-Null
        
        $col = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $col.Name = "IsAdmin"
        $col.HeaderText = "Administrator"
        $col.Width = 100
        $dgv.Columns.Add($col) | Out-Null
        
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = "Groups"
        $col.HeaderText = "Groups"
        $col.Width = 250
        $dgv.Columns.Add($col) | Out-Null
        
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = "Description"
        $col.HeaderText = "Description"
        $col.Width = 250
        $dgv.Columns.Add($col) | Out-Null
        
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = "LastLogon"
        $col.HeaderText = "Last Logon"
        $col.Width = 150
        $dgv.Columns.Add($col) | Out-Null
        
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = "PasswordSet"
        $col.HeaderText = "Password Set"
        $col.Width = 120
        $dgv.Columns.Add($col) | Out-Null
        
        # Double-click to open properties
        $dgv.Add_CellDoubleClick({
            Show-UserPropertiesDialog
        })
        
        $form.Controls.Add($dgv)
        $script:DataGrid = $dgv
        
        # Status Bar
        $status = New-Object System.Windows.Forms.StatusStrip
        $status.Dock = "Bottom"
        $form.Controls.Add($status)
        
        $statusLbl = New-Object System.Windows.Forms.ToolStripStatusLabel
        $statusLbl.Spring = $true
        $statusLbl.TextAlign = "MiddleLeft"
        $statusLbl.Text = "Ready"
        $status.Items.Add($statusLbl) | Out-Null
        $script:StatusLabel = $statusLbl
        
        $btnRefresh = New-Object System.Windows.Forms.ToolStripButton
        $btnRefresh.Text = "Refresh (F5)"
        $btnRefresh.Add_Click({ Update-UserGrid })
        $status.Items.Add($btnRefresh) | Out-Null
        
        # Form Load
        $form.Add_Load({
            try {
                Write-StatusMessage "xsukax Windows Users Manager initialized" -Type "Success"
                Update-UserGrid
            }
            catch {
                Write-ErrorLog "Form load failed: $($_.Exception.Message)"
                Show-Dialog "Error loading users: $($_.Exception.Message)" "Error" -Icon Error
            }
        })
        
        # Key Preview for F5
        $form.KeyPreview = $true
        $form.Add_KeyDown({
            if ($_.KeyCode -eq [System.Windows.Forms.Keys]::F5) {
                Update-UserGrid
                $_.Handled = $true
            }
        })
        
        $script:MainForm = $form
        return $form
    }
    catch {
        Write-ErrorLog "New-MainForm failed: $($_.Exception.Message)"
        throw
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

Write-Host "Starting xsukax Windows Users Manager v2.1..." -ForegroundColor Cyan
Write-Host "Created by: xsukax" -ForegroundColor Green
Write-Host "Error log: $script:ErrorLogPath" -ForegroundColor Yellow

$mainForm = New-MainForm
[void]$mainForm.ShowDialog()

} catch {
    # Fatal error - show to user
    $errorMsg = "FATAL ERROR: $($_.Exception.Message)`n`nStack Trace:`n$($_.ScriptStackTrace)`n`nError log: $script:ErrorLogPath"
    Write-ErrorLog $errorMsg
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Fatal Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}