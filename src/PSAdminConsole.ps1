#-------------------------------------------------------------------------------------
# function Show-DomainStatus
#-------------------------------------------------------------------------------------
# Function to show the various status page for
# application and process scheduler services 
#-------------------------------------------------------------------------------------
function Show-DomainStatus($rowNum, [string]$stsCommand)
{
    # Get server details
    $serverName = $global:dtaServers.Rows[$rowNum].Cells[0].Value
    $PSAdminCmd = $global:dtaServers.Rows[$rowNum].Cells[7].Value
    $domainName = $global:dtaServers.Rows[$rowNum].Cells[4].Value
    $serviceType = $global:dtaServers.Rows[$rowNum].Cells[3].Value
    
    if ($serviceType -eq "App") 
    {
        $cmdString1 = "-c"
    }
    if ($serviceType -eq "Prcs") 
    {
        $cmdString1 = "-p"
    }

    # Execute PSAdmin command on remote server
    $cmd = [Scriptblock]::create('cmd.exe /c ' + $global:cmdStringPre + $PSAdminCmd + ' ' + $cmdString1 + ' ' + $stsCommand +' -d ' + $domainName + ' 2>&1')
    $str = invoke-command -ComputerName $serverName -ScriptBlock $cmd | Out-String
    $GridArray = @()

    # Clean up output so it can be processed
    $serverResp = $str.split('>')[1]
    $serverRespArray = $serverResp.Split([Environment]::NewLine)

    # Loop through output from PSAdmin command
    $n = 0
    foreach ($serverRespRow in $serverRespArray) 
    {
        if ($serverRespRow -ne "") # Don't process blank rows
        {
            $n = $n+1
            if ($n -gt 2) # Don't read output headings
            {
                # Create new object to hold results
                $myObject = New-Object System.Object

                if ($stsCommand -eq "qstatus")
                {
                    $myObject | Add-Member -type NoteProperty -name "Server Name" -Value $serverName
                    $myObject | Add-Member -type NoteProperty -name "Domain" -Value $domainName
                    $myObject | Add-Member -type NoteProperty -name "Prog Name" -Value $serverRespRow.substring(0,$serverRespRow.indexof("."))
                    $myObject | Add-Member -type NoteProperty -name "Queue Name" -Value $serverRespRow.substring(15,11).trim()
                    $myObject | Add-Member -type NoteProperty -name "# Serve" -Value $serverRespRow.substring(27,7).trim()
                    $myObject | Add-Member -type NoteProperty -name "Wk Queued" -Value $serverRespRow.substring(35,9).trim()
                    $myObject | Add-Member -type NoteProperty -name "# Queued" -Value $serverRespRow.substring(46,8).trim()
                    $myObject | Add-Member -type NoteProperty -name "Ave Len" -Value $serverRespRow.substring(56,8).trim()
                }
                if ($stsCommand -eq "sstatus")
                {
                    $myObject | Add-Member -type NoteProperty -name "Server Name" -Value $serverName
                    $myObject | Add-Member -type NoteProperty -name "Domain" -Value $domainName
                    $myObject | Add-Member -type NoteProperty -name "Service" -Value $serverRespRow.substring(0,$serverRespRow.indexof("."))
                    $myObject | Add-Member -type NoteProperty -name "Queue Name" -Value $serverRespRow.substring(15,11).trim()
                    $myObject | Add-Member -type NoteProperty -name "Group" -Value $serverRespRow.substring(27,8).trim()
                    $myObject | Add-Member -type NoteProperty -name "Rq Done" -Value $serverRespRow.substring(44,6).trim()
                    $myObject | Add-Member -type NoteProperty -name "Load Done" -Value $serverRespRow.substring(51,9).trim()
                    $myObject | Add-Member -type NoteProperty -name "Status" -Value $serverRespRow.substring(61,9).trim()                
                }
                if ($stsCommand -eq "cstatus")
                {
                    $myObject | Add-Member -type NoteProperty -name "Server Name" -Value $serverName
                    $myObject | Add-Member -type NoteProperty -name "Domain" -Value $domainName
                    $myObject | Add-Member -type NoteProperty -name "User Name" -Value $serverRespRow.substring(16,15).trim()
                    $myObject | Add-Member -type NoteProperty -name "Client Name" -Value $serverRespRow.substring(31,14).trim()
                    $myObject | Add-Member -type NoteProperty -name "Time" -Value $serverRespRow.substring(45,11).trim()
                    $myObject | Add-Member -type NoteProperty -name "Status" -Value $serverRespRow.substring(56,7).trim()
                }
                
                # Add object to an array
                $GridArray += $myObject
            }
        }
    }

    # Output results to grid
    $GridArray | Out-GridView -Title "$serverName $domainName - $serviceType"
}
 
#-------------------------------------------------------------------------------------
# function Get-ServerList
#-------------------------------------------------------------------------------------
# Loop through the servers and determine what
# application, process scheduler and web servers
# are configured and their status
#------------------------------------------------------------------------------------- 
function Get-ServerList
{
    $serverArrayList = New-Object System.Collections.ArrayList 

    # Import xml file to grid array
    if ($actionIn -eq "open") 
    {
        $global:GridArray2 = @()
        $global:GridArray2 = Import-Clixml $filenameIn
    }

    # Process grid if new server list given
    if ($actionIn -eq "new")
    {
        # Import txt file of servers
        $appServers = @()
        $PSAC_ServerList = get-Content $filenameIn
        ForEach ($PSAC_Server in $PSAC_ServerList)
        {
            $appServers += $PSAC_Server
        }  

        $global:GridArray2 = @()
        $lblStatus.Text = "Interrogating servers..."
        $frmMain.refresh() 

        #Loop through App Server Array
        foreach ($Server in $appServers) 
        {
            if ($Server.trim() -ne "")
            {
                get-ServerDomains ($Server)
            }
        }
        $lblStatus.Text = ""
    }

    # Add results to grid
    $serverArrayList.AddRange($global:GridArray2)
    $global:dtaServers.DataSource = $serverArrayList
    
    if ($global:dtaServers.RowCount -gt 0)
    {
        # Hide columns
        $global:dtaServers.Columns["PS_HOME"].Visible = $false
        $global:dtaServers.Columns["PSADMIN"].Visible = $false
        

        # Format form
        $btnGo.Enabled = $true
        $cboActions.Enabled = $true
        $mnuRefresh.Enabled = $true
        $mnuSave.Enabled = $true
    
        if ($actionIn -eq "open") 
        {
            Get-DomainStatusAll
        }
        
        Set-StatusColours
        $global:dtaServers.AutoResizeColumns()
    }
    else
    {
        # Format form
        $btnGo.Enabled = $false
        $cboActions.Enabled = $false
        $mnuRefresh.Enabled = $false
        $mnuSave.Enabled = $false
    }
    
    $frmMain.refresh() 
}

#-------------------------------------------------------------------------------------
# Function Get-DomainStatusAll
#-------------------------------------------------------------------------------------
# Loop through all the rows in the grid and get the current domain status 
#-------------------------------------------------------------------------------------
function Get-DomainStatusAll { 

    # Run through the server list in grid and get the current status
    $lblStatus.Text = "Obtaining service status..."

    # Clear Status column
    for ($i=0; $i -lt $global:dtaServers.RowCount; $i++)
    {
        $global:dtaServers.Rows[$i].Cells[5].Value = ""
        $global:dtaServers.Rows[$i].Cells[5].Style.backcolor = 'white' 
    }
    $frmMain.refresh() 

    # Loop through all the grid rows and obtain status of each domain
    for ($i=0; $i -lt $global:dtaServers.RowCount; $i++)
    {
        Get-DomainStatus($i)
    }
    Set-StatusColours
    
    # Clean up form
    $lblStatus.Text = ""
    $global:dtaServers.AutoResizeColumns()
    $frmMain.refresh() 
} 

#-------------------------------------------------------------------------------------
# Function Set-StatusColours
#-------------------------------------------------------------------------------------
# Loop through all the grid rows and colour the status column
#-------------------------------------------------------------------------------------
function Set-StatusColours {
    for ($i=0; $i -lt $global:dtaServers.RowCount; $i++)
    {
        if ($global:dtaServers.Rows[$i].Cells[5].Value.substring(0,7) -ne "Started" -And $global:dtaServers.Rows[$i].Cells[5].Value.substring(0,7) -ne "Running")
        {
            $global:dtaServers.Rows[$i].Cells[5].Style.backcolor = 'red' 
        }
    }
}

#-------------------------------------------------------------------------------------
# function GenerateForm
#-------------------------------------------------------------------------------------
# Function to build the initial form
#-------------------------------------------------------------------------------------
function GenerateForm 
{ 
     
    # Form Objects 
    $frmMain = New-Object System.Windows.Forms.Form 
    $lblStatus = New-Object System.Windows.Forms.Label 
    $btnGo = New-Object System.Windows.Forms.Button 
    $global:dtaServers = New-Object System.Windows.Forms.DataGridView
    $cboActions = New-Object System.Windows.Forms.ComboBox
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState 
    $mnuMain = new-object System.Windows.Forms.MenuStrip
    $mnuFile = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuOpen = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuNew = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuSave = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuExit = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuRefresh = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuHelp = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuPSACHelp = new-object System.Windows.Forms.ToolStripMenuItem
    $mnuAbout = new-object System.Windows.Forms.ToolStripMenuItem
     
    # Event Script Blocks 
    # Go... Button 
    $btnGo_OnClick=  
    { 
        $Choice = $cboActions.SelectedItem.ToString()
        
        if ($Choice -eq "App / Prcs Server Status") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App" -Or $global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Prcs") {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Started") {
                        Show-DomainStatus $_.Index "sstatus"
                    }
                }
            }
        }
        
        if ($Choice -eq "App / Prcs Client Status") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App" -Or $global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Prcs") {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Started") {
                        Show-DomainStatus $_.Index  "cstatus"
                    }
                }
            }
        }

        if ($Choice -eq "App / Prcs Queue Status") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App" -Or $global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Prcs") {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Started") {
                        Show-DomainStatus $_.Index  "qstatus"
                    }
                }
            }
        }
        
        if ($Choice -eq "Stop Web Service") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Web") {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Running") {
                        Stop-Service -InputObject $(Get-Service -Computer $global:dtaServers.Rows[$_.Index].Cells[0].Value -Name $global:dtaServers.Rows[$_.Index].Cells[4].Value)
                        $global:dtaServers.Rows[$_.Index].Cells[5].Value = 'Hit Refresh!'
                        $global:dtaServers.AutoResizeColumns()
                        [System.Windows.Forms.MessageBox]::Show('Request sent to stop service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value)
                    }
                    else
                    {
                        [System.Windows.Forms.MessageBox]::Show('Error: Service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value + ' not currently running')
                    }
                }
            }
        }
        
        if ($Choice -eq "Start Web Service") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Web") {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Stopped") {
                        Start-Service -InputObject $(Get-Service -Computer $global:dtaServers.Rows[$_.Index].Cells[0].Value -Name $global:dtaServers.Rows[$_.Index].Cells[4].Value)
                        $global:dtaServers.Rows[$_.Index].Cells[5].Value = 'Hit Refresh!'
                        $global:dtaServers.AutoResizeColumns()
                        [System.Windows.Forms.MessageBox]::Show('Request sent to start service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value)
                    }
                    else
                    {
                        [System.Windows.Forms.MessageBox]::Show('Error: Service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value + ' not currently stopped')
                    }
                }
            }
        }
        
        if ($Choice -eq "Restart Web Service") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Web") {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Running") {
                        Restart-Service -InputObject $(Get-Service -Computer $global:dtaServers.Rows[$_.Index].Cells[0].Value -Name $global:dtaServers.Rows[$_.Index].Cells[4].Value)
                        $global:dtaServers.Rows[$_.Index].Cells[5].Value = 'Hit Refresh!'
                        $global:dtaServers.AutoResizeColumns()
                        [System.Windows.Forms.MessageBox]::Show('Request sent to restart service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value)
                    }
                    else
                    {
                        [System.Windows.Forms.MessageBox]::Show('Error: Service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value + ' not currently running')
                    }
                }
            }
        }
        
        if ($Choice -eq "Start App / Prcs Domain") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App" -Or $global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Prcs")
                {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Stopped") 
                    {
                        if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App") {$serviceCode='-c'} else {$serviceCode='-p'}
                        $cmd = [Scriptblock]::create('cmd.exe /c ' + $global:cmdStringPre + $global:dtaServers.Rows[$_.Index].Cells[7].Value + ' ' + $serviceCode +' start -d ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value)
                        $str = invoke-command -ComputerName $global:dtaServers.Rows[$_.Index].Cells[0].Value -ScriptBlock $cmd | Out-Null
                        $global:dtaServers.Rows[$_.Index].Cells[5].Value = 'Hit Refresh!'
                        $global:dtaServers.AutoResizeColumns()
                        [System.Windows.Forms.MessageBox]::Show('Request sent to start service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value)
                    }
                    else
                    {
                        [System.Windows.Forms.MessageBox]::Show('Error: Service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value + ' currently running')
                    }
                }
            }
        }
        if ($Choice -eq "Stop App / Prcs Domain") {
            $global:dtaServers.SelectedRows| ForEach-Object {
                if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App" -Or $global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "Prcs")
                {
                    if ($global:dtaServers.Rows[$_.Index].Cells[5].Value.substring(0,7) -eq "Started") 
                    {
                        if ($global:dtaServers.Rows[$_.Index].Cells[3].Value -eq "App") {$serviceCode='-c'} else {$serviceCode='-p'}
                        $cmd = [Scriptblock]::create('cmd.exe /c ' + $global:cmdStringPre + $global:dtaServers.Rows[$_.Index].Cells[7].Value + ' ' + $serviceCode +' stop -d ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value)
                        $str = invoke-command -ComputerName $global:dtaServers.Rows[$_.Index].Cells[0].Value -ScriptBlock $cmd | Out-Null
                        $global:dtaServers.Rows[$_.Index].Cells[5].Value = 'Hit Refresh!'
                        $global:dtaServers.AutoResizeColumns()
                        [System.Windows.Forms.MessageBox]::Show('Request sent to stop service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value)
                    }
                    else
                    {
                        [System.Windows.Forms.MessageBox]::Show('Error: Service ' + $global:dtaServers.Rows[$_.Index].Cells[4].Value + ' on ' + $global:dtaServers.Rows[$_.Index].Cells[0].Value + ' not currently running')
                    }
                }
            }
        }
    } 


    # File...Open Menu Option
    function OnClick_mnuOpen($Sender,$e){
        $xmlFileName = Open-XMLFile(".\")
        if ($xmlFileName -ne "")
        {
            $actionIn = "open"
            $filenameIn = $xmlFileName
            Get-ServerList  
        }
    }

    # File...New Menu Option
    function OnClick_mnuNew($Sender,$e){
        $txtFileName = Open-TxtFile(".\")
        if ($txtFileName -ne "")
        {
            $actionIn = "new"
            $filenameIn = $txtFileName
            Get-ServerList  
        }
    }

    # File...Save Menu Option
    function OnClick_mnuSave($Sender,$e){
        $fileNameOut = Open-SaveFile(".\")
        if ($fileNameOut -ne "")
        {
            $global:GridArray2 | Export-Clixml -Path $fileNameOut
        }
    }

    # File...Exit Menu Option
    function OnClick_mnuExit($Sender,$e){
        $frmMain.Close() 
    }

    # Refresh... Menu Option
    function OnClick_mnuRefresh($Sender,$e){
        $actionIn = "refresh"
        Get-DomainStatusAll
    }

    # Help...PSACHelp Menu Option
    function OnClick_mnuPSACHelp($Sender,$e){
        Invoke-Expression .\PSAC_Help.html
    }

    # Help...About Option
    function OnClick_mnuAbout($Sender,$e){
        Show-About
    }
    
    # Form object definitions
    # Form Window
    $frmMain.Text = "PeopleSoft Admin Console" 
    $frmMain.Name = "frmMain" 
    $frmMain.DataBindings.DefaultDataSourceUpdateMode = 0 
    $frmMain.MaximizeBox = $false
    $frmMain.FormBorderStyle = 'Fixed3D'
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 820 
    $System_Drawing_Size.Height = 414 
    $frmMain.ClientSize = $System_Drawing_Size 

    # Label
    $lblStatus.TabIndex = 4 
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 600
    $System_Drawing_Size.Height = 20 
    $lblStatus.Size = $System_Drawing_Size 
    $lblStatus.Text = " " 
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 13 
    $System_Drawing_Point.Y = 27 
    $lblStatus.Location = $System_Drawing_Point 
    $lblStatus.DataBindings.DefaultDataSourceUpdateMode = 0 
    $lblStatus.Name = "lblStatus" 
    $frmMain.Controls.Add($lblStatus) 

    # Go... Button 
    $btnGo.TabIndex = 2 
    $btnGo.Name = "btnGo" 
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 75 
    $System_Drawing_Size.Height = 23 
    $btnGo.Size = $System_Drawing_Size 
    $btnGo.UseVisualStyleBackColor = $True 
    $btnGo.Text = "Go..." 
    $btnGo.Enabled = $false
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 450 
    $System_Drawing_Point.Y = 378 
    $btnGo.Location = $System_Drawing_Point 
    $btnGo.DataBindings.DefaultDataSourceUpdateMode = 0 
    $btnGo.add_Click($btnGo_OnClick) 
    $frmMain.Controls.Add($btnGo) 
     
    # Data Grid 
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 795 
    $System_Drawing_Size.Height = 308 
    $global:dtaServers.Size = $System_Drawing_Size 
    $global:dtaServers.DataBindings.DefaultDataSourceUpdateMode = 0 
    $global:dtaServers.Name = "dtaServers" 
    $global:dtaServers.DataMember = "" 
    $global:dtaServers.TabIndex = 0 
    $global:dtaServers.SelectionMode = 'FullRowSelect'
    $global:dtaServers.MultiSelect = $true
    $global:dtaServers.readonly = $true
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 13 
    $System_Drawing_Point.Y = 48 
    $global:dtaServers.Location = $System_Drawing_Point 
    $frmMain.Controls.Add($global:dtaServers) 
    
    $global:dtaServers.add_ColumnHeaderMouseClick($global:dtaServers_ColumnHeaderMouseClick)
    

    # Combo Box (drop down list of commands)
    $cboActions.Location = New-Object System.Drawing.Size(170,378) 
    $cboActions.Size = New-Object System.Drawing.Size(260,20) 
    $cboActions.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList;
    $cboActions.Height = 80
    $cboActions.Enabled = $false
    [void] $cboActions.Items.Add("Choose Action...")
    [void] $cboActions.Items.Add("App / Prcs Server Status")
    [void] $cboActions.Items.Add("App / Prcs Client Status")
    [void] $cboActions.Items.Add("App / Prcs Queue Status")
    [void] $cboActions.Items.Add("Stop Web Service")
    [void] $cboActions.Items.Add("Start Web Service")
    [void] $cboActions.Items.Add("Restart Web Service")
    [void] $cboActions.Items.Add("Start App / Prcs Domain")
    [void] $cboActions.Items.Add("Stop App / Prcs Domain")
    
    $cboActions.SelectedItem = $cboActions.Items[0]
    $frmMain.Controls.Add($cboActions) 

    # Menu Bar
    $mnuMain.Items.AddRange(@(
        $mnuFile,
        $mnuRefresh,
        $mnuSave,
        $mnuExit,
        $mnuNew,
        $mnuHelp,
        $mnuPSACHelp,
        $mnuAbout))
    
    $mnuMain.Location = new-object System.Drawing.Point(0, 0)
    $mnuMain.Name = "mnuMain"
    $mnuMain.Size = new-object System.Drawing.Size(354, 24)
    $mnuMain.TabIndex = 0
    $mnuMain.Text = "menuStrip1"

    # File...
    $mnuFile.DropDownItems.AddRange(@($mnuOpen,$mnuNew,$mnuSave,$mnuExit))
    $mnuFile.Name = "mnuFile"
    $mnuFile.Size = new-object System.Drawing.Size(35, 20)
    $mnuFile.Text = "&File"

    # File...Open
    $mnuOpen.Name = "mnuOpen"
    $mnuOpen.Size = new-object System.Drawing.Size(152, 22)
    $mnuOpen.Text = "&Open..."
    $mnuOpen.Add_Click( { OnClick_mnuOpen $mnuOpen $EventArgs} )

    # File...New
    $mnuNew.Name = "mnuNew"
    $mnuNew.Size = new-object System.Drawing.Size(152, 22)
    $mnuNew.Text = "&New"
    $mnuNew.Add_Click( { OnClick_mnuNew $mnuNew $EventArgs} )

    # File...Save
    $mnuSave.Name = "mnuSave"
    $mnuSave.Size = new-object System.Drawing.Size(152, 22)
    $mnuSave.Text = "&Save As..."
    $mnuSave.Enabled = $false
    $mnuSave.Add_Click( { OnClick_mnuSave $mnuSave $EventArgs} )

    # File...Exit
    $mnuExit.Name = "mnuExit"
    $mnuExit.Size = new-object System.Drawing.Size(152, 22)
    $mnuExit.Text = "E&xit"
    $mnuExit.Add_Click( { OnClick_mnuExit $mnuExit $EventArgs} )

    # Refresh...
    $mnuRefresh.Name = "mnuRefresh"
    $mnuRefresh.Size = new-object System.Drawing.Size(51, 20)
    $mnuRefresh.Text = "&Refresh All..."
    $mnuRefresh.Enabled = $false
    $mnuRefresh.Add_Click( { OnClick_mnuRefresh $mnuRefresh $EventArgs} )

    # Help
    $mnuHelp.DropDownItems.AddRange(@($mnuPSACHelp,$mnuAbout))
    $mnuHelp.Name = "mnuHelp"
    $mnuHelp.Size = new-object System.Drawing.Size(152, 22)
    $mnuHelp.Text = "&Help"

    # Help...PSACHelp
    $mnuPSACHelp.Name = "mnuPSASHelp"
    $mnuPSACHelp.Size = new-object System.Drawing.Size(152, 22)
    $mnuPSACHelp.Text = "PeopleSoft Admin Console Help"
    $mnuPSACHelp.Add_Click( { OnClick_mnuPSACHelp $mnuPSASHelp $EventArgs} )

    # Help...About
    $mnuAbout.Name = "mnuAbout"
    $mnuAbout.Size = new-object System.Drawing.Size(152, 22)
    $mnuAbout.Text = "About..."
    $mnuAbout.Add_Click( { OnClick_mnuAbout $mnuAbout $EventArgs} )

    
    $frmMain.Controls.Add($mnuMain)
    $frmMain.MainMenuStrip = $mnuMain

    # Create Right-Click Menu
    $showCompMan=
    {
        $cmd = [Scriptblock]::create('compmgmt.msc /computer=' + $global:dtaServers.Rows[$rowHit].Cells[0].Value)
        $str = invoke-command -ScriptBlock $cmd | Out-String
    }
    
    $invokeRDP=
    {
        $cmd = [Scriptblock]::create('mstsc /v:' + $global:dtaServers.Rows[$rowHit].Cells[0].Value)
        $str = invoke-command -ScriptBlock $cmd | Out-String
    }

    $showProcesses=
    {
        GenerateProcessForm($global:dtaServers.Rows[$rowHit].Cells[0].Value)
    }

    $refreshStatus=
    {
        Get-DomainStatus($rowHit)
    }


    $contextMenuStrip1=New-Object System.Windows.Forms.ContextMenuStrip

    # Menu Item Title - Computer Management
    [System.Windows.Forms.ToolStripItem]$toolStripTitle = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripTitle.Enabled = $false
    $contextMenuStrip1.Items.Add($toolStripTitle) | Out-Null

    # Menu Item 1 - Computer Management
    [System.Windows.Forms.ToolStripItem]$toolStripItem1 = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripItem1.Text = "Computer Management"
    $toolStripItem1.add_Click($showCompMan)
    $contextMenuStrip1.Items.Add($toolStripItem1) | Out-Null

    # Menu Item 2 - RDP
    [System.Windows.Forms.ToolStripItem]$toolStripItem2 = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripItem2.Text = "RDP"
    $toolStripItem2.add_Click($invokeRDP)
    $contextMenuStrip1.Items.Add($toolStripItem2) | Out-Null

    # Menu Item 3 - Show Processes
    [System.Windows.Forms.ToolStripItem]$toolStripItem3 = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripItem3.Text = "Processes"
    $toolStripItem3.add_Click($showProcesses)
    $contextMenuStrip1.Items.Add($toolStripItem3) | Out-Null

    # Menu Item 4 - Refesh Status
    [System.Windows.Forms.ToolStripItem]$toolStripItem4 = New-Object System.Windows.Forms.ToolStripMenuItem
    $toolStripItem4.Text = "Refresh Status"
    $toolStripItem4.add_Click($refreshStatus)
    $contextMenuStrip1.Items.Add($toolStripItem4) | Out-Null

    
    # Create event of mouse down on datagrid and show menu when right-click occurs
    $global:dtaServers.add_MouseDown({
        $sender = $args[0]
        [System.Windows.Forms.MouseEventArgs]$e= $args[1]

        if ($e.Button -eq  [System.Windows.Forms.MouseButtons]::Right)
        {
            [System.Windows.Forms.DataGridView+HitTestInfo] $hit = $global:dtaServers.HitTest($e.X, $e.Y);
            if ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::Cell)
            {
                # Select the row on the grid
                $rowHit = $hit.RowIndex
                $global:dtaServers.ClearSelection()
                $global:dtaServers.CurrentCell = $global:dtaServers.Rows[$rowHit].Cells[0];
                $global:dtaServers.Rows.Item($rowHit).Selected = $true
                
                # Set the pop up menu title
                $toolStripTitle.Text = $global:dtaServers.Rows[$rowHit].Cells[0].Value + ":"
                
                # Disable Processes menu if server discovery is in error
                if ($global:dtaServers.Rows[$rowHit].Cells[5].Value -eq "Error  ") {
                    $toolStripItem3.Enabled = $false
                }
                else
                {
                    $toolStripItem3.Enabled = $true
                }
                $contextMenuStrip1.Show($global:dtaServers, $e.X, $e.Y)
            }

        }
    })

    # Save the initial state of the form 
    $InitialFormWindowState = $frmMain.WindowState 
    
    # Show the Form 
    $frmMain.ShowDialog()| Out-Null 
    $frmMain.Dispose()
}

#-------------------------------------------------------------------------------------
# function Open-XMLFile
#-------------------------------------------------------------------------------------
# Function to show file dialogue box to get the
# xml file containing the detailed server list
#-------------------------------------------------------------------------------------
function Open-XMLFile($initialDirectory)
{
    if ($Host.name -eq "ConsoleHost") {$ShowHelp = $true}
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect = $false
        ShowHelp = $ShowHelp
        }
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Server Detail List (*.xml)| *.xml"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.filename
}

#-------------------------------------------------------------------------------------
# function Open-TxtFile
#-------------------------------------------------------------------------------------
# Function to show file dialogue box to get the
# txt file containing the basic server list
#-------------------------------------------------------------------------------------
function Open-TxtFile($initialDirectory)
{
    if ($Host.name -eq "ConsoleHost") {$ShowHelp = $true}
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect = $false
        ShowHelp = $ShowHelp
        }
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Server List (*.txt)| *.txt"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.filename
}

#-------------------------------------------------------------------------------------
# function Open-SaveFile
#-------------------------------------------------------------------------------------
# Function to show file dialogue box to get the
# xml filename for the detailed server list output
#-------------------------------------------------------------------------------------
function Open-SaveFile($initialDirectory)
{
    if ($Host.name -eq "ConsoleHost") {$ShowHelp = $true}
    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        ShowHelp = $ShowHelp
        }
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "Server Detail List (*.xml)| *.xml"
    [void]$OpenFileDialog.ShowDialog()
    $OpenFileDialog.filename
}

#-------------------------------------------------------------------------------------
# function Show-About
#-------------------------------------------------------------------------------------
# Function to show the Help...About form
#-------------------------------------------------------------------------------------
function Show-About
{
    Add-Type -AssemblyName System.Windows.Forms

    $frmAbout = New-Object system.Windows.Forms.Form
    $frmAbout.Text = "About..."
    $frmAbout.TopMost = $true
    $frmAbout.Width = 315
    $frmAbout.Height = 164
    $CenterScreen = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $frmAbout.StartPosition = $CenterScreen;
    $frmAbout.MaximizeBox = $false
    $frmAbout.MinimizeBox = $false
    $frmAbout.FormBorderStyle = 'Fixed3D'

    $lblTitle = New-Object system.windows.Forms.Label
    $lblTitle.Text = "PeopleSoft Admin Console"
    $lblTitle.AutoSize = $true
    $lblTitle.Width = 25
    $lblTitle.Height = 10
    $lblTitle.location = new-object system.drawing.point(63,14)
    $lblTitle.Font = "Microsoft Sans Serif,10,style=Bold"
    $frmAbout.controls.Add($lblTitle)

    $lblVersion = New-Object system.windows.Forms.Label
    $lblVersion.Text = "version 1.1"
    $lblVersion.AutoSize = $true
    $lblVersion.Width = 25
    $lblVersion.Height = 10
    $lblVersion.location = new-object system.drawing.point(114,39)
    $lblVersion.Font = "Microsoft Sans Serif,10"
    $frmAbout.controls.Add($lblVersion)

    $lblCopyright = New-Object system.windows.Forms.Label
    $lblCopyright.Text = "© Richard Wood 2018"
    $lblCopyright.AutoSize = $true
    $lblCopyright.Width = 25
    $lblCopyright.Height = 10
    $lblCopyright.location = new-object system.drawing.point(81,86)
    $lblCopyright.Font = "Microsoft Sans Serif,9"
    $frmAbout.controls.Add($lblCopyright)

    [void]$frmAbout.ShowDialog()
    $frmAbout.Dispose()
}

#-------------------------------------------------------------------------------------

# Import the Assemblies 
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null 
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null 

# Import Scripts
. ".\lib\serverTaskManager.ps1"
. ".\lib\getServerEnvs.ps1"

# Settings
$global:cmdStringPre = 'set PS_CUST_HOME=e:\PS_HOME `& set PS_APP_HOME=e:\PS_HOME `& '

# Call function to show main form
GenerateForm

#-------------------------------------------------------------------------------------