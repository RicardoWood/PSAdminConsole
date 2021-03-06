function Get-ProcessInfo ([string]$serverName)
{ 
    $arrTaskList = New-Object System.Collections.ArrayList 
    $Script:procInfo = Invoke-Command -ComputerName $serverName {Get-Process | select-object Id,Name,Path,Description,VM,WS,CPU,Company| sort -Property Name}

    $procInfo = $procInfo| Select Id,Name,Description,@{Name='Virtual[MB]';Expression={[math]::Round($_.vm / 1024kb,2)}},@{Name='WorkingSet[MB]';Expression={[math]::Round($_.ws / 1024kb,2)}},CPU,Company,Path
    
    $arrTaskList.AddRange($procInfo) 
    $dataTaskList.DataSource = $arrTaskList 
    $dataTaskList.AutoResizeColumns()
    $frmTaskList.refresh() 
} 
 
# Generate Process Form Function 
function GenerateProcessForm ([string]$serverName)
{ 

    # Import the Assemblies 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    # Form Objects 
    $frmTaskList = New-Object System.Windows.Forms.Form 
    $btnKill = New-Object System.Windows.Forms.Button 
    $button1 = New-Object System.Windows.Forms.Button 
    $dataTaskList = New-Object System.Windows.Forms.DataGridView
    $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState 
 
    #---------------------------------------------- 
    # Event Script Blocks 
    #---------------------------------------------- 
    $button1_OnClick=  
    { 
        Get-ProcessInfo($serverName)
    } 
 
    $btnKill_OnClick=  
    { 
        $dataTaskList.SelectedRows| ForEach-Object{
            $procid = $dataTaskList.Rows[$_.Index].Cells[0].Value
            $procName = $dataTaskList.Rows[$_.Index].Cells[1].Value
            $msgBoxInput =  [System.Windows.Forms.MessageBox]::Show('Kill Process '+$procid + ': ' + $procName,'Kill Process?','YesNo','Error')
            switch ($msgBoxInput) {
                'Yes' {
                        $cmd = [Scriptblock]::create('Get-Process -Id '+$procid+'| Stop-Process')
                        Invoke-Command -ComputerName $serverName -ScriptBlock $cmd
                        [System.Windows.Forms.MessageBox]::Show('Request sent to kill process ' + $procid + ': ' + $procName + ' sent.')
                        Get-ProcessInfo($serverName)
                      }
            }
        }
    } 
 
    $OnLoadForm_UpdateGrid= 
    { 
        Get-ProcessInfo($serverName)
    } 
 
    #---------------------------------------------- 
    # Generate Form Code 
    $frmTaskList.Text = $serverName
    $frmTaskList.Name = "form1" 
    $frmTaskList.DataBindings.DefaultDataSourceUpdateMode = 0 
    $frmTaskList.DataBindings.DefaultDataSourceUpdateMode = 0 
    $frmTaskList.MaximizeBox = $false
    $frmTaskList.MinimizeBox = $false
    $frmTaskList.FormBorderStyle = 'Fixed3D'
    $CenterScreen = [System.Windows.Forms.FormStartPosition]::CenterScreen;
    $frmTaskList.StartPosition = $CenterScreen;
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 700 
    $System_Drawing_Size.Height = 414 
    $frmTaskList.ClientSize = $System_Drawing_Size 
 
    $btnKill.TabIndex = 2 
    $btnKill.Name = "button2" 
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 75 
    $System_Drawing_Size.Height = 23 
    $btnKill.Size = $System_Drawing_Size 
    $btnKill.UseVisualStyleBackColor = $True 
     
    $btnKill.Text = "Kill Process" 
     
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 610 
    $System_Drawing_Point.Y = 378 
    $btnKill.Location = $System_Drawing_Point 
    $btnKill.DataBindings.DefaultDataSourceUpdateMode = 0 
    $btnKill.add_Click($btnKill_OnClick) 
     
    $frmTaskList.Controls.Add($btnKill) 
     
    $button1.TabIndex = 1 
    $button1.Name = "button1" 
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 75 
    $System_Drawing_Size.Height = 23 
    $button1.Size = $System_Drawing_Size 
    $button1.UseVisualStyleBackColor = $True 
     
    $button1.Text = "Refresh" 
     
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 13 
    $System_Drawing_Point.Y = 379 
    $button1.Location = $System_Drawing_Point 
    $button1.DataBindings.DefaultDataSourceUpdateMode = 0 
    $button1.add_Click($button1_OnClick) 
     
    $frmTaskList.Controls.Add($button1) 
     
    $System_Drawing_Size = New-Object System.Drawing.Size 
    $System_Drawing_Size.Width = 674
    $System_Drawing_Size.Height = 348 
    $dataTaskList.Size = $System_Drawing_Size 
    $dataTaskList.DataBindings.DefaultDataSourceUpdateMode = 0 
    $dataTaskList.Name = "dataGrid1" 
    $dataTaskList.DataMember = "" 
    $dataTaskList.TabIndex = 0 
    $dataTaskList.SelectionMode = 'FullRowSelect'
    $System_Drawing_Point = New-Object System.Drawing.Point 
    $System_Drawing_Point.X = 13 
    $System_Drawing_Point.Y = 13 
    $dataTaskList.Location = $System_Drawing_Point 
     
    $frmTaskList.Controls.Add($dataTaskList) 
     
    #endregion Generated Form Code 
     
    #Save the initial state of the form 
    $InitialFormWindowState = $frmTaskList.WindowState 
     
    #Add Form event 
    $frmTaskList.add_Load($OnLoadForm_UpdateGrid) 
     
    #Show the Form 
    [void]$frmTaskList.ShowDialog()
    $frmTaskList.Dispose()
     
} # End GenerateProcessForm Function 
     
