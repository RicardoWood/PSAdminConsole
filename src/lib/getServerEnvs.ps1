#-------------------------------------------------------------------------------------
# Function get-ServerDomains
#-------------------------------------------------------------------------------------
# Find all the domains running on the server, currently looks for:
#   - Application Servers
#   - Process Schedulers
#   - Web Servers
#-------------------------------------------------------------------------------------
function get-ServerDomains ([string]$serverName) {

    $serverName = $serverName.ToUpper()
    write-host 'Interrogating server:' $serverName
    $lblStatus.Text = "Interrogating server list..."
            
    # Obtain environment from server name
    switch -wildcard ($serverName)
    {
        "*HCM*"  {$envName = $serverName.SubString(7,3)}
        "*IH*"   {$envName = $serverName.SubString(6,3)}
        "*FSCM*" {$envName = $serverName.SubString(8,3)}
        default  {$envName = "UNKNOWN"}
    }

    # Look up PS_HOME environment variable
    $EnvObj = @(Get-WMIObject -Class Win32_Environment -ComputerName $serverName -EA Stop)
    $Env = $EnvObj | Where-Object {$_.Name -eq "PS_HOME"} 
    $PSHomeEnv = $Env.VariableValue
    
    # If the PS_HOME env Var exists then use psadmin.exe to find all the app and prcs domains
    if ($PSHomeEnv -ne "") 
    {
        $PSAdminCmd = $PSHomeEnv + '\bin\server\WINX86\psadmin.exe'

        $cmd = [Scriptblock]::create('cmd.exe /c set PS_CUST_HOME=e:\PS_HOME `& set PS_APP_HOME=e:\PS_HOME `& ' + $PSAdminCmd + ' -envsummary')
        $str = invoke-command -ComputerName $serverName -ScriptBlock $cmd -ErrorAction Stop | Out-String
    
        $serverRespArray = $str.Split([Environment]::NewLine)
        foreach ($serverRespRow in $serverRespArray) {
            if (($serverRespRow -ne "") -And ($serverRespRow.substring(0,1).trim() -ne "-")) # Don't process blank or dashed rows 
            {
                $serverRespRow = $serverRespRow.PadRight(60,' ')
                if ($serverRespRow.substring(2,19).trim() -eq "PeopleTools Version") {
                    $toolsRel = $serverRespRow.substring(22,7).trim()
                }
                
                if ($serverRespRow.substring(0,30).trim() -eq "PeopleSoft Application Servers") {
                    $serviceType = 'App'
                }
                if ($serverRespRow.substring(0,36).trim() -eq "PeopleSoft Process Scheduler Servers") {
                    $serviceType = 'Prcs'
                }
                
                if ($serverRespRow.substring(3,1).trim() -eq ")") {
                    $svrDomain = $serverRespRow.substring(5,12).trim()
                    $domStatus = $serverRespRow.substring(18,35).trim()
                    
                    #write-host $serverName $SvrDomain $serviceType $toolsRel $domStatus $PSAdminCmd $PSHomeEnv
                    
                    $myObject = New-Object System.Object
                    $myObject | Add-Member -type NoteProperty -name "Server Name" -Value $serverName
                    $myObject | Add-Member -type NoteProperty -name "Environment" -Value $envName
                    $myObject | Add-Member -type NoteProperty -name "Tools Release" -Value $toolsRel
                    $myObject | Add-Member -type NoteProperty -name "Service Type" -Value $serviceType
                    $myObject | Add-Member -type NoteProperty -name "Domain" -Value $svrDomain
                    $myObject | Add-Member -type NoteProperty -name "Status" -Value $domStatus
                    $myObject | Add-Member -type NoteProperty -name "PS_HOME" -Value $PSHomeEnv
                    $myObject | Add-Member -type NoteProperty -name "PSADMIN" -Value $PSAdminCmd
                    $global:GridArray2 += $myObject
                
                }
            }
        }
    }
    
    # Get Web server services
    $serviceType = 'Web'
    get-service -computername $serverName -Name peoplesoft-* | Select Name, Status | ForEach {
        $myObject = New-Object System.Object
        $myObject | Add-Member -type NoteProperty -name "Server Name" -Value $serverName
        $myObject | Add-Member -type NoteProperty -name "Environment" -Value $envName
        $myObject | Add-Member -type NoteProperty -name "Tools Release" -Value $toolsRel
        $myObject | Add-Member -type NoteProperty -name "Service Type" -Value $serviceType
        $myObject | Add-Member -type NoteProperty -name "Domain" -Value $_.Name.ToString()
        get-service -computername $serverName -Name $_.Name.ToString() -ErrorAction Stop | Select Status | ForEach {
            $serviceStatus = $_.Status.ToString()
        }
        $myObject | Add-Member -type NoteProperty -name "Status" -Value $serviceStatus
        $myObject | Add-Member -type NoteProperty -name "PS_HOME" -Value ""
        $myObject | Add-Member -type NoteProperty -name "PSADMIN" -Value ""
        $global:GridArray2 += $myObject
    }
    
    $lblStatus.Text = ""
}


#-------------------------------------------------------------------------------------
# Function Get-DomainStatus
#-------------------------------------------------------------------------------------
# Get domain status for given row
#-------------------------------------------------------------------------------------
function Get-DomainStatus($rowNum) {


        write-host 'Obtaining status for' $global:dtaServers.Rows[$rowNum].Cells[0].Value $global:dtaServers.Rows[$rowNum].Cells[3].Value $global:dtaServers.Rows[$rowNum].Cells[4].Value
        
        if ($global:dtaServers.Rows[$rowNum].Cells[3].Value -eq "Web")
        {
            try
            {
                get-service -computername $global:dtaServers.Rows[$rowNum].Cells[0].Value -Name $global:dtaServers.Rows[$rowNum].Cells[4].Value -ErrorAction Stop | Select Status | ForEach {
                    $global:dtaServers.Rows[$rowNum].Cells[5].Value = $_.Status.ToString()
                }
            }
            Catch
            {
                $global:dtaServers.Rows[$rowNum].Cells[5].Value = "Error  "
                write-host 'Unable to get status of' $global:dtaServers.Rows[$rowNum].Cells[0].Value $global:dtaServers.Rows[$rowNum].Cells[3].Value $global:dtaServers.Rows[$rowNum].Cells[4].Value
                Write-Host $_.Exception.Message -ForegroundColor Green
            }
            
        }
        
        if ($global:dtaServers.Rows[$rowNum].Cells[3].Value -eq "App")
        {
            Try
            {
                $cmd = [Scriptblock]::create('cmd.exe /c ' + $global:cmdStringPre + $global:dtaServers.Rows[$rowNum].Cells[7].Value + ' -c status -d ' + $global:dtaServers.Rows[$rowNum].Cells[4].Value)
                $str = invoke-command -ComputerName $global:dtaServers.Rows[$rowNum].Cells[0].Value -ScriptBlock $cmd -ErrorAction Stop | Out-String
                $global:dtaServers.Rows[$rowNum].Cells[5].Value = $str
            }
            Catch
            {
                $global:dtaServers.Rows[$rowNum].Cells[5].Value = "Error  "
                write-host 'Unable to get status of' $global:dtaServers.Rows[$rowNum].Cells[0].Value $global:dtaServers.Rows[$rowNum].Cells[3].Value $global:dtaServers.Rows[$rowNum].Cells[4].Value
                Write-Host $_.Exception.Message -ForegroundColor Green
            }
        }
        
        if ($global:dtaServers.Rows[$rowNum].Cells[3].Value -eq "Prcs")
        {
            Try
            {
                $cmd = [Scriptblock]::create('cmd.exe /c ' + $global:cmdStringPre + $global:dtaServers.Rows[$rowNum].Cells[7].Value + ' -p status -d ' + $global:dtaServers.Rows[$rowNum].Cells[4].Value)
                $str = invoke-command -ComputerName $global:dtaServers.Rows[$rowNum].Cells[0].Value -ScriptBlock $cmd -ErrorAction Stop | Out-String
                $global:dtaServers.Rows[$rowNum].Cells[5].Value = $str
            }
            Catch
            {
                $global:dtaServers.Rows[$rowNum].Cells[5].Value = "Error  "
                write-host 'Unable to get status of' $global:dtaServers.Rows[$rowNum].Cells[0].Value $global:dtaServers.Rows[$rowNum].Cells[3].Value $global:dtaServers.Rows[$rowNum].Cells[4].Value
                Write-Host $_.Exception.Message -ForegroundColor Green
            }
        }

}

