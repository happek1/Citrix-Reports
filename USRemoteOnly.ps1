#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# Script Name: RemoteReportcitrixonly.ps1
# Modified: 6/24/2016
#
# Script Requirements: Citrix.XenApp.Commands
# Description: This script gets Health Status of XenApp Farm Remote Servers 
#              Based on script by Jason Poyner's (http://deptive.co.nz/xenapp-farm-health-report)
#              and Stan Czerno (http://www.czerno.com/blog/post/2014/06/12/powershell-script-to-monitor-a-citrix-xenapp-farm-s-health)
#              
# Modified by khappe to only include Remote Desk Citrix information, script creates a PS Session to a XenApp Admin server then performs PS commands
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# User Definable Variables
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


# Email Settings
# Multiple email addresses example: "email@domain.com,email2@domain.com"
$emailFrom     = "noc@reedsmith.com"
$emailTo       = "noc@reedsmith.com"
#$emailCC       = "citrixsupport@reedsmith.com, khappe@reedsmith.com" 
$smtpServer    = "smtp-pdc"
$SendAlerts    = $false

$SendEmailWarnings = $false
$SendEmailErrors = $true

$StrGroup = "Desktop Remote"
$excludedFolders = @("")


# Optional: Excluded folders from health check. 
# Example: @("Servers/Application", "Servers/Std Instances")
$excludedFolders = @("")

# Server to be excluded with any particular name or name filter. 
# Example: @("SRV1","SRV2")
$ServerFilter = @("")

# The maximum uptime days a server can report green. 
$maxUpTimeDays = 3

# License Server Name for license usage.
$LicenseServer = ""

#Admin Server for PSSESSSION
$AdminServer = ""

# License Type to be defined 
# Example: @("MPS_PLT_CCU", "MPS_ENT_CCU", "XDT_ENT_UD") 
#$LicenseTypes = @("MPS_ENT_CCU")
$LicenseTypes = @("MPS_PLT_CCU", "MPS_ENT_CCU", "XDT_ENT_UD")

# Alert Spam Protection Timeframe, in seconds
$EmailAlertsLagTime = "1800"

# Webserver Page
$HTMLPage   = "US Desktop Remote Status.html"


# TimeFrame for Emails
$int_Email_Start = 5
$int_Email_End = 20


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# END - User Definable Variables
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# System Variables
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


# Script Start Time
$script:startTime = Get-Date

# Reads the current directory path from the location of this file
$currentDir = Split-Path $MyInvocation.MyCommand.Path
$currentScriptName = $MyInvocation.MyCommand.Name

#added by Kris - creates PSSESSIOn for Remote Citrix PS Tools
#Loads the Citrix Modules
$PSSessionOption = New-PSSessionOption -OpenTimeOut 600000 -OperationTimeout 600000
$s = new-pssession -computer $AdminServer -SessionOption $PSSessionOption
Invoke-Command -session $s -script { asnp Citrix* }
Import-PSSession -session $s -module Citrix*


# Get farm details once to use throughout the script
$FarmDetails = Get-XAFarm 
$CitrixFarmName = $FarmDetails.FarmName
$WebPageTitle = "$CitrixFarmName Health Status"

# Email Subject with the farm name
$emailSubject  = "US Desktop Remote Status" 

# Log files created in the location of script. 
$LogFile = Join-Path $currentDir ("US Desktop Remote Status.log")
$PreviousLogFile = Join-Path $currentDir ("US Desktop Remote Status_PreviousRun.log")
$ResultsHTML = Join-Path $currentDir ("US Desktop Remote Status.html")
$AlertsFile = Join-Path $currentDir ("US Desktop Remote Status_Alerts.log")
$PreviousAlertsFile = Join-Path $currentDir ("US Desktop Remote Status_PreviousAlerts.log")
$AlertsEmailFile = Join-Path $currentDir ("US Desktop Remote Status_Email.log")

# Table headers
$headerNames  = "Logons", "UpTimeDays", "ServerLoad","ActiveUsers", "DiscUsers"
$headerWidths =    "6",      "8",           "8",            "6",         "6"

# Cell Colors
$ErrorStyle = "style=""background-color: #000000; color: #FF3300;"""
$WarningStyle = "style=""background-color: #000000;color: #FFFF00;"""


# The variable to count the server names
[int]$TotalServers = 0; $TotalServersCount = 0

$allResults = @{}



#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# END - System Variables
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# Functions
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


Function LogMe() 
    {
    
    Param ( [parameter(Mandatory = $true, ValueFromPipeline = $true)] $logEntry,
	   [switch]$display,
	   [switch]$error,
	   [switch]$warning
	   )

    If ($error) { Write-Host "$logEntry" -Foregroundcolor Red; $logEntry = "[ERROR] $logEntry" }
	ElseIf ($warning) { Write-Host "$logEntry" -Foregroundcolor Yellow; $logEntry = "[WARNING] $logEntry"}
	ElseIf ($display) { Write-Host "$logEntry" -Foregroundcolor Green; $logEntry = "$logEntry" }
    Else { Write-Host "$logEntry"; $logEntry = "$logEntry" }

	$logEntry | Out-File $LogFile -Append
    
    } #End Function: LogMe


Function CheckLicense() 
    {

	If (!$LicenseServer) { "No License Server Name defined" | LogMe -Error; $LicenseResult = " Error; Check Detailed Logs "; return $LicenseResult }
	If (!$LicenseTypes) { "No License Type defined" | LogMe -Error; $LicenseResult = " Error; Check Detailed Logs "; return $LicenseResult }
	
	[int]$TotalLicense = 0; [int]$InUseLicense = 0; [int]$PercentageLS = 0; $LicenseResult = " "
	
	Try 
	    {
        
        If (Get-Service -Display "Citrix Licensing" -ComputerName $LicenseServer -ErrorAction Stop) { "Citrix Licensing service is available." | LogMe -Display }
		Else { "Citrix Licensing' service is NOT available." | LogMe -Error; Return "Error; Check Logs" }

		Try
            {
         	
            If ($licensePool = gwmi -class "Citrix_GT_License_Pool" -Namespace "ROOT\CitrixLicensing" -comp $LicenseServer -ErrorAction Stop) 
				{
                
                "License Server WMI Class file found" | LogMe -Display 
				$LicensePool | ForEach-Object{ 
                    
                    Foreach ($Ltype in $LicenseTypes)
						{

                        If ($_.PLD -match $Ltype) { $TotalLicense = $TotalLicense + $_.count; $InUseLicense = $InUseLicense + $_.InUseCount }

                        }

                    }
	
				"The total number of licenses available: $TotalLicense " | LogMe -Display
				"The number of licenses are in use: $InUseLicense " | LogMe -Display
				If (!(($InUseLicense -eq 0) -or ($TotalLicense -eq 0 ))) { $PercentageLS = (($InUseLicense / $TotalLicense ) * 100); $PercentageLS = "{0:N2}" -f $PercentageLS }
				
				If ($PercentageLS -gt 90) { "The License usage is $PercentageLS % " | LogMe -Error }
				ElseIf ($PercentageLS -gt 80) { "The License usage is $PercentageLS % " | LogMe -Warning }
				Else { "The License usage is $PercentageLS % " | LogMe -Display }
	
				$LicenseResult = "$InUseLicense/$TotalLicense [ $PercentageLS % ]"; return $LicenseResult

                }

            }
        Catch
            { 
            
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            "License Server WMI Class file failed. An Error Occured while capturing the License information" | LogMe -Error
            "You may need to uninstall your License Server and reinstall." | LogMe -Error
            "There are known issues with doing an in place upgrade of the license service." | LogMe -Error
			$LicenseResult = " Error; Check Detailed Logs "; return $LicenseResult 

            }

        }
    Catch { "Error returned while checking the Licensing Service. Server may be down or some permission issue" | LogMe -error; return "Error; Check Detailed Logs" }

    } #End Function: CheckLicense


Function CheckContextSwitch() 
    {
	
    Param ($hostname)

	Try 
        {

        $ContextSwitch = (get-wmiobject -Computer $hostname -Class "Win32_PerfFormattedData_PerfOS_System" -ErrorAction Stop | Select-Object ContextSwitchesPersec )
        $ContextSwitchPref = $ContextSwitch.ContextSwitchesPersec

        return $ContextSwitchPref

        }
    Catch { "Error returned while checking the Context Switch performance. Perfmon Counters may be at fault." | LogMe -error; return 101 } 

    } #End Function: CheckContextSwitch
    

Function Ping ([string]$hostname, [int]$timeout) 
    {

    $ping = new-object System.Net.NetworkInformation.Ping #creates a ping object
	
    Try { $result = $ping.send($hostname, $timeout).Status.ToString() }
    Catch { $result = "Failed" }
	
    return $result
    
    } #End Function: Ping



Function writeHtmlHeader 
    { 

	Param ($title, $fileName)
	
    $date = ( Get-Date -format g)
    $head = @"
    <html>
    <head>
    <meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
    <meta http-equiv="refresh" content="60">
    <title>$title</title>
    <STYLE TYPE="text/css">
    <!--
    td {
        font-family: Tahoma;
        font-size: 12px;
        border-top: 1px solid #999999;
        border-right: 1px solid #999999;
        border-bottom: 1px solid #999999;
        border-left: 1px solid #999999;
        padding-top: 0px;
        padding-right: 0px;
        padding-bottom: 0px;
        padding-left: 0px;
        overflow: hidden;}

    .header {
	    font-family: Tahoma;
		font-size: 12px;
		font-weight:bold;
		border-top: 1px solid #999999;
		border-right: 1px solid #999999;
		border-bottom: 1px solid #999999;
		border-left: 1px solid #999999;
		padding-top: 0px;
		padding-right: 0px;
		padding-bottom: 0px;
		padding-left: 0px;
        overflow: hidden;
		color:#FFFFFF;
		text-shadow:2px 2px 10px #000000;

        }
    body {
        margin-left: 5px;
        margin-top: 5px;
        margin-right: 0px;
        margin-bottom: 10px;
        table {
            table-layout:fixed;
            border: thin solid #FFFFFF;}
	.shadow {
		height: 1em;
		filter: Glow(Color=#000000,
		Direction=135,
		Strength=5);}
        -->
    </style>
    </head>
    <body>
    
        <table width='100%'>
        <tr bgcolor='#CCCCCC'>
        <td width=33% align='center' valign="middle">
        <font face='Tahoma' color='#8A0808' size='2'><strong>Site ($StrGroup) Last Queried: $date</strong></font>
        </td>
        </tr>
        </table>
       
        <table width='100%'>
        <tr bgcolor='#CCCCCC'>
        <td width=50% align='center' valign="middle">
        <font face='Tahoma' color='#003399' size='2'><strong>Number of Servers in Site: $TotalServersCount</strong></font>
        <td width=50% align='center' valign="middle">
        <font face='tahoma' color='#003399' size='2'><strong>Citrix License Usage:  $LicenseReport</strong></font>
        </td>
        </tr>
        </table>
        <table width='100%'>
        <tr bgcolor='#CCCCCC'>
        <td width=50% align='center' valign="middle">
        <font face='Tahoma' color='#003399' size='2'><strong>Number of Available Servers in Remote Silo: $CountIn </strong></font>
        <td width=50% align='center' valign="middle">
        <font face='tahoma' color='#003399' size='2'><strong>Number of Unavailable Servers in Remote Silo: $CountOut </strong></font>
        </td>
        </tr>
        </table>
        
"@

    $head | Out-File $fileName

    } #End Function: writeHtmlHeader 


Function writeTableHeader 
    { 
	
    Param ($fileName)
	
    $tableHeader = @"
    <table width='100%'><tbody>
    <tr bgcolor=#CCCCCC>
    <td width='12%' align='center'><strong>ServerName</strong></td>
"@

    $i = 0

    While ($i -lt $headerNames.count)
        {

        $headerName = $headerNames[$i]
        $headerWidth = $headerWidths[$i]
        #$tableHeader += "<td width='" + $headerWidth + "%' align='center'><strong>$headerName</strong></td>"
        $tableHeader += "<td align='center'><strong>$headerName</strong></td>"
        $i++
        
        }

    $tableHeader += "</tr>"
    $tableHeader | Out-File $fileName -append 

    } #End Function: writeTableHeader 


Function writeData 
    {

	Param ($data, $fileName)
	
	$data.Keys | sort | foreach {

        $tableEntry += "<tr>"
    	$computerName = $_
	    #$tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'><a href='.\HostStatus.html?$computerName'>$computerName</a></font></td>")
        $tableEntry += ("<td bgcolor='#CCCCCC' align=center><font color='#003399'>$computerName</font></td>")

	    $headerNames | foreach {
            Try
                {
			
                If ($data.$computerName.$_[1] -eq $null ) { $bgcolor = "#FF0000"; $fontColor = "#FFFFFF"; $testResult = "Err" }
                Else
                    {
				
                    If ($data.$computerName.$_[0] -eq "SUCCESS") { $bgcolor = "#387C44"; $fontColor = "#FFFFFF" }
	    			ElseIf ($data.$computerName.$_[0] -eq "WARNING") { $bgcolor = "#F5DA81"; $fontColor = "#000000" }
		    		ElseIf ($data.$computerName.$_[0] -eq "ERROR") { $bgcolor = "#FF0000"; $fontColor = "#000000" }
			    	Else { $bgcolor = "#CCCCCC"; $fontColor = "#003399" }
				
            	    $testResult = $data.$computerName.$_[1]

    				}
            
                }
            Catch { $bgcolor = "#CCCCCC"; $fontColor = "#003399"; $testResult = "N/A" }

		    $tableEntry += ("<td bgcolor='" + $bgcolor + "' align=center><font color='" + $fontColor + "'>$testResult</font></td>")

		    }

	    $tableEntry += "</tr>"

	    }

	$tableEntry | Out-File $fileName -append

    } #End Function: writeData


Function FindErrors 
    {
    
	Param ($data)

    Add-Content $AlertsFile "Server,Type,Component,Value"

    If (Test-Path $AlertsEmailFile) { RM $AlertsEmailFile } 
   
    Add-Content $AlertsEmailFile "Server,Type,Component,Value,Status"

    $data.Keys | sort | foreach {

        $computerName = $_
        $headerNames | foreach {
        
            Try
                {

                If ($data.$computerName.$_[1] -eq $null ) { $testResult = "Err" }
                Else
                    {

                    If (($data.$computerName.$_[0] -eq "WARNING") -Or ($data.$computerName.$_[0] -eq "ERROR"))
                        { 
                        
                        $strPreviousAlert =""

                        $alertServer = $computerName 
                        $alertType = $data.$computerName.$_[0]
                        #$alertValue = $data.$computerName.$_[1] | out-string
                        $alertValue = $data.$computerName.$_[1]
                        $alertComp = $_

                        $strOutput = $AlertServer +',' +$alertType +',' +$AlertComp +',' +$AlertValue

                        Add-Content $AlertsFile $strOutput

                        If (Test-Path $PreviousAlertsFile)
                            {
                        
                            $strPreviousAlert = Import-CSV $PreviousAlertsFile | Where { $_.Server -eq $AlertServer -And $_.Type -eq $alertType -And $_.Component -eq $alertComp }
                                    
                            If (($strPreviousAlert -ine "") -And ($strPreviousAlert -ine $null))
                                { 

                                ForEach ($aline in $strPreviousAlert)
                                    {

                                    If ($alertComp -ieq $aline.Component)
                                        {
                                                                        
                                        If ($alertValue -ine $aline.Value)
                                            {

                                            If ($alertValue -gt $aline.Value)
                                                {

                                                Write-Host "Alert Exisits: " $strOutput
                                                Write-Host $AlertServer "Alert " $alertComp "value has Increased: " $aline.Value " -> " $alertValue

                                                If ($alertType -ieq "ERROR")
                                                    {

                                                    $strEmailOutput = $strOutput +',Increased'
                                                    Add-Content $AlertsEmailFile $strEmailOutput

                                                    }

                                                }

                                            }

                                        If ($alertType -ine $aline.Type)
                                            {

                                            Write-Host "Alert Type Changed:" $strOutput

                                            If ($alertType -eq "ERROR")
                                                {

                                                $strEmailOutput = $strOutput +',Changed'
                                                Add-Content $AlertsEmailFile $strEmailOutput

                                                }

                                            }
                                        }
                                    Else
                                        {

                                        Write-Host "Alert Does not Exisit:" $strOutput

                                        $strEmailOutput = $strOutput +',New'
                                        Add-Content $AlertsEmailFile $strEmailOutput

                                        }

                                    }

                                }
                            Else
                                {

                                Write-Host "Alert Does not Exisit:" $strOutput

                                $strEmailOutput = $strOutput +',New'
                                Add-Content $AlertsEmailFile $strEmailOutput

                                }

                            }
                        Else
                            {
					
                            Write-Host "Alert Does not Exisit:" $strOutput

                            $strEmailOutput = $strOutput +',New'
                            Add-Content $AlertsEmailFile $strEmailOutput

                            }

                        }

                    $testResult = $data.$computerName.$_[1]

				    }
                } 
            Catch { $testResult = "N/A" }

            }

        }

    } #End Function: FindErrors
 


Function GetElapsedTime([datetime]$starttime) 
    {

    $runtime = $(get-date) - $starttime
    $retStr = [string]::format("{0} minute(s) and {1} sec(s)", $runtime.minutes, $runtime.seconds)
    $retStr

    } #End Function: GetElapsedTime 

Function writeHtmlFooter { 
	param($fileName)

$elapsed = GetElapsedTime $script:startTime


$footer=("<br><font face='HP Simplified' color='#003399' size='2'><br><I>Total Elapsed Script Time  {0}.<br> Script Hosted on server {1}.<br>Script Name : {2}<br>Script Path: {3}</font>" -f ($elapsed),$env:COMPUTERNAME,$currentScriptName,$currentdir) 
@"
</table>
</body>
</html>
"@ | Out-File $FileName -append
$footer | Out-File $FileName -append

}


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# END - Functions
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# Main Program
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


If (Test-Path $LogFile) { Copy $LogFile $PreviousLogFile }

If (Test-Path $AlertsFile) 
    {

    Copy $AlertsFile $PreviousAlertsFile 
    Clear-Content $AlertsFile 
        
    }

RM $LogFile -force -EA SilentlyContinue

"Script Started at $script:startTime" | LogMe -display 
"Processing Location: $StrGroup" | LogMe -display
" " | LogMe -display


"Checking Citrix License usage on $LicenseServer" | LogMe -Display 
$LicenseReport = CheckLicense


" " | LogMe; "Checking Citrix XenApp Server Health." | LogMe ; " " | LogMe


Get-XAServer | where-object {$_.FolderPath -match $StrGroup} | Sort-Object ServerName | % { 

    $tests = @{}	
    
    # Check If Server is in Excluded Folder path or server list
    #If ($excludedFolders -contains $_.FolderPath) { $_.FolderPath + " in excluded Server folder - skipping" | LogMe -Display; "" | LogMe; return }
    #If ($ServerFilter -contains $_.ServerName) { $_.ServerName + " is excluded in the Server List  - skipping" | LogMe -Display; "" | LogMe; return }

    [int]$TotalServers = [int]$TotalServers + 1; $server = $_.ServerName

    "Server Name: $server" | LogMe

    $ProcessTime = Get-Date -format R
    "Processing Time: $ProcessTime" | LogMe

   
$sessions = Get-XASession -Farm | where-object {$_.ServerName -match $WorkerGroups}

    # Ping Remote Server
    $result = Ping $server 1000
    If ($result -ne "SUCCESS") 
        { 

        $tests.Ping = "ERROR", $result; "NOT able to ping - skipping " | LogMe -error 

        }
    Else 
        {   

        $tests.Ping = "SUCCESS", $result;  "Server is responding to ping" | LogMe -Display
        

        # Check Logon Mode
        $logonMode = $_.LogonMode
        If ($logonMode -ine 'AllowLogOns'){ "Logon Mode: $logonMode " | LogMe -error; $tests.Logons = "ERROR", "Disabled" }
        Else { "Logon Mode: $logonMode " | LogMe -Display; $tests.Logons = "SUCCESS", "Enabled" }
        

    	# Get Active Sessions
	    $activeServerSessions = [array]($sessions | ? {$_.State -eq "Active" -and $_.Protocol -ne "Console" -and $_.ServerName -match $server})

	    If ($activeServerSessions) { $totalActiveServerSessions = $activeServerSessions.count }
    	Else { $totalActiveServerSessions = 0 }

        $tests.ActiveUsers = "SUCCESS", $totalActiveServerSessions 

        $ICAActiveSessions = [array]($sessions | ? {$_.State -eq "Active" -and $_.Protocol -eq "Ica" -and $_.ServerName -match $server})
        $ICATotalActiveServerSessions = $ICAActiveSessions.Count
        
        $RDPActiveSessions = [array]($sessions | ? {$_.State -eq "Active" -and $_.Protocol -eq "Rdp" -and $_.ServerName -match $server})
        $RDPTotalActiveServerSessions = $RDPActiveSessions.Count
        
        "Active sessions: ICA $ICATotalActiveServerSessions | RDP $RDPTotalActiveServerSessions" | LogMe -display

    
        # Get Disconnected Sessions
	    $discServerSessions = [array]($sessions | ? {$_.State -eq "Disconnected" -and $_.Protocol -ne "Console" -and $_.ServerName -match $server})
	
        If ($discServerSessions) { $totalDiscServerSessions = $discServerSessions.count } 
	    Else { $totalDiscServerSessions = 0 }
    
        $tests.DiscUsers = "SUCCESS", $totalDiscServerSessions 

        $ICADiscSessions = [array]($sessions | ? {$_.State -eq "Disconnected" -and $_.Protocol -eq "Ica" -and $_.ServerName -match $server})
        $ICATotalDiscServerSessions = $ICADiscSessions.Count
    
        $RDPDiscSessions = [array]($sessions | ? {$_.State -eq "Disconnected" -and $_.Protocol -eq "Rdp" -and $_.ServerName -match $server})
        $RDPTotalDiscServerSessions = $RDPDiscSessions.Count
    
        "Disconnected sessions: ICA $ICATotalDiscServerSessions | RDP $RDPTotalDiscServerSessions" | LogMe -display
   

        # Warning If Disconnected Sessions Greater Than Active Sessions.
        #If ($totalDiscServerSessions -gt $totalActiveServerSessions) { $tests.DiscUsers = "WARNING", $totalDiscServerSessions }
		# Coloring of connected users
		If ($totalActiveServerSessions -gt 0) { $tests.ActiveUsers = "WARNING", $totalActiveServerSessions }
		If ($totalDiscServerSessions -gt 0) { $tests.DiscUsers = "WARNING", $totalDiscServerSessions }
		
            	
        # Check Server Load
	    If ($tests.ping[0] -eq "Success")
            {
            
            $CurrentServerLoad = Get-XAServerLoad -ServerName $server

		    If ([int] $CurrentServerLoad.load -lt 7500) 
                { 
			
                If ([int] $CurrentServerLoad.load -eq 0) { $tests.ActiveUsers = "SUCCESS", $totalActiveServerSessions; $tests.DiscUsers = "SUCCESS", $totalDiscServerSessions }

				"Serverload is normal [ $CurrentServerload ]" | LogMe -display; $tests.Serverload = "SUCCESS", ($CurrentServerload.load) 

                }
			ElseIf ([int] $CurrentServerLoad.load -lt 8500) { "Serverload is Medium [ $CurrentServerload ]" | LogMe -warning; $tests.Serverload = "WARNING", ($CurrentServerload.load) }
			ElseIf ([int] $CurrentServerLoad.load -eq 20000) { "Serverload Fault [ Could not Contact License Server ]" | LogMe -Error; $tests.Serverload = "ERROR", "LS Err" }    
			ElseIf ([int] $CurrentServerLoad.load -eq 99999) { "Serverload Fault [ No Load Evaluator is Configured ]" | LogMe -Error; $tests.Serverload = "ERROR", "No LE" }
			ElseIf ([int] $CurrentServerLoad.load -eq 10000) { "Serverload Full [ $CurrentServerload ]" | LogMe -Error; $tests.Serverload = "ERROR", ($CurrentServerload.load) }
			Else { "Serverload is High [ $CurrentServerload ]" | LogMe -error; $tests.Serverload = "ERROR", ($CurrentServerload.load) }

            }
        Else { "Server load can't be determine since PING failed " | LogMe -error; $tests.Serverload = "ERROR", "IMA Err" }

        $CurrentServerLoad = 0

	
        # Test WMI
        $tests.WMI = "ERROR","Error"
	    Try { $wmi=Get-WmiObject -class Win32_OperatingSystem -computer $_.ServerName } 
	    Catch {	$wmi = $null }
		
        # Perform WMI related checks and time check
	    If ($wmi -ne $null) 
            {

		    $tests.WMI = "SUCCESS", "Success"; "WMI connection success" | LogMe -display
		    $LBTime=$wmi.ConvertToDateTime($wmi.Lastbootuptime)
		    [TimeSpan]$uptime=New-TimeSpan $LBTime $(get-date)
		    
            If ($uptime.days -gt $maxUpTimeDays) { "Server reboot warning, last reboot: {0:D}" -f $LBTime | LogMe -warning; $tests.UptimeDays = "WARNING", $uptime.days } 
            Else { "Server uptime days: $uptime" | LogMe -display; $tests.UptimeDays = "SUCCESS", $uptime.days } 
            
            } 
        Else { "WMI connection failed - check WMI for corruption" | LogMe -error }

            
        If (Get-WmiObject win32_computersystem -ComputerName $server -ErrorAction SilentlyContinue) 
            {
            
            $tests.RPC = "SUCCESS", "Success"; "RPC responded" | LogMe -Display  
            
            }
        Else { $tests.RPC = "ERROR", "No Response"; "RPC failed" | LogMe -error }
        }
    
	$allResults.$server = $tests

    " " | LogMe -display


    }

#Server Counts
$TotalServersCount = [int]$TotalServers

$SiloServers = Get-XAServer | where-object {$_.FolderPath -match $StrGroup} 
$Countin = ForEach-Object {[array] ($SiloServers | ? {$_.FolderPath -match $StrGroup} | where {$_.LogonMode -match "Allow*"} ) } | Measure | % { $_.Count}
$Countout = ForEach-Object {[array] ($SiloServers | ? {$_.FolderPath -match $StrGroup} | where {$_.LogonMode -notmatch "Allow*"} ) } | Measure | % { $_.Count}


# Write Html
("Saving results to html report: " + $ResultsHTML) | LogMe 

writeHtmlHeader $WebPageTitle $ResultsHTML
writeTableHeader $ResultsHTML
$allResults | sort-object -property FolderPath | % { writeData $allResults $ResultsHTML }
writeHtmlFooter $ResultsHTML


# Copying Html to Web Server
#("Copying $ResultsHTML to: " + $HTMLServer +"\" +$HTMLPage) | LogMe 


#Try { Copy-Item $ResultsHTML $HTMLServer\$HTMLPage }
#Catch
    #{ 

    #"Error Copying $ResultsHTML to $HTMLServer\$HTMLPage" | LogMe -error
    #$_.Exception.Message | LogMe -error
    
    #}

" " | LogMe -display

#"Checking Alerts" | LogMe -display
# Process for Errors
#FindErrors $allResults

"" | LogMe -display
"Script Completed" | LogMe -display

"" | LogMe -display
"Script Ended at $(get-date)" | LogMe -display

#$elapsed = GetElapsedTime $script:startTime
#"Total Elapsed Script Time: $elapsed" | LogMe -display


#Script Cleanup
$allResults = $null
$ErrorsandWarnings = $null
$script:EchoAlerts = $null
$script:EchoErrors = $null
$tests = $null
Get-PSSession | Remove-PSSession

$emailbody = Get-Content '.\US Desktop Remote Status.html' -Raw
Send-MailMessage -to $emailTo -From $emailFrom -Subject $emailSubject -BodyAsHtml $emailbody -SmtpServer Smtp-pdc.domain.com
Start-Sleep -s 5

#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# END - Main Program
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------