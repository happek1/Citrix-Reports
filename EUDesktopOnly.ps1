#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# Script Name: USDesktopOnly.ps1
# Modified: 7/22/2016
#
# Script Requirements: Citrix.XenApp.Commands
# Description: This script gets Health Status of XenApp Farm Remote Servers and Desktop Silo Infomation
#              Based on script by Jason Poyner's (http://deptive.co.nz/xenapp-farm-health-report)
#              and Stan Czerno (http://www.czerno.com/blog/post/2014/06/12/powershell-script-to-monitor-a-citrix-xenapp-farm-s-health)
#              
# Modified by khappe to only include only relevant information, script creates a PS Session to a XenApp Admin server then performs PS commands
# Note: Remote PS Session must be enabled on the Admin Server of script will not work
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# User Definable Variables
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


# Email Settings
# Multiple email addresses example: "email@domain.com,email2@domain.com"
$emailFrom     = "noc@domain.com"
$emailTo       = "noc@domain.com"
$emailCC       = "@domain.com" 
$smtpServer    = "Smtp-pdc.domain.com"
$SendAlerts    = $false

$SendEmailWarnings = $false
$SendEmailErrors = $true

##String Groups##
$StrGroup = "Desktop - Prod"
$excludedFolders = @("")

# Optional: Excluded folders from health check. 
# Example: @("Servers/Application", "Servers/Std Instances")
$excludedFolders = @("")

# Server to be excluded with any particular name or name filter. 
# Example: @("SRV1","SRV2")
$ServerFilter = @("") 

# The maximum uptime days a server can report green. 
$maxUpTimeDays = 2

#Admin Server for PSSESSSION
$AdminServer = "Server"

#Noc Admin Server for Email
$nocadmin = "Server"

# License Type to be defined 
# Example: @("MPS_PLT_CCU", "MPS_ENT_CCU", "XDT_ENT_UD") 
#$LicenseTypes = @("MPS_ENT_CCU")
$LicenseTypes = @("MPS_PLT_CCU", "MPS_ENT_CCU", "XDT_ENT_UD")

# Alert Spam Protection Timeframe, in seconds
$EmailAlertsLagTime = "1800"

# Webserver Page
$HTMLPage   = "EU Desktop Only Report.html"


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
$emailSubject  = "EU Desktop Only Report" 

# Log files created in the location of script. 
$LogFile = Join-Path $currentDir ("EU Desktop Only Report.log")
$PreviousLogFile = Join-Path $currentDir ("EU Desktop Only Report_PreviousRun.log")
$ResultsHTML = Join-Path $currentDir ("EU Desktop Only Report.html")
$AlertsFile = Join-Path $currentDir ("EU Desktop Only Report_Alerts.log")
$PreviousAlertsFile = Join-Path $currentDir ("EU Desktop Only Report_PreviousAlerts.log")
$AlertsEmailFile = Join-Path $currentDir ("EU Desktop Only Report_Email.log")

# Table headers
$headerNames  = "Logons", "UpTimeDays", "ServerLoad","ActiveUsers", "DiscUsers"
$headerWidths =    "6",      "8",           "8",            "6",         "6"

# Cell Colors
$ErrorStyle = "style=""background-color: #000000; color: #FF3300;"""
$WarningStyle = "style=""background-color: #000000;color: #FFFF00;"""


# The variable to count the server names
[int]$TotalServers = 0; $TotalServersCount = 0

$allResults = @{}
$allResults2 = @{}


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
	</table>
        
        <table width='100%'>
        <tr bgcolor='#CCCCCC'>
        <td width=33% align='center' valign="middle">
        <font face='Tahoma' color='#8A0808' size='2'><strong> Desktop Silo Information</strong></font>
        </td>
        </tr>
        </table>
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


#"Checking Citrix License usage on $LicenseServer" | LogMe -Display 
#$LicenseReport = CheckLicense


" " | LogMe; "Checking Citrix XenApp Server Health." | LogMe ; " " | LogMe


Get-XAServer | where-object {$_.FolderPath -match $StrGroup} | Sort-Object ServerName | % { 

    $tests = @{}	
    
    # Check If Server is in Excluded Folder path or server list (remove the # to turn on the filter)
    
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
        If ($totalDiscServerSessions -gt $totalActiveServerSessions) { $tests.DiscUsers = "WARNING", $totalDiscServerSessions }

            	
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
			Else { "Serverload is High [ $CurrentServerload ]" |  -error; $tests.Serverload = "ERROR", ($CurrentServerload.load) }

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


# Write Html
("Saving results to html report: " + $ResultsHTML) | LogMe 

writeHtmlHeader $WebPageTitle $ResultsHTML
writeTableHeader $ResultsHTML
$allResults | sort-object -property FolderPath | % { writeData $allResults $ResultsHTML }
writeHtmlFooter $ResultsHTML


"" | LogMe -display
"Script Completed" | LogMe -display

"" | LogMe -display
"Script Ended at $(get-date)" | LogMe -display


#Script Cleanup
$allResults = $null
$allResults2 = $null
$ErrorsandWarnings = $null
$script:EchoAlerts = $null
$script:EchoErrors = $null
$tests = $null
$tests2 = $null

#Email the report 
#$emailbody = Get-Content "D:\xenapp\EU Desktop Only Report_Alerts.html"
#EU Desktop Only Report_Alerts.html"' -Raw
#$currentDir ("EU Desktop Only Report_Alerts.html")

#$N = New-PSSession -ComputerName $nocadmin
#Invoke-Command -Session $N {Send-MailMessage -to $($args[0]) -CC $($args[1]) -From $($args[2]) -Subject $($args[3])  -BodyAsHtml $($args[4]) -SmtpServer $($args[5]) } -ArgumentList $emailTo,$emailCC,$emailFrom,$emailSubject,$emailbody,$smtpServer
#Start-Sleep -s 1

#Get-PSSession | Remove-PSSession
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# END - Main Program
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------