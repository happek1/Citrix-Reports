#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# Script Name: XenApp_Sites_Status.ps1
# Modified: 30/01/2015
# Created By: Aaron Argent
#
# Script Requirements: Citrix.XenApp.Commands
# Description: This script gets Health Status of XenApp Farm Servers 
#              Based on script by Jason Poyner's (http://deptive.co.nz/xenapp-farm-health-report)
#              and Stan Czerno (http://www.czerno.com/blog/post/2014/06/12/powershell-script-to-monitor-a-citrix-xenapp-farm-s-health)
#              Displays WebPage for Different Citrix Zones and Sends Email Notification
#              Run Script eg. "Powershell.exe -ExecutionPolicy Bypass -File D:\Scripts\XenAppStatus\XenApp_Sites_Status65.ps1 ADE"
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


Param (
   [string[]]$LocationName
)

switch ($LocationName) 
    { 

        'ADE' {

            $StrGroup = "ADE"
            $excludedFolders = @("Servers/AKL","Servers/BNE","Servers/MEL","Servers/SYD","Servers")

            }
        'AKL' {

            $StrGroup = "AKL"
            $excludedFolders = @("Servers/ADE","Servers/BNE","Servers/MEL","Servers/SYD","Servers")

            }
        'BNE' {

            $StrGroup = "BNE"
            $excludedFolders = @("Servers/ADE","Servers/AKL","Servers/MEL","Servers/SYD","Servers")

            }
        'MEL' {

            $StrGroup = "MEL"
            $excludedFolders = @("Servers/ADE","Servers/AKL","Servers/BNE","Servers/SYD","Servers")

            }
        'SYD' {

            $StrGroup = "SYD"
            $excludedFolders = @("Servers/ADE","Servers/AKL","Servers/BNE","Servers/MEL","Servers")

            }
        default {

            $StrGroup = "BNE"
            $excludedFolders = @("Servers/ADE","Servers/AKL","Servers/MEL","Servers/SYD","Servers")

            }
    }

Clear-host


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# User Definable Variables
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------


# Email Settings
# Multiple email addresses example: "email@domain.com,email2@domain.com"
$emailFrom     = ""
$emailTo       = ""
$emailCC       = "" 
$smtpServer    = ""
$SendAlerts    = $false

$SendEmailWarnings = $false
$SendEmailErrors = $true

# Optional: Define the Load Evaluator to be checked in all the servers. 
# Example: @("Default","Advanced"). 
$defaultLE = @("")

# Relative path to the PVS vDisk write cache file
$PvsWriteCache   = "d$\.vdiskcache"
$PvsWriteCache2  = "d$\vdiskdif.vhdx"

# Optional: Excluded folders from health check. 
# Example: @("Servers/Application", "Servers/Std Instances")
#$excludedFolders = @("Servers")

# Server to be excluded with any particular name or name filter. 
# Example: @("SRV1","SRV2")
$ServerFilter = @("")

# The maximum uptime days a server can report green. 
$maxUpTimeDays = 10

# Ports used for the functionality check.
$RDPPort = "3389"
$ICAPort = "1494"
$sessionReliabilityPort = "2598"

# Test Session reliability Port?
$TestSessionReliabilityPort = "true"

# License Server Name for license usage.
$LicenseServer = "SERVERNAME"

# License Type to be defined 
# Example: @("MPS_PLT_CCU", "MPS_ENT_CCU", "XDT_ENT_UD") 
#$LicenseTypes = @("MPS_ENT_CCU")
$LicenseTypes = @("XDT_ENT_UD")

# Alert Spam Protection Timeframe, in seconds
$EmailAlertsLagTime = "1800"

# Image Name for Webpage Logo. Store in Script Folder
$ImageName1 = "logo.jpg" 

# Webserver Shared Folder Location
$HTMLServer = "\\\c$\inetpub\wwwroot"
$HTMLPage   = "XenApp65_" +$StrGroup +".html"

# Webpage URL
$TopLevelURL = "http://bnevppvs1/"
$MonitorURL = $TopLevelURL + $HTMLPage

# Location of PVS Cache CSV File
$CSVFile = $HTMLServer +"\PVSCacheCSV.csv"

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

# Define the Global variable and Check for the Citrix Snapin
If ((Get-PSSnapin "Citrix.*" -EA silentlycontinue) -eq $null)
    {
	
    Try { Add-PSSnapin Citrix.* -ErrorAction Stop }
	Catch { Write-Error "Error loading XenApp Powershell snapin"; Return } 
    
    }

# Get farm details once to use throughout the script
$FarmDetails = Get-XAFarm 
$CitrixFarmName = $FarmDetails.FarmName
$WebPageTitle = "$CitrixFarmName Health Status"

# Email Subject with the farm name
$emailSubject  = "Citrix Health Alert: " + $CitrixFarmName + "Farm - Site: " + $StrGroup 

# Log files created in the location of script. 
$LogFile = Join-Path $currentDir ("XenAppStatus_"+$StrGroup+".log")
$PreviousLogFile = Join-Path $currentDir ("XenAppStatus_"+$StrGroup+"_PreviousRun.log")
$ResultsHTML = Join-Path $currentDir ("XenAppStatus_"+$StrGroup+".html")
$AlertsFile = Join-Path $currentDir ("XenAppStatus_"+$StrGroup+"_Alerts.log")
$PreviousAlertsFile = Join-Path $currentDir ("XenAppStatus_"+$StrGroup+"_PreviousAlerts.log")
$AlertsEmailFile = Join-Path $currentDir ("XenAppStatus_"+$StrGroup+"_Email.log")

# Table headers
$headerNames  = "Ping", "Logons", "ActiveUsers", "DiscUsers", "LoadEvaluator", "ServerLoad", "RDPPort", "ICAPort", "SRPort", "AvgCPU", "MemUsg", "ContextSwitch", "DiskFree", "IMA", "CitrixPrint", "Spooler", "WMI", "RPC", "UptimeDays", "RAMCache", "vDiskCache", "vDiskImage"
$headerWidths = "6",      "6",      "6",           "6",         "6",          "6",          "4",       "4",     "5",      "5",      "5",   "5",  "6",  "6",    "6",           "8",       "8",   "6",   "8",   "8",   "8",   "8"

# Cell Colors
$ErrorStyle = "style=""background-color: #000000; color: #FF3300;"""
$WarningStyle = "style=""background-color: #000000;color: #FFFF00;"""

# Farm Average
$CPUAverage = 0 
$RAMAverage = 0 

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


Function CheckService() 
    {

	Param ($ServiceName)
	
    Try 
	    {

        If (!(Get-Service -Display $ServiceName -ComputerName $server -ErrorAction Stop) ) { "$ServiceName is not available..." | LogMe -display; $ServiceResult = "N/A" }
    	Else
            {
        	
            If ((Get-Service -Display $ServiceName -ComputerName $server -ErrorAction Stop).Status -Match "Running") { "$ServiceName is running" | LogMe -Display; $ServiceResult = "Success" }
        	Else
                {
                
                "$ServiceName is not running"  | LogMe -error
				
                Try
                    {
                    
                    Start-Service -InputObject $(Get-Service -Display $ServiceName -ComputerName $server) -ErrorAction Stop 
					"Start command sent for $ServiceName"  | LogMe -warning
            		Start-Sleep 5 # Sleep a bit to allow service to start
					If ((Get-Service -Display $ServiceName -ComputerName $server).Status -Match "Running") { "$ServiceName is now running" | LogMe -Display; $ServiceResult = "Success" }
					Else {	"$ServiceName failed to start."  | LogMe -error	$ServiceResult = "Error" }

                    }
                Catch { "Start command failed for $ServiceName. You need to check the server." | LogMe -Error; return "Error" } 
                
                } 
            }

    	return $ServiceResult

	    } 
    Catch { "Error while checking the Service. Server may be down or has a permission issue." | LogMe -error; return "Error" } 

    } #End Function: CheckService


Function CheckCpuUsage() 
    {
	
    Param ($hostname)

	Try 
        {
        
        $CpuUsage = (get-counter -ComputerName $hostname -Counter "\Processor(_Total)\% Processor Time" -SampleInterval 1 -MaxSamples 5 -ErrorAction Stop | select -ExpandProperty countersamples | select -ExpandProperty cookedvalue | Measure-Object -Average).average
    	$CpuUsage = "{0:N1}" -f $CpuUsage; return $CpuUsage

        }
    Catch { "Error returned while checking the CPU usage. Perfmon Counters may be at fault." | LogMe -error; return 101 } 

    } #End Function: CheckCpuUsage


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
    
 
Function CheckMemoryUsage() 
    { 
	
    Param ($hostname)
    
    Try 
	    {
        
        $SystemInfo = (Get-WmiObject -Computername $hostname -Class Win32_OperatingSystem -ErrorAction Stop | Select-Object TotalVisibleMemorySize, FreePhysicalMemory)
    	$TotalRAM = $SystemInfo.TotalVisibleMemorySize/1MB 
    	$FreeRAM = $SystemInfo.FreePhysicalMemory/1MB 
    	$UsedRAM = $TotalRAM - $FreeRAM 
    	$RAMPercentUsed = ($UsedRAM / $TotalRAM) * 100 
    	$RAMPercentUsed = "{0:N2}" -f $RAMPercentUsed
    	
        return $RAMPercentUsed

        }
    Catch { "Error returned while checking the Memory usage. Perfmon Counters may be at fault" | LogMe -error; return 101 } 
    
    } #End Function: CheckMemoryUsage


Function Ping ([string]$hostname, [int]$timeout) 
    {

    $ping = new-object System.Net.NetworkInformation.Ping #creates a ping object
	
    Try { $result = $ping.send($hostname, $timeout).Status.ToString() }
    Catch { $result = "Failed" }
	
    return $result
    
    } #End Function: Ping


Function Check-Port() 
    { 

	Param ([string]$hostname, [string]$port)

    Try { $socket = new-object System.Net.Sockets.TcpClient($hostname, $Port) -ErrorAction Stop } #creates a socket connection to see if the port is open
    Catch { $socket = $null; "Socket connection on $port failed" | LogMe -error; return $false }

	If ($socket -ne $null) 
        { 
        
        "Socket Connection on $port Successful" | LogMe -Display

        If ($port -eq "1494") 
            {
        
            $stream   = $socket.GetStream() #gets the output of the response
		    $buffer   = new-object System.Byte[] 1024
		    $encoding = new-object System.Text.AsciiEncoding

		    Start-Sleep -Milliseconds 500 #records data for half a second			
		
            While ($stream.DataAvailable) 
                {

                $read     = $stream.Read($buffer, 0, 1024)  
			    $response = $encoding.GetString($buffer, 0, $read)
		
			    If ($response -like '*ICA*'){ "ICA protocol responded" | LogMe -Display; return $true } 

			    }

			"ICA did not respond correctly" | LogMe -error; return $false 
    
            }
        Else { return $true }

        }
    Else { "Socket connection on $port failed" | LogMe -error; return $false }

    } #End Function: Check-Port


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
        font-family: Lao UI;
        font-size: 11px;
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
		font-size: 40px;
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
        <table class="header" width='100%'>
        <tr bgcolor='#FAFAFA'>
        <td style="text-align: center; text-shadow: 2px 2px 2px #ff0000;">
        <img src="$ImageName1">
        </td>
        <td class="header" width='826' align='center' valign="middle" style="background-image: url('ekg_wide.jpg'); background-repeat: no-repeat; background-position: center; ">
        <p class="shadow"> $CitrixFarmName<br>Health Status </p>
        </tr>
        </table>
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
        <font face='Tahoma' color='#003399' size='2'><strong>Active Sessions:  $TotalActiveSessions [ ICA $TotalICAActiveSessions | RDP $TotalRDPActiveSessions ]</strong></font>
        <td width=50% align='center' valign="middle">
        <font face='tahoma' color='#003399' size='2'><strong>Disconnected Sessions:  $TotalDisconnectedSessions [ ICA $TotalICADisconnectedSessions | RDP $TotalRDPDisconnectedSessions ]</strong></font>
        </td>
        </tr>
        </table>
        <table width='100%'> 
        <tr bgcolor='#CCCCCC'> 
        <td width=50% align='center' valign="middle"> 
        <font face='Tahoma' color='#003399' size='2'><strong>Average CPU:  $CPUAverage %</strong></font> 
        <td width=50% align='center' valign="middle"> 
        <font face='tahoma' color='#003399' size='2'><strong>Average RAM:  $RAMAverage %</strong></font> 
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
 

Function writeHtmlFooter
    { 
	
    Param ($fileName)

    #$footer=("<br><font face='HP Simplified' color='#003399' size='2'><br><I>Report last updated on {0}.<br> Script Hosted on server {3}.<br>Script Path: {4}</font>" -f (Get-Date -displayhint date),$env:userdomain,$env:username,$env:COMPUTERNAME,$currentDir) 
	
    $footer = @"
    </table>
    </body>
    </html>
"@

    $footer | Out-File $FileName -append

    } #End Function: writeHtmlFooter


Function GetElapsedTime([datetime]$starttime) 
    {

    $runtime = $(get-date) - $starttime
    $retStr = [string]::format("{0} sec(s)", $runtime.TotalSeconds)
    $retStr

    } #End Function: GetElapsedTime 


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


$sessions = Get-XASession -Farm | where-object {$_.ServerName -match $StrGroup}
    

Get-XAServer | where-object {$_.ServerName -match $StrGroup} | Sort-Object ServerName | % { 

    $tests = @{}	
    
    # Check If Server is in Excluded Folder path or server list
    #If ($excludedFolders -contains $_.FolderPath) { $_.ServerName + " in excluded Server folder - skipping" | LogMe -Display; "" | LogMe; return }
    #If ($ServerFilter -contains $_.ServerName) { $_.ServerName + " is excluded in the Server List  - skipping" | LogMe -Display; "" | LogMe; return }

    [int]$TotalServers = [int]$TotalServers + 1; $server = $_.ServerName

    "Server Name: $server" | LogMe

    $ProcessTime = Get-Date -format R
    "Processing Time: $ProcessTime" | LogMe

#    $WorkerGroups = $null, (Get-XAWorkerGroup -ServerName $server | % {$_.WorkerGroupName})
    
#    If ($WorkerGroups -eq $null) 
#        {
#
#        "WorkerGroups: N/A "
#        $tests.WorkerGroups = $null, "N/A" 
#
#        } 
#    Else
#        {
#
#        "WorkerGroups:$WorkerGroups " 
#        $tests.WorkerGroups = $WorkerGroups 
#
#        }
   

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
        
#        If ($_.LogOnsEnabled -eq $false) { "Logons are disabled " | LogMe -error; $tests.Logons = "ERROR", "Disabled" } 
#        Else { "Logons are enabled " | LogMe -Display; $tests.Logons = "SUCCESS","Enabled" }


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


    	# Check Load Evaluator
	    $LEFlag = 0; $CurrentLE = ""
	    $CurrentLE = (Get-XALoadEvaluator -ServerName $server).LoadEvaluatorName

	    Foreach ($LElist in $defaultLE) 
            {

            If ($CurrentLE -match $LElist)
                {    
            
                "Default Load Evaluator assigned" | LogMe -display
                $tests.LoadEvaluator = "SUCCESS", $CurrentLE
                $LEFlag = 1; break 
                
                }

            }

        If ($LEFlag -eq 0 )
            {

            If ($CurrentLE -match "Offline") 
                {

		        "Server is in Offline LE; Please check the box" | LogMe -error
		        $tests.LoadEvaluator = "WARNING", $CurrentLE 
                
                }
            Else
                {

                "Non-default Load Evaluator assigned" | LogMe -warning
                $tests.LoadEvaluator = "WARNING", $CurrentLE 

                }
        
            }


        # Test RDP Port
        If (Check-Port $server $RDPPort) { $tests.RDPPort = "SUCCESS", "Success" }
        Else { $tests.RDPPort = "ERROR", "$vDiskCacheSize" }        


        # Test ICA Port
        If (Check-Port $server $ICAPort) { $tests.ICAPort = "SUCCESS", "Success" }
        Else { $tests.ICAPort = "ERROR","No Response" }


        # Test Session Reliability Port   
        If ($TestSessionReliabilityPort -eq "True")
            {
            
            If (Check-Port $server $sessionReliabilityPort) { $tests.SRPort = "SUCCESS", "Success" }
            Else { $tests.SRPort = "ERROR", "No Response" }

            }


        # Check Average CPU Value
        $AvgCPUval = CheckCpuUsage ($server)

        If ( [int] $AvgCPUval -lt 80) { "CPU usage is normal [ $AvgCPUval % ]" | LogMe -display; $tests.AvgCPU = "SUCCESS", ($AvgCPUval) }
		ElseIf ([int] $AvgCPUval -lt 90) { "CPU usage is medium [ $AvgCPUval % ]" | LogMe -warning; $tests.AvgCPU = "WARNING", ($AvgCPUval) }  	
		ElseIf ([int] $AvgCPUval -lt 95) { "CPU usage is high [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", ($AvgCPUval) }
		ElseIf ([int] $AvgCPUval -eq 101) { "CPU usage test failed" | LogMe -error; $tests.AvgCPU = "ERROR", "Err" }
        Else { "CPU usage is Critical [ $AvgCPUval % ]" | LogMe -error; $tests.AvgCPU = "ERROR", ($AvgCPUval) }  
        
        $CPUAverage =  $CPUAverage + $AvgCPUval 
 
        $AvgCPUval = 0
 

        # Check Memory Usage
        $UsedMemory = CheckMemoryUsage ($server)

        If ( [int] $UsedMemory -lt 80) { "Memory usage is normal [ $UsedMemory % ]" | LogMe -display; $tests.MemUsg = "SUCCESS", ($UsedMemory) }
		ElseIf ([int] $UsedMemory -lt 85) { "Memory usage is medium [ $UsedMemory % ]" | LogMe -warning; $tests.MemUsg = "WARNING", ($UsedMemory) }  	
		ElseIf ([int] $UsedMemory -lt 90) { "Memory usage is high [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", ($UsedMemory) }
		ElseIf ([int] $UsedMemory -eq 101) { "Memory usage test failed" | LogMe -error; $tests.MemUsg = "ERROR", "Err" }
        Else { "Memory usage is Critical [ $UsedMemory % ]" | LogMe -error; $tests.MemUsg = "ERROR", ($UsedMemory) }  

        $RAMAverage = $RAMAverage + $UsedMemory 

		$UsedMemory = 0 


        # Check Context Switching Value
        $ContextSwitchval = CheckContextSwitch ($server)

        If ( [int] $ContextSwitchval -lt 45000) { "Context Switches is normal [ $ContextSwitchval ]" | LogMe -display; $tests.ContextSwitch = "SUCCESS", ($ContextSwitchval) }
		ElseIf ([int] $ContextSwitchval -gt 45000) { "Context Switches is medium [ $ContextSwitchval ]" | LogMe -warning; $tests.ContextSwitch = "WARNING", ($ContextSwitchval) }  	
		ElseIf ([int] $ContextSwitchval -gt 60000) { "Context Switches is high [ $ContextSwitchval ]" | LogMe -error; $tests.ContextSwitch = "ERROR", ($ContextSwitchval) }
		ElseIf ([int] $ContextSwitchval -eq 0) { "Context Switches test failed" | LogMe -error; $tests.ContextSwitch = "ERROR", "Err" }
        Else { "Context Switches is Critical [ $ContextSwitchval % ]" | LogMe -error; $tests.ContextSwitch = "ERROR", ($ContextSwitchval) }  

        $ContextSwitchval = 0


        # Check Disk Usage
        $HardDisk = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='C:'" | Select-Object Size,FreeSpace

        $DiskTotalSize = $HardDisk.Size
        $DiskFreeSpace = $HardDisk.FreeSpace

        $PercentageDS = (($DiskFreeSpace / $DiskTotalSize ) * 100); $PercentageDS = "{0:N2}" -f $PercentageDS

        If ( [int] $PercentageDS -gt 15) { "Disk Free is normal [ $PercentageDS % ]" | LogMe -display; $tests.DiskFree = "SUCCESS", ($PercentageDS) }
		ElseIf ([int] $PercentageDS -lt 15) { "Disk Free is Low [ $PercentageDS % ]" | LogMe -warning; $tests.DiskFree = "WARNING", ($PercentageDS) }  	
		ElseIf ([int] $PercentageDS -lt 5) { "Disk Free is Critical [ $PercentageDS % ]" | LogMe -error; $tests.DiskFree = "ERROR", ($PercentageDS) }
		ElseIf ([int] $PercentageDS -eq 0) { "Disk Free test failed" | LogMe -error; $tests.DiskFree = "ERROR", "Err" }
        Else { "Disk Free is Critical [ $PercentageDS % ]" | LogMe -error; $tests.DiskFree = "ERROR", ($PercentageDS) }  
		
        $PercentageDS = 0


        # Check Services
        $ServiceOP = CheckService ("Citrix Independent Management Architecture")
        If ($ServiceOP -eq "Error")  { $tests.IMA = "ERROR", $ServiceOP }
        Else { $tests.IMA = "SUCCESS", $ServiceOP }

        $ServiceOP = CheckService ("Print Spooler")
        If ($ServiceOP -eq "Error")  { $tests.Spooler = "ERROR", $ServiceOP }
        Else { $tests.Spooler = "SUCCESS", $ServiceOP }

        $ServiceOP = CheckService ("Citrix Print Manager Service")
        If ($ServiceOP -eq "Error")  { $tests.CitrixPrint = "ERROR", $ServiceOP }
        Else { $tests.CitrixPrint = "SUCCESS", $ServiceOP }
		

        # Check Server Load
	    If ($tests.IMA[0] -eq "Success")
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
        Else { "Server load can't be determine since IMA failed " | LogMe -error; $tests.Serverload = "ERROR", "IMA Err" }

        $CurrentServerLoad = 0

	
        # Test WMI
        $tests.WMI = "ERROR","Error"
	    Try { $wmi=Get-WmiObject -class Win32_OperatingSystem -computer $_.ServerName } 
	    Catch {	$wmi = $null }
		
        # Perform WMI related checks
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


        # Check PVS
        If (Test-Path \\$Server\c$\Personality.ini)
            {

            $RAMCacheSize = 0
            $vDiskCacheSize = 0
        
            $PvsWriteCacheUNC = Join-Path "\\$Server" $PvsWriteCache 
            $vDiskexists  = Test-Path $PvsWriteCacheUNC

            If ($vDiskexists -eq $False) 
                {
            
                $PvsWriteCacheUNC = Join-Path "\\$Server" $PvsWriteCache2
                $vDiskexists = Test-Path $PvsWriteCacheUNC
            
                $CacheSize = Import-CSV $CSVFile | Where {$_.Server -eq $server} 
            
                If ($CacheSize.CacheSize -ine "")
                    { 
                 
                    $RAMCacheSize = $CacheSize.CacheSize
                
                    }

	    		}

		    If ($vDiskexists -eq $True)
			    {

                $CacheDisk = [long] ((get-childitem $PvsWriteCacheUNC -force).length)
                $HardDiskD = Get-WmiObject Win32_LogicalDisk -ComputerName $server -Filter "DeviceID='D:'" | Select-Object FreeSpace
                $DiskDFreeSpace = $HardDiskD.FreeSpace
                $PercentageDS = (($CacheDisk / $DiskDFreeSpace ) * 100); $PercentageDS = "{0:N2}" -f $PercentageDS
                $vDiskCacheSize = [math]::Round($PercentageDS)
                
			    }              

            If ($RAMCacheSize -ine 0)
                {

                If ( [int] $RAMCacheSize -lt 80) { "PVS RAM Cache usage is normal [ $RAMCacheSize % ]" | LogMe -display; $tests.RAMCache = "SUCCESS", ($RAMCacheSize) }
                ElseIf ([int] $RAMCacheSize -lt 90) { "PVS RAM Cache usage is medium [ $RAMCacheSize % ]" | LogMe -warning; $tests.RAMCache = "WARNING", ($RAMCacheSize) }  	
		        ElseIf ([int] $RAMCacheSize -lt 95) { "PVS RAM Cache usage is high [ $RAMCacheSize % ]" | LogMe -error; $tests.RAMCache = "ERROR", ($RAMCacheSize) }
		        ElseIf ([int] $RAMCacheSize -eq 101) { "PVS RAM Cache usage test failed" | LogMe -error; $tests.RAMCache = "ERROR", "Err" }
                Else { "PVS RAM Cache usage is Critical [ $RAMCacheSize % ]" | LogMe -error; $tests.RAMCache = "ERROR", ($RAMCacheSize) }  

                }
            Else { $tests.RAMCache = "SUCCESS", "N/A" }
                
            If ($vDiskCacheSize -ine 0)
                {

                If ( [int] $vDiskCacheSize -lt 80) { "PVS vDisk Cache usage is normal [ $vDiskCacheSize % ]" | LogMe -display; $tests.vDiskCache = "SUCCESS", ($vDiskCacheSize) }
                ElseIf ([int] $vDiskCacheSize -lt 90) { "PVS vDisk Cache usage is medium [ $vDiskCacheSize % ]" | LogMe -warning; $tests.vDiskCache = "WARNING", ($vDiskCacheSize) }  	
		        ElseIf ([int] $vDiskCacheSize -lt 95) { "PVS vDisk Cache usage is high [ $vDiskCacheSize % ]" | LogMe -error; $tests.vDiskCache = "ERROR", ($vDiskCacheSize) }
		        ElseIf ([int] $vDiskCacheSize -eq 101) { "PVS vDisk Cache usage test failed" | LogMe -error; $tests.vDiskCache = "ERROR", "Err" }
                Else { "PVS vDisk Cache usage is Critical [ $vDiskCacheSize % ]" | LogMe -error; $tests.vDiskCache = "ERROR", ($vDiskCacheSize) }  

                }        
            Else { $tests.vDiskCache = "SUCCESS", "N/A" }

		    $VDISKImage = get-content \\$Server\c$\Personality.ini | Select-String "Diskname" | Out-String | % { $_.substring(12)}
            If ($VDISKImage -ine "")
                {
            
			    "vDisk Detected: $VDISKImage" | LogMe -display; $tests.vDiskImage = "SUCCESS", ($VDISKImage)
			    #$tests.vDiskImage = "SUCCESS", $VDISKImage
            
                }
            Else 
                {
            
			    "vDisk Unknown"  | LogMe -error; $tests.vDiskImage = "ERROR", "Err"
			    #$tests.vDiskImage = "WARNING", $VDISKImage
            
			    }   
            
		    }
	    Else { $tests.RAMCache = "SUCCESS", "N/A";  $tests.vDiskCache = "SUCCESS", "N/A";  $tests.vDiskImage = "SUCCESS", "N/A" }


        } 


	$allResults.$server = $tests

    " " | LogMe -display


    }


# DISPLAY TOTALS
$TotalServersCount = [int]$TotalServers

$ActiveICAUsers = [array]($sessions | ? {$_.State -eq "Active" -and $_.Protocol -eq "Ica"})
$ActiveRDPUsers = [array]($sessions | ? {$_.State -eq "Active" -and $_.Protocol -eq "Rdp"})
$DiscICAUsers = [array]($sessions | ? {$_.State -eq "Disconnected" -and $_.Protocol -eq "Ica"})
$DiscRDPUsers = [array]($sessions | ? {$_.State -eq "Disconnected" -and $_.Protocol -eq "Rdp"})
#$ActiveUsers = [array]($sessions | ? {$_.State -eq "Active" -and $_.Protocol -ne "Console"})
#$DiscUsers = [array]($sessions | ? {$_.State -eq "Disconnected" -and $_.Protocol -ne "Console"})

#If ($ActiveUsers) { $TotalActiveSessions = $ActiveUsers.count } Else { $TotalActiveSessions = 0 }
#If ($DiscUsers) { $TotalDisconnectedSessions = $DiscUsers.count } Else { $TotalDisconnectedSessions = 0 }
If ($ActiveICAUsers) { $TotalICAActiveSessions = $ActiveICAUsers.count } Else { $TotalICAActiveSessions = 0 }
If ($DiscICAUsers) { $TotalICADisconnectedSessions = $DiscICAUsers.count } Else { $TotalICADisconnectedSessions = 0 }
If ($ActiveICAUsers) { $TotalRDPActiveSessions = $ActiveRDPUsers.count } Else { $TotalRDPActiveSessions = 0 }
If ($DiscRDPUsers) { $TotalRDPDisconnectedSessions = $DiscRDPUsers.count } Else { $TotalRDPDisconnectedSessions = 0 }

$TotalActiveSessions = $TotalICAActiveSessions+$TotalRDPActiveSessions
$TotalDisconnectedSessions = $TotalICADisconnectedSessions+$TotalRDPDisconnectedSessions

"Total Number of Servers: $TotalServersCount" | LogMe 
"Total Active Applications: $TotalActiveSessions" | LogMe 
"Total Disconnected Sessions: $TotalDisconnectedSessions" | LogMe  

$CPUAverage = "{0:N2}" -f ($CPUAverage / $TotalServersCount) 
$RAMAverage = "{0:N2}" -f ($RAMAverage / $TotalServersCount) 

" " | LogMe -display


# Write Html
("Saving results to html report: " + $ResultsHTML) | LogMe 

writeHtmlHeader $WebPageTitle $ResultsHTML
writeTableHeader $ResultsHTML
$allResults | sort-object -property FolderPath | % { writeData $allResults $ResultsHTML }
writeHtmlFooter $ResultsHTML


# Copying Html to Web Server
("Copying $ResultsHTML to: " + $HTMLServer +"\" +$HTMLPage) | LogMe 

Try { Copy-Item $ResultsHTML $HTMLServer\$HTMLPage }
Catch
    { 

    "Error Copying $ResultsHTML to $HTMLServer\$HTMLPage" | LogMe -error
    $_.Exception.Message | LogMe -error
    
    }

" " | LogMe -display

"Checking Alerts" | LogMe -display
# Process for Errors
FindErrors $allResults


#Checking Email Send Timeframe
[int]$hour = get-date -format HH
If($hour -lt $int_Email_Start -or $hour -gt $int_Email_End){ $EmailSendAllowed = $False; "Email Send Disabled: Outside Timeframe" | LogMe -display }
Else { $EmailSendAllowed = $True; "Email Send Allowed" | LogMe -display }


If ((Test-Path $AlertsEmailFile) -And ($SendAlerts -eq $True) -And ($EmailSendAllowed -eq $True))
    {

    $EmailAlertWarnings = ""
    $EmailAlertErrors = ""

    If (((Import-Csv $AlertsEmailFile | Measure-Object | Select-Object).Count) -ge 1)
        {

        $EmailAlertErrorsCount = 0
        $EmailAlertWarningsCount = 0
        $BODY = ""

        Import-CSV $AlertsEmailFile | Sort-Object Type, Server | Foreach-Object{

            $strAlertServer = ""
            $strAlertComponent = ""
            $strAlertValue = ""
            $strAlertStatus = ""

            $strAlertServer = $_.Server
            $strAlertComponent = $_.Component
            $strAlertValue = $_.Value
            $strAlertStatus = $_.Status

            $strAlertOutput = 'Server: <strong>' +$strAlertServer +'</strong><br>Component: <strong>' +$strAlertComponent +'</strong><br>Value: <strong>' +$strAlertValue +'</strong><br>Status: <strong>' +$strAlertStatus +'</strong></p>'

            If (($_.Type -ieq "ERROR") -And ($SendEmailErrors))
                {

                $EmailAlertErrors += '<p' +$WarningStyle +'>' +$strAlertOutput
                $EmailAlertErrorsCount ++

                }

            If (($_.Type -ieq "WARNING") -And ($SendEmailWarnings))
                {

                $EmailAlertWarnings += '<p' +$ErrorStyle +'>' +$strAlertOutput
                $EmailAlertWarningsCount ++
                            
                }

            }

        If ($EmailAlertErrorsCount -gt 0)
            {

            $BODY = $BODY +"<p><strong>ERRORS Detected:</strong></p>"

            $BODY = $BODY +$EmailAlertErrors

            }

        If ($EmailAlertWarningsCount -gt 0)
            {

            $BODY = $BODY +"<p><strong>WARNINGS Detected:</strong></p>"

            $BODY = $BODY +$EmailAlertWarnings

            }

        $BODY = $BODY +"Check " +$MonitorURL +" for more information"

        If ((!$emailFrom) -Or (!$emailTo) -Or (!$smtpServer))
            {

            If (!$emailFrom) { $MailFlag = $True; Write-Warning "From Email is NULL" | LogMe -error }
            If (!$emailTo) { $MailFlag = $True; Write-Warning "To Email is NULL" | LogMe -error }
            If (!$smtpServer) { $MailFlag = $True; Write-Warning "SMTP Server is NULL" | LogMe -error }

            "Email could not sent. Please Check Configuration and Try again." | LogMe -error

            }
        Else
            {

            If (($EmailAlertWarningsCount -gt 0) -Or ($EmailAlertErrorsCount -gt 0))
                {

                $msg = new-object System.Net.Mail.MailMessage
                $msg.From=$emailFrom
                $msg.to.Add($emailTo)
            
                If ($emailCC) { $msg.cc.add($emailCC) }
            
                $msg.Subject=$emailSubject
                $msg.IsBodyHtml=$true
                $msg.Body=$BODY
                #$msg.Attachments.Add($LogFile)

                $smtp = new-object System.Net.Mail.SmtpClient
                $smtp.host=$smtpServer
            
                Try { $smtp.Send($msg); "" | LogMe -display; "Email Sent" | LogMe -display }
                Catch
                    { 
                
                    "" | LogMe -display
                    "Error Sending Email. See Error Messages Below:" | LogMe -error 
                    $ErrorMessage = $_.Exception.Message
                    $FailedItem = $_.Exception.ItemName
                    $ErrorMessage | LogMe -error
                    $FailedItem | LogMe -error

                    }

                Start-Sleep 2

                $msg.Dispose()

                $EmailAlertErrors = ""
                $EmailAlertWarnings = ""
                $BODY = ""

                }
            Else { "No Alerts to Email" | LogMe -display }

            }

        }

    }


"" | LogMe -display
"Script Completed" | LogMe -display

"" | LogMe -display
"Script Ended at $(get-date)" | LogMe -display

$elapsed = GetElapsedTime $script:startTime
"Total Elapsed Script Time: $elapsed" | LogMe -display


#Script Cleanup
$allResults = $null
$ErrorsandWarnings = $null
$script:EchoAlerts = $null
$script:EchoErrors = $null
$tests = $null


#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------
# END - Main Program
#-------------------------------------------------------------------------------
#-------------------------------------------------------------------------------