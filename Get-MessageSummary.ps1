<#
    .DESCRIPTION
    Get-MessageSummary.ps1
    Matt Krause 2017

    This script creates a HTML report containing the number of sent and received messages per users of a specific domain 
    The script report contains the following:

        * Date Range the report covers
        * User Name
        * Number of messages sent
        * Number of messages received
        * Date the report was run
        
    The script searches all Tranport logs in the environment to pull report information. The script detects which version of Exchange it's being run on and will
    automatically update the command to work with the version of Exchange

    .PARAMETER Domain
    Desired domain for the report data
    This is a required parameter. You will be prompted to enter the domain if you exclude it

    -Domain "domain.com"

    .PARAMETER HTMLReport
    Filename to write HTML Report to
    If no report location paramter is given, the file will be saved to the script location with a file name of "MessageSummary122017121520.html"

    -HTMLReport "C:\Report\reportname.html"

    .PARAMETER Start
    Start date for report data
    If no start date is provided, the script will collect data from the first of the current month

    -StartDate "1/1/2016"
    
    .PARAMETER End
    End date for report data
    If no end date is provided, the script will collect data up to the current date

    -EndDate "12/31/2016"
    
    .EXAMPLE
    Generate the HTML report (usind default dates and location):
    .\Get-MessageSummary.ps1 -Domain "domain.com"

    Generate the HTML report (with location and start/end dates):
      .\Get-MessageSummary.ps1 -Domain "domain.com" -HTMLReport "C:\Temp\Report.html" -StartDate "1/1/2016" -EndDate "1/31/2016"
    #>

#Read in parameters
param(
    [parameter(Position=0,Mandatory=$true,HelpMessage="Insert Domain")][string]$Domain,
    [parameter(Position=1,Mandatory=$false,HelpMessage="HTML Report Location")][string]$HTMLReport="1",
    [parameter(position=2,Mandatory=$false,HelpMEssage="Start Date")][string]$StartDate="1",
    [parameter(position=3,Mandatory=$false,HelpMessage="End Date")][string]$EndDate="1"
    )

#Set date and report file Variables based on provided parameters
$ReportDate = Get-Date
$Month = Get-Date -Format "MM"
$Year = Get-Date -Format "yyyy"

if ($HTMLReport -eq "1")
    {
        if(!$PSScriptRoot) #Make $PSSscriptRoot work with PowerShell v2
            {
                $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
                $ReportLocation = $PSScriptRoot +"\"
                $HTMLReportName = "MessageSummary" + $ReportDate.Month + $ReportDate.Day + $ReportDate.Year + $ReportDate.Hour + $ReportDate.Minute + $ReportDate.Second + ".html"
                $HTMLReport = $ReportLocation+$HTMLReportName
            }
        else
            {
                $ReportLocation = $PSScriptRoot +"\"
                $HTMLReportName = "MessageSummary" + $ReportDate.Month + $ReportDate.Day + $ReportDate.Year + $ReportDate.Hour + $ReportDate.Minute + $ReportDate.Second + ".html"
                $HTMLReport = $ReportLocation+$HTMLReportName 
            }
    }

if ($StartDate -eq "1")
    {
        $StartDate = $Month+"/01/"+$Year
    }

if ($EndDate -eq "1")
    {
        $EndDate = Get-Date -Format "MM/dd/yyyy"
    }

#Set up Array to hold report data
$Global:MailUsers = @()

#------------------------------------------------------Functions------------------------------------------------------
#Function to lookup info for Exchange 2010
Function Run2010
    {
        Param ($DomainUsers,$StartDate,$EndDate)
        ForEach($x in $DomainUsers)
        {
            $ReceiveCount = 0
            $SendCount = 0
            Write-Host "Looking up user"$x.Name -ForegroundColor Yellow
            Get-TransportServer | get-MessageTrackingLog -ResultSize Unlimited -Start "$StartDate" -End "$EndDate" -Sender $x.PrimarySMTPAddress -EventID RECEIVE | ? {$_.Source -eq "STOREDRIVER"}| ForEach {$SendCount++}
            Get-TransportServer | get-MessageTrackingLog -ResultSize Unlimited -Start "$StartDate" -End "$EndDate" -Recipients $x.PrimarySMTPAddress -EventID DELIVER | ForEach {$ReceiveCount++}
            Write-Host "User Received"$ReceiveCount" messages." -ForegroundColor Red -BackgroundColor Yellow
            Write-Host "User Sent"$SendCount" messages." -ForegroundColor Red -BackgroundColor Yellow

            #write the data to the array object
            $MailUser = @()
            $MailUser = New-Object -TypeName PSObject -Property @{
                        Name = $x.Name
                        ReceivedMessages = $ReceiveCount
                        SentMessages = $SendCount
                        }
        
            #write the data to the report array data
            $Global:MailUsers += MailUser
        }
    }
#Function to lookup info for Exchange 2013/2016
Function Run2016
    {
        Param ($DomainUsers,$StartDate,$EndDate)
        ForEach($x in $DomainUsers)
        {
            $ReceiveCount = 0
            $SendCount = 0
            Write-Host "Looking up user"$x.Name -ForegroundColor Yellow
            Get-TransportService | get-MessageTrackingLog -ResultSize Unlimited -Start "$StartDate" -End "$EndDate" -Sender $x.PrimarySMTPAddress -EventID RECEIVE | ? {$_.Source -eq "STOREDRIVER"}| ForEach {$SendCount++}
            Get-TransportService | get-MessageTrackingLog -ResultSize Unlimited -Start "$StartDate" -End "$EndDate" -Recipients $x.PrimarySMTPAddress -EventID DELIVER | ForEach {$ReceiveCount++}
            Write-Host "User Received"$ReceiveCount" messages." -ForegroundColor Red -BackgroundColor Yellow
            Write-Host "User Sent"$SendCount" messages." -ForegroundColor Red -BackgroundColor Yellow

            #write the data to the array object
            $MailUser = @()
            $MailUser = New-Object -TypeName PSObject -Property @{
                        Name = $x.Name
                        ReceivedMessages = $ReceiveCount
                        SentMessages = $SendCount
                        }
            
            #write the data to the report array data
            $Global:MailUsers += $MailUser
        }
    }
#------------------------------------------------------End Functions------------------------------------------------------

#Look Up Users associated with the domain.
$DomainUsers = [array](Get-mailbox -ResultSize Unlimited | where {$_.PrimarySMTPAddress -like "*@$Domain"})

#Lookup the server we are connected to
$SessionServer = Get-PSSession | where {$_.State -eq "Opened"}

#Determine what version of the script to run based on version of Exchange.
$EXversion = Get-ExchangeServer -Identity $SessionServer.ComputerName

If ($EXversion.AdminDisplayVersion -like "Version 14*")
    {
        Run2010 -DomainUsers $DomainUsers -StartDate $StartDate -EndDate $EndDate
    }
else
    {
        Run2016 -DomainUsers $DomainUsers -StartDate $StartDate -EndDate $EndDate
    }

#------------------------------------------------------Build the HTML Report------------------------------------------------------
#Set Up HTML Code for report output
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #FF4500;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
Monthly Message Report for $Domain
</title>
"@
$pre = "<h2><b>Report Data for $Domain</b></h2><br>Data Collected from: $StartDate to $EndDate<br><br>"
$Post = "<br><br><small><i>Report ran on $ReportDate</i></small>"

#Export the HTML report to a file
$MailUsers | Select Name,ReceivedMessages,SentMessages | ConvertTo-Html -Head $Header -PreContent $Pre -PostContent $Post | Out-File $HTMLReport
write-host "The report was saved at"$HTMLReport -ForegroundColor Green
