<#
    .SYNOPSIS
    Creates a HTML Report containing message transport details for a specific domain and date range 
   
    Matt Krause
    Version 1 December 2016
    
    .DESCRIPTION
    
    This script creates a HTML report showing the following information about an Exchange domain. 
            
    * Number of messages sent per user.
    * Number of messages received per user.
        
    This searches all Tranport logs in the environment to pull this information.
    
    .PARAMETER Domain
    Desired domain for the report data
    
    .PARAMETER Start
    Start date for report data
    
    .PARAMETER End
    End date for report data
    
    .PARAMETER HTMLReport
    Filename to write HTML Report to

    .EXAMPLE
    Generate the HTML report 
    .\Get-TransportReport.ps1 -Domain "domain.com" -HTMLReport "C:\Temp\Report.html" -StartDate "1/1/2016" -EndDate "1/31/2016"
    #>

#Read in parameters
param(
    [parameter(Position=0,Mandatory=$true,HelpMessage="Insert Domain")][string]$Domain,
    [parameter(Position=1,Mandatory=$false,HelpMessage="HTML Report Location")][string]$HTMLReport="1",
    [parameter(position=2,Mandatory=$false,HelpMEssage="Start Date")][string]$StartDate="1",
    [parameter(position=3,Mandatory=$false,HelpMessage="End Date")][string]$EndDate="1"
    )

#Set Variables
if ($HTMLReport -eq "1")
    {
        $ReportLocation = $PSScriptRoot +"\"
        $ReportDate = Get-Date
        $HTMLReportName = "MessageSummary" + $ReportDate.Month + $ReportDate.Day + $ReportDate.Year + $ReportDate.Hour + $ReportDate.Minute + $ReportDate.Second + ".html"
        $HTMLReport = $ReportLocation+$HTMLReportName
    }
$Month = Get-Date -Format "MM"
$Year = Get-Date -Format "yyyy"
if ($StartDate -eq "1")
    {
        $StartDate = $Month+"/01/"+$Year
    }
if ($EndDate -eq "1")
    {
        $EndDate = Get-Date -Format "MM/dd/yyyy"
    }

#Set up Array to hold report data
$MailUsers = @()

#Look Up Users associated with the domain.
$DomainUsers = [array](Get-mailbox -ResultSize Unlimited | where {$_.PrimarySMTPAddress -like "*@$Domain"})

#Run the report for Received Messages
ForEach($x in $DomainUsers)
    {
        $ReceiveCount = 0
        $SendCount = 0
        Write-Host "Looking up user"$x.Name -ForegroundColor Yellow
        Get-TransportServer | get-MessageTrackingLog -ResultSize Unlimited -Start "$StartDate" -End "$EndDate" -Sender $x.PrimarySMTPAddress -EventID RECEIVE | ? {$_.Source -eq "STOREDRIVER"}| ForEach {$ReceiveCount++}
        Get-TransportServer | get-MessageTrackingLog -ResultSize Unlimited -Start "$StartDate" -End "$EndDate" -Recipients $x.PrimarySMTPAddress -EventID DELIVER | ForEach {$SendCount++}
        Write-Host "User Received"$ReceiveCount" messages." -ForegroundColor Red -BackgroundColor Yellow
        Write-Host "User Sent"$SendCount" messages." -ForegroundColor Red -BackgroundColor Yellow

        #write the data to the array object
        $mailUser = @()
        $MailUser = New-Object -TypeName PSObject -Property @{
                        Name = $x.Name
                        ReceivedMessages = $ReceiveCount
                        SentMessages = $SendCount
                        }
        

        #write the data to the report array data
        $MailUsers += $MailUser
    }

#Build the HTML Report
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
#un-comment the following line to creat a .csv file
#$MailUsers | Export-Csv -NoTypeInformation -Path $HTMLReportLocation+$HTMLReport+".csv"

#html
$MailUsers | Select Name,ReceivedMessages,SentMessages | ConvertTo-Html -Head $Header -PreContent $Pre -PostContent $Post | Out-File $HTMLReport
write-host "The report was saved at"$HTMLReport -ForegroundColor Green
