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
$Global:MailUsers = @()

#----Functions----
#Function to lookup info for Exchange 2010
Function 2010Lookup
    {
        Param ($DomainUsers)
        write-host "Starting 2010 Version"
        ForEach($x in $DomainUsers,$StartDate,$EndDate)
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
            $global:MailUsers += MailUser
        }
    }
#Function to lookup info for Exchange 2013/2016
Function 2016Lookup
    {
        Param ($DomainUsers,$StartDate,$EndDate)
        write-host "Starting 2013/2016 Version"
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


#Lookup the server we are connected to
$SessionServer = Get-PSSession | where {$_.State -eq "Opened"}
Write-Host "Server is $SessionServer"
#Determine what version of the script to run based on version of Exchange.
$EXversion = Get-ExchangeServer -Identity $SessionServer.ComputerName
Write-Host "Exchange Version $EXversion"

If ($EXversion.AdminDisplayVersion -like "Version 14*")
    {
        Write-Host "Will use Exchange 2010 Version"
        2010Lookup -DomainUsers $DomainUsers -StartDate $StartDate -EndDate $EndDate
    }
else
    {
        Write-Host "Will use 2013/2016 Version"
        2016Lookup -DomainUsers $DomainUsers -StartDate $StartDate -EndDate $EndDate
    }

#Look Up Users associated with the domain.
$DomainUsers = [array](Get-mailbox -ResultSize Unlimited | where {$_.PrimarySMTPAddress -like "*@$Domain"})

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
