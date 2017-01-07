# Get-MessageReport.ps1
PowerShell script to create an HTML report for the number of messages sent and received per user for a specified domain and date range.


http://ehloexchange.com/get-messagesummary-ps1-powershell-script/

Parameters:
-Domain
-HTMLReport
-StartDate
-EndDate

Examples:
C:\> ./Get-MessageSummary.ps1 -Domain "domain.com"

C:\ .\Get-MessageSummary.ps1 -Domain "domain.com" -HTMLReport "C:\report.html" -StartDate "1/1/2017" -EndDate "1/7/2017"

Report Output: