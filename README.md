# Get-MessageSummary.ps1
PowerShell script to create an HTML report for the number of messages sent and received per user for a specified domain and date range.


http://ehloexchange.com/get-messagesummary-ps1-powershell-script/

Parameters:
-Domain
-HTMLReport
-StartDate
-EndDate

Examples:
C:\> .\Get-MessageSummary.ps1 -Domain "sampledomain.com"

C:> .\Get-MessageSummary.ps1 -Domain "sampledomain.com" -HTMLReport "C:\report.html" -StartDate "1/1/2017" -EndDate "1/6/2017"

Report Output:

![forreadme](https://cloud.githubusercontent.com/assets/23706901/21744219/ae9c3f44-d4ce-11e6-9905-ff14a06f5ba4.png)
