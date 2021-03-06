# Pingcastle-Scheduled-Report-By-Mail

Powershell script to automate running PingCastle tool for Active Directory Health audit and sending report by mail
also comparing the scoring results with last run to check if there was a change in scoring.

![image](https://user-images.githubusercontent.com/6965771/153182886-7739fc9e-3bb5-4c59-98d3-53a59c1f2d1a.png)

Features:
1. Automatically downloads latest PingCastle version 
2. Updates PingCastle to newer versions (if already exists)
3. Executes PingCastle full audit in order to generate a domain health report
4. Compare values with previous report
5. Sends the Ad health HTML report using Gmail SMTP service (needs to generate an application token 1st)
6. Option to add and run as a scheduled task 

Note: If you read your e-mails in a browser, you need tp download the HTML report to a directory so it will open correctly

In order to run the script you need to open a Powershell terminal:

PS C:\> .\PingCastleScheduledReport.ps1

In order to be able to send emails using gmail you will need to create a gmail application password using this link:
https://myaccount.google.com/apppasswords

Note: Tested on windows 10 Powershell 5.1

Credits:
Based on idea and script by aikiox / Send-PingCastleReport.ps1
https://gist.github.com/aikiox/98f97ccc092557acc1ea958d65f8f361

