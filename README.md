# Pingcastle-Scheduled-Report-By-Mail

Powershell script to automate running PingCastle tool for Active Directory Health audit and sending report by mail
also comparing the scoring results with last run to check if there was a change in scoring.

Features:
1. Automatically sownloads the latest version PingCastle
2. Updates PingCastle to newer versions if exists
3. Execute PingCastle in order to generate a domain health report
4. Compare values with previous report
5. Sends the HTML report using Gmail
6. Option to run as a scheduled task

Note: If you read your e-mails in a browser, you need tp download the HTML report to a directory so it will open correctly

In order to run the script you need to open a Powershell terminal:

PS C:\> .\PingCastleScheduledReport.ps1

In order to be able to send emails using gmail you will need to create a gmail application password using this link:
https://myaccount.google.com/apppasswords

Note: Tested on windows 10 Powershell 5.1


![image](https://user-images.githubusercontent.com/6965771/153041942-71d6cabf-b8e7-41cf-b5ef-dc3bbf1f9cb1.png)
