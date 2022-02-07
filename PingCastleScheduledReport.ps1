
<#PSScriptInfo

.VERSION 0.1
.GUID dcf37da6-cd01-43c7-8e51-a5ce735aab42
.AUTHOR Romain Tiennot
.COMPANYNAME Colibri SAS / ManoMano
.COPYRIGHT Copyright (c) Colibri SAS / Manomano 2021
.TAGS pingcastle security activedirectory
.PROJECTURI https://gist.github.com/aikiox/98f97ccc092557acc1ea958d65f8f361 
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
Version 1.0: Original published version.

.VERSION 0.2
.GUID dcf37da6-cd01-43c7-8e51-a5ce735aab42
.AUTHOR Omer Friedman
.COMPANYNAME Israel National Cyber Directorate
.TAGS pingcastle security activedirectory
.PROJECTURI 
.EXTERNALMODULEDEPENDENCIES
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
Version 0.2: Updates

#>

<#
.SYNOPSIS
    Example of a script to send the PingCastle report

    Copyright (c) Colibri SAS / Manomano 2021

    Permission to use, copy, modify, and distribute this software for any
    purpose with or without fee is hereby granted, provided that the above
    copyright notice and this permission notice appear in all copies.

    THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
    WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
    MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
    ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
    WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
    ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
    OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
.DESCRIPTION
    Execute PingCastle for generate report
    Compares values to the previous report
    Moves reports to a directory
    Update PingCastle
.EXAMPLE
    PS C:\> Send-PingCastleReport.ps1
#>

CLS

function DownloadFromGithub {
    
    if ((Test-NetConnection github.com).pingsucceeded){
        Write-Host "[OK] You are connected to the Internet, Downloading PingCastle" -ForegroundColor Green
        # Download latest dotnet/codeformatter release from github
        $preRelease = $false
        $appName = "PingCastle"
        $repo = "vletoux/pingcastle"
        $filenamePattern = "$appName*.zip"
        $pathExtract = Join-Path $PSScriptRoot $appName
        $null = New-Item -ItemType Directory -Path $pathExtract -Force

        if ($preRelease) {
            $releasesUri = "https://api.github.com/repos/$repo/releases"
            $downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri)[0].assets | Where-Object name -like $filenamePattern ).browser_download_url
        }
        else {
            $releasesUri = "https://api.github.com/repos/$repo/releases/latest"
            $downloadUri = ((Invoke-RestMethod -Method GET -Uri $releasesUri).assets | Where-Object name -like $filenamePattern ).browser_download_url
        }

        $fileVersion = $downloadUri.Split("/")[-1]
        Write-Host "Downloading latest version $fileVersion to $pathExtract" -ForegroundColor Green
        Invoke-WebRequest -Uri $downloadUri -Out $pathExtract\PingCastle.zip
        Expand-Archive -Path "$pathExtract\$appName.zip" -DestinationPath $pathExtract -Force
    } else {
        Write-Host "[Failed] You are not connected to the internet, Please download and extract PingCastle to $pingCastleFullpath" -ForegroundColor Red
        break
    }
}

function sendReportByMail {
    param([string]$Attachment)
    if (!(Test-Path "$PSScriptRoot\email-creds.clixml"))
    {      
        Write-Host "[Note] In order to send email using Gmail you need to generate a gmail windows desktop application password using this link:" -ForegroundColor Yellow
        Start-Process "https://security.google.com/settings/security/apppasswords"
        Write-Host "https://security.google.com/settings/security/apppasswords" -ForegroundColor Yellow
        Write-Host "(eg. user name = donald@trump.com)"  -ForegroundColor Yellow
        Write-Host "(eg. password = jdynzpsxmjepwxnn)"  -ForegroundColor Yellow
        $credStore = Get-Credential
        $credStore | Export-CliXml "$PSScriptRoot\email-creds.clixml"
        Write-Host "[OK] email credentials were saved to [$PSScriptRoot\email-creds.clixml] file" -ForegroundColor Green
    } 

    $getCred = Import-CliXml "$PSScriptRoot\email-creds.clixml"
    $mailUser = $getCred.UserName.ToString()

    if (!(Test-Path "$PSScriptRoot\email-conf.clixml"))
    {   
       @{
        To = Read-Host "Input main email address you want the report to be sent"
        Cc = Read-Host "Input a cc email address you want the report to be sent"
        } | Export-CliXml "$PSScriptRoot\email-conf.clixml"
    }
    
    $emailConf = Import-CliXml "$PSScriptRoot\email-conf.clixml"
    $To = $emailConf.To
    $Cc = $emailConf.Cc
    $sentDate = Get-Date
    $Subject = "PingCastle Report $sentDate"
    $Body = "This is your PingCastle report"
    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"
    Send-MailMessage -From $mailUser -to $To -Cc $Cc -Subject $Subject `
    -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    -Credential $getCred -Attachments $Attachment
    Write-Host "[OK] Sending report by mail From:$mailUser To:$to Cc:$Cc"-ForegroundColor Green
}

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$score = [pscustomobject]@{
    lowScore = "green"
    mediumScore = "yellow"
    highScore = "magenta"
    criticalScore = "red"
}

#Check if computer is connected to a domain
if ((gwmi win32_computersystem).partofdomain -eq $true) {
    write-host -fore green "[OK] You are connected to domain"    
    
    $domainServer = ($env:LOGONSERVER).Replace("\\","")
    $DomainDNS =  (whoami).split("\")[0]
    $DomainUser = whoami
    $ApplicationName = 'PingCastle'
    $PingCastle = [pscustomobject]@{
        Name            = $ApplicationName
        ProgramPath     = Join-Path $PSScriptRoot $ApplicationName
        ProgramName     = '{0}.exe' -f $ApplicationName
        Arguments       = "--healthcheck --level Full"
        ReportFileName  = 'ad_hc_{0}' -f ($DomainDNS).ToLower()
        ReportFolder    = "Reports"
        ScoreFileName   = '{0}Score.txt' -f $ApplicationName
        ProgramUpdate   = '{0}AutoUpdater.exe' -f $ApplicationName
        ArgumentsUpdate = '--wait-for-days 30'
    }

} else {
    
    if (!(Test-Path "$PSScriptRoot\domain-config.clixml"))
    {
        write-host "[Note] You are not connected to a domain, Please fill in the next details" -ForegroundColor Yellow
    
        @{ 
            domainServer = Read-Host "Input the name or ip of your Domain Controller (eg. DC1)"
            domainDNS = Read-Host "Input FQDN of your organization network (eg. cyber.gov.il)"
        } | Export-CliXml "$PSScriptRoot\domain-config.clixml"
        Write-Host "[Confirmation] Data was saved to file $PSScriptRoot\domain-config.clixml" -ForegroundColor Yellow

    }

    $domainConf = Import-Clixml "$PSScriptRoot\domain-config.clixml"
    $domainServer = $domainConf.domainServer
    $DomainDNS = $domainConf.domainDNS
    
    if (!(Test-Path "$PSScriptRoot\domainUser-creds.clixml"))
    {
        Write-Host "Input full credentials of a user with Domain Admin credentials (eg. cyber\silentbob)"
        $DomainUserCredStore = Get-Credential
        $DomainUserCredStore | Export-CliXml "$PSScriptRoot\domainUser-creds.clixml"
        Write-Host "[OK] Domain user credentials were saved to [$PSScriptRoot\domainUser-creds.clixml] file" -ForegroundColor Green
    }
    $getDomainUserCreds = Import-CliXml "$PSScriptRoot\domainUser-creds.clixml"
    $DomainUser = $getDomainUserCreds.username
    $userPassword = $getDomainUserCreds.GetNetworkCredential().password
    
    $ApplicationName = 'PingCastle'
    $PingCastle = [pscustomobject]@{
        Name            = $ApplicationName
        ProgramPath     = Join-Path $PSScriptRoot $ApplicationName
        ProgramName     = '{0}.exe' -f $ApplicationName
        Arguments       = "--server $domainServer --user $domainUser --password $userPassword --healthcheck --level Full"
        ReportFileName  = 'ad_hc_{0}' -f ($DomainDNS).ToLower()
        ReportFolder    = "Reports"
        ScoreFileName   = '{0}Score.txt' -f $ApplicationName
        ProgramUpdate   = '{0}AutoUpdater.exe' -f $ApplicationName
        ArgumentsUpdate = '--wait-for-days 30'
    }

    $OriginalProgressPreference = $Global:ProgressPreference
    $Global:ProgressPreference = 'SilentlyContinue'
    if (!((Test-NetConnection labdc -WarningAction SilentlyContinue).PingSucceeded)){
        Write-Host "[Failed] Could not ping the domain server [$domainServer], please run the script again" -ForegroundColor Red
        break
    } else {
        Write-Host "[OK] Succeded pinging the domain server [$domainServer]" -ForegroundColor Green
    }

    $Global:ProgressPreference = $OriginalProgressPreference
}

$pingCastleFullpath = Join-Path $PingCastle.ProgramPath $PingCastle.ProgramName
$pingCastleUpdateFullpath = Join-Path $PingCastle.ProgramPath $PingCastle.ProgramUpdate
$pingCastleReportLogs = Join-Path $PingCastle.ProgramPath $PingCastle.ReportFolder
$pingCastleScoreFileFullpath = Join-Path $pingCastleReportLogs $PingCastle.ScoreFileName
$pingCastleReportFullpath = Join-Path $PingCastle.ProgramPath ('{0}.html' -f $PingCastle.ReportFileName)
$pingCastleReportXMLFullpath = Join-Path $PingCastle.ProgramPath ('{0}.xml' -f $PingCastle.ReportFileName)

$pingCastleReportDate = Get-Date -UFormat %Y%m%d_%H%M%S
$pingCastleReportFileNameDate = ('{0}_{1}' -f $pingCastleReportDate, ('{0}.html' -f $PingCastle.ReportFileName))

$sentNotification = $false

$splatProcess = @{
    WindowStyle = 'Hidden'
    Wait        = $true
}
#endregion

# Check if program exist
if (-not(Test-Path $pingCastleFullpath)) {
    Write-Host "[Checking] Path  to PingCastle application not found $pingCastleFullpath" -ForegroundColor Yellow
    DownloadFromGithub
} else {
    Write-Host "[OK] PingCastle was found" -ForegroundColor Green
    # Try to start pingcastle update tool and catch any error
    try {
        Write-Host "[Checking] Trying to update PingCastle to latest version" -ForegroundColor Yellow
        Start-Process -FilePath $pingCastleUpdateFullpath -ArgumentList $PingCastle.ArgumentsUpdate @splatProcess
        Write-Host "[OK] PingCastle version udate completed" -ForegroundColor Green
        }
    Catch {
        if ((Test-NetConnection github.com).pingsucceeded){
            Write-Host  "[Failed] Could not run PingCastle update tool $pingCastleUpdateFullpath" -ForegroundColor Red
          }
    }

}

# Check if log directory exist. If not, create it
if (-not (Test-Path $pingCastleReportLogs)) {
    try {
        $null = New-Item -Path $pingCastleReportLogs -ItemType directory
    }
    Catch {
        Write-Host "[Failed] Error creating reports directory $pingCastleReportLogs"  -ForegroundColor Red
    }
}

# Try to start program and catch any error
try {
    Write-Host "[OK] Start running PingCastle, Please wait..." -ForegroundColor Green
    Push-Location -Path $PingCastle.ProgramPath
    Start-Process -FilePath $pingCastleFullpath -ArgumentList $PingCastle.Arguments @splatProcess
}
Catch {
    Write-Host "[Failed] Error executing $pingCastleFullpath" -ForegroundColor Red
}

# Check if report exist after execution
foreach ($pingCastleTestFile in ($pingCastleReportFullpath, $pingCastleReportXMLFullpath)) {
    if (-not (Test-Path $pingCastleTestFile)) {
        Write-Host "[Failed] Report file not found $pingCastleTestFile" -ForegroundColor Red
    }
}

# Get content on XML file
try {
    $contentPingCastleReportXML = $null
    $contentPingCastleReportXML = (Select-Xml -Path $pingCastleReportXMLFullpath -XPath "/HealthcheckData").node
}
catch {
    Write-Host "[Failed] Unable to read content of report xml file $pingCastleReportXMLFullpath" -ForegroundColor Red
}

# Convert to json all score from xml file
try {
    $contentPingCastleReport = $contentPingCastleReportXMLToJSON = $null
    $contentPingCastleReport = $contentPingCastleReportXML | Select-Object *Score
    $contentPingCastleReportXMLToJSON = $contentPingCastleReport | ConvertTo-Json -Compress
}
catch {
    Write-Host "[Failed] Unable to convert report xml to json" -ForegroundColor Red
}

# Check if PingCastle previous score file exist
if (-not (Test-Path $pingCastleScoreFileFullpath)) {
    # if don't exist, sent report
    $sentNotification = $true
}
else {
    try {
        # Get content of previous PingCastle score
        $contentPingCastleScoreFile = Get-Content $pingCastleScoreFileFullpath

        # Compare value between previous score and current score
        if ($contentPingCastleScoreFile -cne $contentPingCastleReportXMLToJSON) {
            # If value is different, sent report
            $sentNotification = $true
        }
    }
    catch {
        Write-Host -Message ("[Failed] Unable to read content of the last score file {0}" -f $pingCastleScoreFileFullpath) -ForegroundColor Red
        $sentNotification = $true
    }
}

# If content is same, don't sent report
if ($sentNotification -eq $false) {
    Remove-Item ("{0}.{1}" -f (Join-Path $PingCastle.ProgramPath $PingCastle.ReportFileName), '*')
    Write-Host "[OK] Score is the same as the previous run, New report deleted from disk" -ForegroundColor Green
    Pop-Location
    exit
}

Write-Host "=============================================="
Write-Host "PingCastle Report for Domain $DomainDNS"
Write-Host "=============================================="

Switch($contentPingCastleReport.GlobalScore)
{
    {$_ -in 0..24} {Write-Host "GlobalScore:`t" $contentPingCastleReport.GlobalScore "(Low)" -ForegroundColor $score.lowscore}
    {$_ -in 25..49} {Write-Host "GlobalScore:`t" $contentPingCastleReport.GlobalScore "(Medium)" -ForegroundColor $score.mediumScore}
    {$_ -in 50..74} {Write-Host "GlobalScore:`t" $contentPingCastleReport.GlobalScore "(High)" -ForegroundColor $score.highscore}
    {$_ -in 75..100} {Write-Host "GlobalScore:`t" $contentPingCastleReport.GlobalScore "(Critical)" -ForegroundColor $score.criticalscore}
}

Write-Host "=============================================="

Switch($contentPingCastleReport.StaleObjectsScore)
{    
    {$_ -in 0..24} {Write-Host "StaleObjectsScore:`t" $contentPingCastleReport.StaleObjectsScore "(Low)" -ForegroundColor $score.lowscore}
    {$_ -in 25..49} {Write-Host "StaleObjectsScore:`t" $contentPingCastleReport.StaleObjectsScore "(Medium)" -ForegroundColor $score.mediumScore}
    {$_ -in 50..74} {Write-Host "StaleObjectsScore:`t" $contentPingCastleReport.StaleObjectsScore "(High)" -ForegroundColor $score.highscore}
    {$_ -in 75..100} {Write-Host "StaleObjectsScore:`t" $contentPingCastleReport.StaleObjectsScore "(Critical)" -ForegroundColor $score.criticalscore}
}

Switch($contentPingCastleReport.PrivilegiedGroupScore)
{
    {$_ -in 0..24} {Write-Host "PrivilegiedGroupScore:`t" $contentPingCastleReport.PrivilegiedGroupScore "(Low)" -ForegroundColor $score.lowscore}
    {$_ -in 25..49} {Write-Host "PrivilegiedGroupScore:`t" $contentPingCastleReport.PrivilegiedGroupScore "(Medium)" -ForegroundColor $score.mediumScore}
    {$_ -in 50..74} {Write-Host "PrivilegiedGroupScore:`t" $contentPingCastleReport.PrivilegiedGroupScore "(High)" -ForegroundColor $score.highscore}
    {$_ -in 75..100} {Write-Host "PrivilegiedGroupScore:`t" $contentPingCastleReport.PrivilegiedGroupScore "(Critical)" -ForegroundColor $score.criticalscore}
}

Switch($contentPingCastleReport.TrustScore)
{
    {$_ -in 0..24} {Write-Host "TrustScore:`t`t" $contentPingCastleReport.TrustScore  "(Low)" -ForegroundColor $score.lowscore}
    {$_ -in 25..49} {Write-Host "TrustScore:`t`t" $contentPingCastleReport.TrustScore "(Medium)" -ForegroundColor $score.mediumScore}
    {$_ -in 50..74} {Write-Host "TrustScore:`t`t" $contentPingCastleReport.TrustScore "(High)" -ForegroundColor $score.highscore}
    {$_ -in 75..100} {Write-Host "TrustScore:`t`t" $contentPingCastleReport.TrustScore "(Critical)" -ForegroundColor $score.criticalscore}
}

Switch($contentPingCastleReport.AnomalyScore)
{
    {$_ -in 0..24} {Write-Host "AnomalyScore:`t`t" $contentPingCastleReport.AnomalyScore "(Low)" -ForegroundColor $score.lowscore}
    {$_ -in 25..49} {Write-Host "AnomalyScore:`t`t" $contentPingCastleReport.AnomalyScore "(Medium)" -ForegroundColor $score.mediumScore}
    {$_ -in 50..74} {Write-Host "AnomalyScore:`t`t" $contentPingCastleReport.AnomalyScore "(High)" -ForegroundColor $score.highscore}
    {$_ -in 75..100} {Write-Host "AnomalyScore:`t`t" $contentPingCastleReport.AnomalyScore "(Critical)" -ForegroundColor $score.criticalscore}
}


# Update report file with current score
try {
    $contentPingCastleReportXMLToJSON | Out-File $pingCastleScoreFileFullpath -Force
}
Catch {
    Write-Host "[Failed] Error updating report score file $pingCastleScoreFileFullpath" -ForegroundColor Red
}

sendReportByMail $pingCastleReportFullpath

# Move report to logs directory
try {
    $pingCastleMoveFile = (Join-Path $pingCastleReportLogs $pingCastleReportFileNameDate)
    Move-Item -Path $pingCastleReportFullpath -Destination $pingCastleMoveFile
    Write-Host "[Note] You can open report from [$pingCastleMoveFile]" -ForegroundColor Yellow
}
catch {
    Write-Host "[Failed] Error moving report file to the logs directory $pingCastleReportFullpath" -ForegroundColor Red
}

Pop-Location