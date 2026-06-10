#Requires -Version 7
#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Create an Excel overview of all SFTP file transfer configurations.

.DESCRIPTION
    Read all .json configuration files from a folder and create an Excel file
    that gives an overview of all SFTP file transfer configurations combined.

    Each row in the Excel file represents a single Source/Destination path pair.

.PARAMETER ConfigurationJsonFile
    Contains all the parameters used by the script.
    See 'Example.json' for a detailed explanation of parameters.
#>

[CmdLetBinding()]
param (
    [Parameter(Mandatory)]
    [String]$ConfigurationJsonFile
)

begin {
    $ErrorActionPreference = 'stop'

    $systemErrors = [System.Collections.Generic.List[PSObject]]::new()
    $scriptStartTime = Get-Date

    try {
        function Get-StringValueHC {
            param (
                [String]$Name
            )

            if (-not $Name) {
                return $null
            }
            elseif (
                $Name.StartsWith('ENV:', [System.StringComparison]::OrdinalIgnoreCase)
            ) {
                $envVariableName = $Name.Substring(4).Trim()
                $envStringValue = Get-Item -Path "Env:\$envVariableName" -EA Ignore
                if ($envStringValue) {
                    return $envStringValue.Value
                }
                else {
                    throw "Environment variable '$envVariableName' not found."
                }
            }
            else {
                return $Name
            }
        }

        function Get-LogFolderHC {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [string]$Path
            )

            if ($Path -match '^[a-zA-Z]:\\' -or $Path -match '^\\') {
                $fullPath = $Path
            }
            else {
                $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $Path
            }

            if (-not (Test-Path -Path $fullPath -PathType Container)) {
                try {
                    Write-Verbose "Create log folder '$fullPath'"
                    $null = New-Item -Path $fullPath -ItemType Directory -Force
                }
                catch {
                    throw "Failed creating log folder '$fullPath': $_"
                }
            }

            (Resolve-Path $fullPath).ProviderPath
        }

        function Send-MailKitMessageHC {
            [CmdletBinding()]
            param (
                [parameter(Mandatory)]
                [string]$MailKitAssemblyPath,
                [parameter(Mandatory)]
                [string]$MimeKitAssemblyPath,
                [parameter(Mandatory)]
                [string]$SmtpServerName,
                [parameter(Mandatory)]
                [ValidateSet(25, 465, 587, 2525)]
                [int]$SmtpPort,
                [parameter(Mandatory)]
                [string]$Body,
                [parameter(Mandatory)]
                [string]$Subject,
                [parameter(Mandatory)]
                [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
                [string]$From,
                [string]$FromDisplayName,
                [string[]]$To,
                [string[]]$Bcc,
                [int]$MaxAttachmentSize = 20MB,
                [ValidateSet(
                    'None', 'Auto', 'SslOnConnect', 'StartTls', 'StartTlsWhenAvailable'
                )]
                [string]$SmtpConnectionType = 'None',
                [ValidateSet('Normal', 'Low', 'High')]
                [string]$Priority = 'Normal',
                [string[]]$Attachments,
                [PSCredential]$Credential
            )

            begin {
                function Test-IsAssemblyLoaded {
                    param (
                        [String]$Name
                    )
                    foreach ($assembly in [AppDomain]::CurrentDomain.GetAssemblies()) {
                        if ($assembly.FullName -like "$Name, Version=*") {
                            return $true
                        }
                    }
                    return $false
                }

                function Add-Attachments {
                    param (
                        [string[]]$Attachments,
                        [MimeKit.Multipart]$BodyMultiPart
                    )

                    $attachmentList = New-Object System.Collections.ArrayList($null)

                    foreach (
                        $attachmentPath in
                        $Attachments | Sort-Object -Unique
                    ) {
                        try {
                            try {
                                $attachmentItem = Get-Item -LiteralPath $attachmentPath -ErrorAction Stop

                                if ($attachmentItem.PSIsContainer) {
                                    Write-Warning "Attachment '$attachmentPath' is a folder, not a file"
                                    continue
                                }
                            }
                            catch {
                                Write-Warning "Attachment '$attachmentPath' not found"
                                continue
                            }

                            $totalSizeAttachments += $attachmentItem.Length

                            $null = $attachmentList.Add($attachmentItem)

                            if ($totalSizeAttachments -ge $MaxAttachmentSize) {
                                $M = 'The maximum allowed attachment size of {0} MB has been exceeded ({1} MB). No attachments were added to the email. Check the log folder for details.' -f
                                ([math]::Round(($MaxAttachmentSize / 1MB))),
                                ([math]::Round(($totalSizeAttachments / 1MB), 2))

                                Write-Warning $M

                                return [PSCustomObject]@{
                                    AttachmentLimitExceededMessage = $M
                                }
                            }
                        }
                        catch {
                            Write-Warning "Failed to add attachment '$attachmentPath': $_"
                        }
                    }

                    foreach (
                        $attachmentItem in
                        $attachmentList
                    ) {
                        try {
                            Write-Verbose "Add mail attachment '$($attachmentItem.Name)'"

                            $attachment = New-Object MimeKit.MimePart

                            $memoryStream = New-Object System.IO.MemoryStream

                            try {
                                $fileStream = [System.IO.File]::OpenRead($attachmentItem.FullName)
                                $fileStream.CopyTo($memoryStream)
                            }
                            finally {
                                if ($fileStream) {
                                    $fileStream.Dispose()
                                }
                            }

                            $memoryStream.Position = 0

                            $attachment.Content = New-Object MimeKit.MimeContent($memoryStream)

                            $attachment.ContentDisposition = New-Object MimeKit.ContentDisposition

                            $attachment.ContentTransferEncoding = [MimeKit.ContentEncoding]::Base64

                            $attachment.FileName = $attachmentItem.Name

                            $bodyMultiPart.Add($attachment)
                        }
                        catch {
                            Write-Warning "Failed to add attachment '$attachmentItem': $_"
                        }
                    }
                }

                try {
                    if (-not ($To -or $Bcc)) {
                        throw "Either 'To' to 'Bcc' is required for sending emails"
                    }

                    foreach ($email in $To) {
                        if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                            throw "To email address '$email' not valid."
                        }
                    }

                    foreach ($email in $Bcc) {
                        if ($email -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                            throw "Bcc email address '$email' not valid."
                        }
                    }

                    if (-not(Test-IsAssemblyLoaded -Name 'MimeKit')) {
                        try {
                            Write-Verbose "Load MimeKit assembly '$MimeKitAssemblyPath'"
                            Add-Type -Path $MimeKitAssemblyPath
                        }
                        catch {
                            throw "Failed to load MimeKit assembly '$MimeKitAssemblyPath': $_"
                        }
                    }

                    if (-not(Test-IsAssemblyLoaded -Name 'MailKit')) {
                        try {
                            Write-Verbose "Load MailKit assembly '$MailKitAssemblyPath'"
                            Add-Type -Path $MailKitAssemblyPath
                        }
                        catch {
                            throw "Failed to load MailKit assembly '$MailKitAssemblyPath': $_"
                        }
                    }
                }
                catch {
                    throw "Failed to send email to '$To': $_"
                }
            }

            process {
                try {
                    $message = New-Object -TypeName 'MimeKit.MimeMessage'

                    $bodyPart = New-Object MimeKit.TextPart('html')
                    $bodyPart.Text = $Body

                    $bodyMultiPart = New-Object MimeKit.Multipart('mixed')
                    $bodyMultiPart.Add($bodyPart)

                    if ($Attachments) {
                        $params = @{
                            Attachments   = $Attachments
                            BodyMultiPart = $bodyMultiPart
                        }
                        $addAttachments = Add-Attachments @params

                        if ($addAttachments.AttachmentLimitExceededMessage) {
                            $bodyPart.Text += '<p><i>{0}</i></p>' -f
                            $addAttachments.AttachmentLimitExceededMessage
                        }
                    }

                    $message.Body = $bodyMultiPart

                    $fromAddress = New-Object MimeKit.MailboxAddress(
                        $FromDisplayName, $From
                    )
                    $message.From.Add($fromAddress)

                    foreach ($email in $To) {
                        $message.To.Add($email)
                    }

                    foreach ($email in $Bcc) {
                        $message.Bcc.Add($email)
                    }

                    $message.Subject = $Subject

                    switch ($Priority) {
                        'Low' {
                            $message.Headers.Add('X-Priority', '5 (Lowest)')
                            break
                        }
                        'Normal' {
                            $message.Headers.Add('X-Priority', '3 (Normal)')
                            break
                        }
                        'High' {
                            $message.Headers.Add('X-Priority', '1 (Highest)')
                            break
                        }
                        default {
                            throw "Priority type '$_' not supported"
                        }
                    }

                    $smtp = New-Object -TypeName 'MailKit.Net.Smtp.SmtpClient'

                    try {
                        $smtp.Connect(
                            $SmtpServerName, $SmtpPort,
                            [MailKit.Security.SecureSocketOptions]::$SmtpConnectionType
                        )
                    }
                    catch {
                        throw "Failed to connect to SMTP server '$SmtpServerName' on port '$SmtpPort' with connection type '$SmtpConnectionType': $_"
                    }

                    if ($Credential) {
                        try {
                            $smtp.Authenticate(
                                $Credential.UserName,
                                $Credential.GetNetworkCredential().Password
                            )
                        }
                        catch {
                            throw "Failed to authenticate with user name '$($Credential.UserName)' to SMTP server '$SmtpServerName': $_"
                        }
                    }

                    Write-Verbose "Send mail to '$To' with subject '$Subject'"

                    $null = $smtp.Send($message)
                }
                catch {
                    throw "Failed to send email to '$To': $_"
                }
                finally {
                    if ($smtp) {
                        $smtp.Disconnect($true)
                        $smtp.Dispose()
                    }
                    if ($message) {
                        $message.Dispose()
                    }
                }
            }
        }

        #region Import .json file
        Write-Verbose "Import .json file '$ConfigurationJsonFile'"

        $jsonFileItem = Get-Item -LiteralPath $ConfigurationJsonFile -ErrorAction Stop

        $jsonFileContent = Get-Content $jsonFileItem -Raw -Encoding UTF8 |
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        Write-Verbose 'Test .json file properties'

        try {
            @(
                'Path'
            ).where(
                { -not $jsonFileContent.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            #region Test Path exists
            $jsonPath = $jsonFileContent.Path

            if (-not (Test-Path -LiteralPath $jsonPath -PathType Container)) {
                throw "Path '$jsonPath' not found"
            }
            #endregion
        }
        catch {
            throw "Input file '$ConfigurationJsonFile': $_"
        }
        #endregion
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message

        return
    }
}

process {
    if ($systemErrors) { return }

    try {
        #region Read all .json files
        Write-Verbose "Read .json files from '$jsonPath'"

        $jsonFiles = Get-ChildItem -LiteralPath $jsonPath -Filter '*.json' -File

        if (-not $jsonFiles) {
            throw "No .json files found in '$jsonPath'"
        }

        Write-Verbose "Found $($jsonFiles.Count) .json file(s)"
        #endregion

        #region Flatten data to rows
        $excelData = foreach ($jsonFile in $jsonFiles) {
            try {
                Write-Verbose "Read '$($jsonFile.FullName)'"

                $fileContent = Get-Content $jsonFile.FullName -Raw -Encoding UTF8 |
                ConvertFrom-Json

                if (-not $fileContent.Tasks) {
                    Write-Warning "No 'Tasks' found in '$($jsonFile.Name)'"
                    continue
                }

                foreach ($task in $fileContent.Tasks) {
                    foreach ($action in $task.Actions) {
                        foreach ($path in @($action.Paths)) {
                            [PSCustomObject]@{
                                FileName           = $jsonFile.BaseName
                                TaskName           = $task.TaskName
                                SftpComputerName   = $task.Sftp.ComputerName
                                SftpPort           = if ($task.Sftp.Port) { $task.Sftp.Port } else { 22 }
                                Source             = $path.Source
                                Destination        = $path.Destination
                                MatchFileNameRegex = $task.Option.MatchFileNameRegex
                                OverwriteFile      = $task.Option.OverwriteFile
                                ExcludeZeroSizeFile = $task.Option.ExcludeZeroSizeFile
                            }
                        }
                    }
                }
            }
            catch {
                $systemErrors.Add(
                    [PSCustomObject]@{
                        DateTime = Get-Date
                        Message  = "Failed reading '$($jsonFile.Name)': $_"
                    }
                )

                Write-Warning $systemErrors[-1].Message
            }
        }
        #endregion
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
}

end {
    try {
        $settings = $jsonFileContent.Settings

        $scriptName = $settings.ScriptName
        $sendMail = $settings.SendMail
        $saveLogFiles = $settings.SaveLogFiles

        $allLogFilePaths = @()
        $logFolderPath = $null

        #region Get script name
        if (-not $scriptName) {
            Write-Warning "No 'Settings.ScriptName' found in import file."
            $scriptName = 'Default script name'
        }
        #endregion

        #region Create log files
        try {
            $logFolder = Get-StringValueHC $saveLogFiles.Where.Folder

            if ($logFolder) {
                #region Get log folder
                try {
                    $logFolderPath = Get-LogFolderHC -Path $logFolder

                    Write-Verbose "Log folder '$logFolderPath'"
                }
                catch {
                    throw "Failed creating log folder '$logFolder': $_"
                }
                #endregion

                #region Export Excel file
                if ($excelData) {
                    $logFilePath = Join-Path -Path $logFolderPath -ChildPath (
                        '{0} - {1} ({2}).xlsx' -f
                        $scriptStartTime.ToString('yyyy_MM_dd_HHmmss'),
                        $scriptName,
                        $jsonFileItem.BaseName
                    )

                    Write-Verbose "Export $($excelData.Count) rows to '$logFilePath'"

                    $excelParams = @{
                        Path          = $logFilePath
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = 'Overview'
                        TableName     = 'Overview'
                        Verbose       = $false
                    }

                    $excelData | Export-Excel @excelParams

                    $allLogFilePaths += $logFilePath
                }
                #endregion

                #region Export system errors
                if ($systemErrors) {
                    $errorLogFilePath = Join-Path -Path $logFolderPath -ChildPath (
                        '{0} - {1} ({2}) - System errors.xlsx' -f
                        $scriptStartTime.ToString('yyyy_MM_dd'),
                        $scriptName,
                        $jsonFileItem.BaseName
                    )

                    $excelParams = @{
                        Path          = $errorLogFilePath
                        AutoSize      = $true
                        FreezeTopRow  = $true
                        WorksheetName = 'Errors'
                        TableName     = 'Errors'
                        Verbose       = $false
                    }

                    $systemErrors | Export-Excel @excelParams

                    $allLogFilePaths += $errorLogFilePath
                }
                #endregion
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed creating log file in folder '$($saveLogFiles.Where.Folder)': $_"
                }
            )

            Write-Warning $systemErrors[-1].Message
        }
        #endregion

        #region Remove old log files
        if ($saveLogFiles.DeleteLogsAfterDays -gt 0 -and $logFolderPath) {
            $cutoffDate = (Get-Date).AddDays(-$saveLogFiles.DeleteLogsAfterDays)

            Write-Verbose "Removing log files older than $cutoffDate from '$logFolderPath'"

            Get-ChildItem -Path $logFolderPath -File |
            Where-Object { $_.LastWriteTime -lt $cutoffDate } |
            ForEach-Object {
                try {
                    Write-Verbose "Deleting old log file '$_'"
                    Remove-Item -Path $_.FullName -Force
                }
                catch {
                    $systemErrors.Add(
                        [PSCustomObject]@{
                            DateTime = Get-Date
                            Message  = "Failed to remove file '$_': $_"
                        }
                    )

                    Write-Warning $systemErrors[-1].Message
                }
            }
        }
        #endregion

        #region Send email
        try {
            $isSendMail = $false

            switch ($sendMail.When) {
                'Never' {
                    break
                }
                'Always' {
                    $isSendMail = $true
                    break
                }
                'OnError' {
                    if ($systemErrors.Count) {
                        $isSendMail = $true
                    }
                    break
                }
                'OnErrorOrAction' {
                    if ($systemErrors.Count -or $excelData) {
                        $isSendMail = $true
                    }
                    break
                }
                default {
                    throw "SendMail.When '$($sendMail.When)' not supported. Supported values are 'Never', 'Always', 'OnError' or 'OnErrorOrAction'."
                }
            }

            if ($isSendMail) {
                #region Test mandatory fields
                @{
                    'From'                = $sendMail.From
                    'Smtp.ServerName'     = $sendMail.Smtp.ServerName
                    'Smtp.Port'           = $sendMail.Smtp.Port
                    'AssemblyPath.MailKit' = $sendMail.AssemblyPath.MailKit
                    'AssemblyPath.MimeKit' = $sendMail.AssemblyPath.MimeKit
                }.GetEnumerator() |
                Where-Object { -not $_.Value } | ForEach-Object {
                    throw "Input file property 'Settings.SendMail.$($_.Key)' cannot be blank"
                }
                #endregion

                $mailParams = @{
                    From                = Get-StringValueHC $sendMail.From
                    Subject             = "$scriptName"
                    SmtpServerName      = Get-StringValueHC $sendMail.Smtp.ServerName
                    SmtpPort            = Get-StringValueHC $sendMail.Smtp.Port
                    MailKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MailKit
                    MimeKitAssemblyPath = Get-StringValueHC $sendMail.AssemblyPath.MimeKit
                }

                $mailParams.Body = @"
<!DOCTYPE html>
<html>
<head>
<style type="text/css">
    body {
        font-family:verdana;
        font-size:14px;
        background-color:white;
    }
    table {
        border-collapse:collapse;
        border:0px none;
        padding:3px;
        text-align:left;
    }
    td, th {
        border-collapse:collapse;
        border:1px none;
        padding:3px;
        text-align:left;
    }
    #aboutTable th {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
    #aboutTable td {
        color: rgb(143, 140, 140);
        font-weight: normal;
    }
</style>
</head>
<body>
<table>
    <h1>$scriptName</h1>
    <hr size="2" color="#06cc7a">

    $($sendMail.Body)

    $(
        if ($systemErrors.Count) {
            '<table>
                <tr style="background-color: #ffe5ec;">
                    <th>System errors</th>
                    <td>{0}</td>
                </tr>
            </table>' -f $($systemErrors.Count)
        }
    )

    <p>Found <b>$($jsonFiles.Count)</b> .json file(s) with <b>$(@($excelData).Count)</b> path pair(s).</p>

    $(
        if ($allLogFilePaths) {
            '<p><i>* Check the attachment(s) for details</i></p>'
        }
    )

    <hr size="2" color="#06cc7a">
    <table id="aboutTable">
        $(
            if ($scriptStartTime) {
                $runTime = New-TimeSpan -Start $scriptStartTime -End (Get-Date)
                '<tr>
                    <th>Duration</th>
                    <td>{0:00}:{1:00}:{2:00}</td>
                </tr>' -f
                $runTime.Hours, $runTime.Minutes, $runTime.Seconds
            }
        )
        $(
            if ($logFolderPath) {
                '<tr>
                    <th>Log files</th>
                    <td><a href="{0}">Open log folder</a></td>
                </tr>' -f $logFolderPath
            }
        )
        <tr>
            <th>Computer</th>
            <td>$env:COMPUTERNAME</td>
        </tr>
    </table>
</table>
</body>
</html>
"@

                if ($sendMail.FromDisplayName) {
                    $mailParams.FromDisplayName = Get-StringValueHC $sendMail.FromDisplayName
                }

                if ($sendMail.Subject) {
                    $mailParams.Subject = '{0}, {1}' -f
                    $mailParams.Subject, $sendMail.Subject
                }

                if ($sendMail.To) {
                    $mailParams.To = $sendMail.To
                }

                if ($sendMail.Bcc) {
                    $mailParams.Bcc = $sendMail.Bcc
                }

                if ($systemErrors.Count) {
                    $mailParams.Priority = 'High'
                    $mailParams.Subject = '{0} error{1}, {2}' -f
                    $systemErrors.Count,
                    $(if ($systemErrors.Count -ne 1) { 's' }),
                    $mailParams.Subject
                }

                if ($allLogFilePaths) {
                    $mailParams.Attachments = $allLogFilePaths |
                    Sort-Object -Unique
                }

                if ($sendMail.Smtp.ConnectionType) {
                    $mailParams.SmtpConnectionType = Get-StringValueHC $sendMail.Smtp.ConnectionType
                }

                #region Create SMTP credential
                $smtpUserName = Get-StringValueHC $sendMail.Smtp.UserName
                $smtpPassword = Get-StringValueHC $sendMail.Smtp.Password

                if ($smtpUserName -and $smtpPassword) {
                    try {
                        $securePassword = ConvertTo-SecureString -String $smtpPassword -AsPlainText -Force

                        $credential = New-Object System.Management.Automation.PSCredential($smtpUserName, $securePassword)

                        $mailParams.Credential = $credential
                    }
                    catch {
                        throw "Failed to create credential: $_"
                    }
                }
                elseif ($smtpUserName -or $smtpPassword) {
                    throw "Both 'Settings.SendMail.Smtp.Username' and 'Settings.SendMail.Smtp.Password' are required when authentication is needed."
                }
                #endregion

                Send-MailKitMessageHC @mailParams
            }
        }
        catch {
            $systemErrors.Add(
                [PSCustomObject]@{
                    DateTime = Get-Date
                    Message  = "Failed sending email: $_"
                }
            )

            Write-Warning $systemErrors[-1].Message
        }
        #endregion
    }
    catch {
        $systemErrors.Add(
            [PSCustomObject]@{
                DateTime = Get-Date
                Message  = $_
            }
        )

        Write-Warning $systemErrors[-1].Message
    }
    finally {
        if ($systemErrors) {
            $M = 'Found {0} system error{1}' -f
            $systemErrors.Count,
            $(if ($systemErrors.Count -ne 1) { 's' })
            Write-Warning $M

            $systemErrors | ForEach-Object {
                Write-Warning $_.Message
            }

            Write-Warning 'Exit script with error code 1'
            exit 1
        }
        else {
            Write-Verbose 'Script finished successfully'
        }
    }
}
