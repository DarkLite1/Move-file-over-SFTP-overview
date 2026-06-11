#Requires -Modules Pester
#Requires -Modules ImportExcel
#Requires -Version 7

BeforeAll {
    $testInputFile = @{
        Path     = $null
        Settings = @{
            ScriptName   = 'Test (Brecht)'
            SendMail     = @{
                From         = 'm@example.com'
                To           = @('007@example.com')
                Subject      = 'Email subject'
                Body         = 'Email body'
                Smtp         = @{
                    ServerName     = 'SMTP_SERVER'
                    Port           = 25
                    ConnectionType = 'StartTls'
                    UserName       = 'bob'
                    Password       = 'pass'
                }
                AssemblyPath = @{
                    MailKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MailKit.4.11.0\lib\net8.0\MailKit.dll'
                    MimeKit = 'C:\Program Files\PackageManagement\NuGet\Packages\MimeKit.4.11.0\lib\net8.0\MimeKit.dll'
                }
            }
            SaveLogFiles = @{
                Where               = @{
                    Folder         = (New-Item 'TestDrive:/log' -ItemType Directory).FullName
                }
                DeleteLogsAfterDays = 30
            }
        }
    }

    $testOutParams = @{
        FilePath = (New-Item 'TestDrive:/Test.json' -ItemType File).FullName
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ConfigurationJsonFile = $testOutParams.FilePath
    }

    function Copy-ObjectHC {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [Object]$InputObject
        )

        $jsonString = $InputObject | ConvertTo-Json -Depth 100
        $deepCopy = $jsonString | ConvertFrom-Json
        return $deepCopy
    }

    function Test-NewJsonFileHC {
        try {
            if (-not $testNewInputFile) {
                throw "Variable 'testNewInputFile' cannot be blank"
            }

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams
        }
        catch {
            throw "Failure in Test-NewJsonFileHC: $_"
        }
    }

    function Get-StringValueHC {
        param(
            [String]$Name
        )
        $Name
    }

    function Send-MailKitMessageHC {
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
            [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
            [string]$From,
            [parameter(Mandatory)]
            [string]$Body,
            [parameter(Mandatory)]
            [string]$Subject,
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
    }

    function Remove-TestFileHC {
        Get-ChildItem -Path $testInputFile.Settings.SaveLogFiles.Where.Folder -File -EA Ignore |
        Remove-Item -Force

        if ($testJsonFolder) {
            Get-ChildItem -Path $testJsonFolder -File -EA Ignore |
            Remove-Item -Force
        }
    }

    Mock Send-MailKitMessageHC
    Mock Write-EventLog
}

Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ConfigurationJsonFile') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}

Describe 'create a system error when' {
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ConfigurationJsonFile = 'nonExisting.json'

            { .$testScript @testNewParams } | Should -Not -Throw

            $LASTEXITCODE | Should -Be 1
        }
        It "property 'Path' is missing" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Path = $null

            Test-NewJsonFileHC

            { .$testScript @testParams } | Should -Not -Throw

            $LASTEXITCODE | Should -Be 1

            Should -Invoke Send-MailKitMessageHC -Times 1
        }
        It "property 'Path' folder does not exist" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Path = 'TestDrive:\nonExistingFolder'

            Test-NewJsonFileHC

            { .$testScript @testParams } | Should -Not -Throw

            $LASTEXITCODE | Should -Be 1

            Should -Invoke Send-MailKitMessageHC -Times 1
        }
    }
}

Describe 'when the input file is valid' {
    BeforeAll {
        $testJsonFolder = (New-Item 'TestDrive:/jsonFiles' -ItemType Directory).FullName
    }

    Context 'with a single task and single path' {
        BeforeAll {
            #region Create test JSON files
            $testJsonContent = @{
                Tasks = @(
                    @{
                        TaskName = 'Upload files'
                        Sftp     = @{
                            ComputerName = 'sftp.server.com'
                            Port         = 22
                            Credential   = @{
                                UserName = 'user1'
                                Password = 'pass1'
                            }
                        }
                        Option   = @{
                            OverwriteFile       = $true
                            ExcludeZeroSizeFile = $false
                            MatchFileNameRegex  = '\.txt$'
                        }
                        Actions  = @(
                            @{
                                ComputerName = 'PC1'
                                Paths        = @(
                                    @{
                                        Source      = '\\server\share'
                                        Destination = 'sftp:/upload/'
                                    }
                                )
                            }
                        )
                    }
                )
            }

            $testJsonContent | ConvertTo-Json -Depth 7 |
            Out-File (Join-Path $testJsonFolder 'Single.json')

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Path = $testJsonFolder

            Test-NewJsonFileHC
            #endregion

            .$testScript @testParams
        }

        It 'should create an Excel file with 1 row' {
            $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

            $testExcelFile = Get-ChildItem -Path $testLogFolder -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '*System errors*' }

            $testExcelFile | Should -Not -BeNullOrEmpty

            $excelContent = Import-Excel -Path $testExcelFile.FullName

            @($excelContent).Count | Should -Be 1

            $excelContent[0].FileName | Should -Be 'Single'
            $excelContent[0].TaskName | Should -Be 'Upload files'
            $excelContent[0].SftpComputerName | Should -Be 'sftp.server.com'
            $excelContent[0].SftpPort | Should -Be 22
            $excelContent[0].Source | Should -Be '\\server\share'
            $excelContent[0].Destination | Should -Be 'sftp:/upload/'
            $excelContent[0].MatchFileNameRegex | Should -Be '\.txt$'
            $excelContent[0].OverwriteFile | Should -BeTrue
            $excelContent[0].ExcludeZeroSizeFile | Should -BeFalse
        }

        AfterAll { Remove-TestFileHC }
    }

    Context 'with multiple tasks and multiple paths' {
        BeforeAll {
            #region Create test JSON files
            $testJsonContent = @{
                Tasks = @(
                    @{
                        TaskName = 'Task A'
                        Sftp     = @{
                            ComputerName = 'sftp1.example.com'
                            Port         = 2222
                            Credential   = @{
                                UserName = 'user1'
                                Password = 'pass1'
                            }
                        }
                        Option   = @{
                            OverwriteFile       = $false
                            ExcludeZeroSizeFile = $true
                            MatchFileNameRegex  = '\.csv$'
                        }
                        Actions  = @(
                            @{
                                ComputerName = 'PC1'
                                Paths        = @(
                                    @{
                                        Source      = '\\server\shareA'
                                        Destination = 'sftp:/folderA/'
                                    }
                                    @{
                                        Source      = '\\server\shareB'
                                        Destination = 'sftp:/folderB/'
                                    }
                                )
                            }
                        )
                    }
                    @{
                        TaskName = 'Task B'
                        Sftp     = @{
                            ComputerName = 'sftp2.example.com'
                            Credential   = @{
                                UserName = 'user2'
                                Password = 'pass2'
                            }
                        }
                        Option   = @{
                            OverwriteFile       = $true
                            ExcludeZeroSizeFile = $false
                            MatchFileNameRegex  = '.*'
                        }
                        Actions  = @(
                            @{
                                ComputerName = 'PC2'
                                Paths        = @{
                                    Source      = 'sftp:/downloads/'
                                    Destination = '\\server\shareC'
                                }
                            }
                        )
                    }
                )
            }

            $testJsonContent | ConvertTo-Json -Depth 7 |
            Out-File (Join-Path $testJsonFolder 'Multi.json')

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Path = $testJsonFolder

            Test-NewJsonFileHC

            .$testScript @testParams
        }

        It 'should create an Excel file with 3 rows' {
            $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

            $testExcelFile = Get-ChildItem -Path $testLogFolder -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '*System errors*' }

            $testExcelFile | Should -Not -BeNullOrEmpty

            $excelContent = Import-Excel -Path $testExcelFile.FullName

            @($excelContent).Count | Should -Be 3
        }

        It 'should have correct data for task A path 1' {
            $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

            $testExcelFile = Get-ChildItem -Path $testLogFolder -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '*System errors*' }

            $excelContent = Import-Excel -Path $testExcelFile.FullName

            $row = $excelContent | Where-Object {
                $_.TaskName -eq 'Task A' -and $_.Source -eq '\\server\shareA'
            }

            $row.FileName | Should -Be 'Multi'
            $row.SftpComputerName | Should -Be 'sftp1.example.com'
            $row.SftpPort | Should -Be 2222
            $row.Destination | Should -Be 'sftp:/folderA/'
            $row.MatchFileNameRegex | Should -Be '\.csv$'
            $row.OverwriteFile | Should -BeFalse
            $row.ExcludeZeroSizeFile | Should -BeTrue
        }

        It 'should default SftpPort to 22 when not specified' {
            $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

            $testExcelFile = Get-ChildItem -Path $testLogFolder -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '*System errors*' }

            $excelContent = Import-Excel -Path $testExcelFile.FullName

            $row = $excelContent | Where-Object {
                $_.TaskName -eq 'Task B'
            }

            $row.SftpPort | Should -Be 22
        }

        AfterAll { Remove-TestFileHC }
    }

    Context 'with Paths as a single object instead of array' {
        BeforeAll {
            #region Create test JSON file with single Paths object
            $testJsonContent = @{
                Tasks = @(
                    @{
                        TaskName = 'Single path object'
                        Sftp     = @{
                            ComputerName = 'sftp.test.com'
                            Credential   = @{
                                UserName = 'user1'
                                Password = 'pass1'
                            }
                        }
                        Option   = @{
                            OverwriteFile       = $false
                            ExcludeZeroSizeFile = $false
                            MatchFileNameRegex  = '\.pdf$'
                        }
                        Actions  = @(
                            @{
                                ComputerName = $null
                                Paths        = @{
                                    Source      = '\\server\pdfs'
                                    Destination = 'sftp:/incoming/'
                                }
                            }
                        )
                    }
                )
            }

            $testJsonContent | ConvertTo-Json -Depth 7 |
            Out-File (Join-Path $testJsonFolder 'SinglePath.json')

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Path = $testJsonFolder

            Test-NewJsonFileHC

            .$testScript @testParams
        }

        It 'should still create 1 row in the Excel file' {
            $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

            $testExcelFile = Get-ChildItem -Path $testLogFolder -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '*System errors*' }

            $testExcelFile | Should -Not -BeNullOrEmpty

            $excelContent = Import-Excel -Path $testExcelFile.FullName

            @($excelContent).Count | Should -Be 1

            $excelContent[0].Source | Should -Be '\\server\pdfs'
            $excelContent[0].Destination | Should -Be 'sftp:/incoming/'
        }

        AfterAll { Remove-TestFileHC }
    }

    Context 'with multiple .json files in the folder' {
        BeforeAll {
            #region Create test JSON files
            @{
                Tasks = @(
                    @{
                        TaskName = 'File1 task'
                        Sftp     = @{
                            ComputerName = 'sftp1.com'
                            Port         = 22
                            Credential   = @{
                                UserName = 'u1'
                                Password = 'p1'
                            }
                        }
                        Option   = @{
                            OverwriteFile       = $true
                            ExcludeZeroSizeFile = $false
                            MatchFileNameRegex  = '.*'
                        }
                        Actions  = @(
                            @{
                                ComputerName = 'PC1'
                                Paths        = @(
                                    @{
                                        Source      = '\\srv\a'
                                        Destination = 'sftp:/a/'
                                    }
                                )
                            }
                        )
                    }
                )
            } | ConvertTo-Json -Depth 7 |
            Out-File (Join-Path $testJsonFolder 'File1.json')

            @{
                Tasks = @(
                    @{
                        TaskName = 'File2 task'
                        Sftp     = @{
                            ComputerName = 'sftp2.com'
                            Port         = 2222
                            Credential   = @{
                                UserName = 'u2'
                                Password = 'p2'
                            }
                        }
                        Option   = @{
                            OverwriteFile       = $false
                            ExcludeZeroSizeFile = $true
                            MatchFileNameRegex  = '\.xml$'
                        }
                        Actions  = @(
                            @{
                                ComputerName = 'PC2'
                                Paths        = @(
                                    @{
                                        Source      = 'sftp:/out/'
                                        Destination = '\\srv\b'
                                    }
                                )
                            }
                        )
                    }
                )
            } | ConvertTo-Json -Depth 7 |
            Out-File (Join-Path $testJsonFolder 'File2.json')

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Path = $testJsonFolder

            Test-NewJsonFileHC

            .$testScript @testParams
        }

        It 'should create an Excel file with rows from both files' {
            $testLogFolder = $testInputFile.Settings.SaveLogFiles.Where.Folder

            $testExcelFile = Get-ChildItem -Path $testLogFolder -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike '*System errors*' }

            $testExcelFile | Should -Not -BeNullOrEmpty

            $excelContent = Import-Excel -Path $testExcelFile.FullName

            @($excelContent).Count | Should -Be 2

            ($excelContent | Where-Object { $_.FileName -eq 'File1' }).TaskName |
            Should -Be 'File1 task'

            ($excelContent | Where-Object { $_.FileName -eq 'File2' }).TaskName |
            Should -Be 'File2 task'
        }

        AfterAll { Remove-TestFileHC }
    }
}

Describe 'SendMail.When' {
    BeforeAll {
        $testJsonFolder = (New-Item 'TestDrive:/jsonMailTest' -ItemType Directory).FullName

        @{
            Tasks = @(
                @{
                    TaskName = 'Mail test'
                    Sftp     = @{
                        ComputerName = 'sftp.test.com'
                        Port         = 22
                        Credential   = @{
                            UserName = 'u'
                            Password = 'p'
                        }
                    }
                    Option   = @{
                        OverwriteFile       = $false
                        ExcludeZeroSizeFile = $false
                        MatchFileNameRegex  = '.*'
                    }
                    Actions  = @(
                        @{
                            ComputerName = $null
                            Paths        = @(
                                @{
                                    Source      = '\\srv\data'
                                    Destination = 'sftp:/data/'
                                }
                            )
                        }
                    )
                }
            )
        } | ConvertTo-Json -Depth 7 |
        Out-File (Join-Path $testJsonFolder 'MailTest.json')
    }

    It "should send mail" {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Path = $testJsonFolder

        Test-NewJsonFileHC

        .$testScript @testParams

        Should -Invoke Send-MailKitMessageHC -Times 1
    }

    AfterAll { Remove-TestFileHC }
}

Describe 'no .json files in the folder' {
    BeforeAll {
        $testEmptyFolder = (New-Item 'TestDrive:/emptyJsonFolder' -ItemType Directory).FullName

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Path = $testEmptyFolder

        Test-NewJsonFileHC
    }

    It 'should create a system error' {
        { .$testScript @testParams } | Should -Not -Throw

        $LASTEXITCODE | Should -Be 1
    }

    AfterAll { Remove-TestFileHC }
}
