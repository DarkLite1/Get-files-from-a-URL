#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $realCmdLet = @{
        StartJob = Get-Command Start-Job
    }

    $testInputFile = @{
        MailTo                 = 'bob@contoso.com'
        DropFolder             = (New-Item "TestDrive:/Get files" -ItemType Directory).FullName
        ExcelFileWorksheetName = 'FilesToDownload'
        MaxConcurrentJobs      = 5
    }

    $testData = @(
        [PSCustomObject]@{
            Url                = 'http://something/1'
            FileName           = 'File1.pdf'
            FilePath           = '*\{0}\{1}' -f 'Supplier A', 'File1.pdf'
            DownloadFolderName = 'Supplier A'
            DownloadedOn       = $null
            Error              = 'Download failed:*'
        }
        [PSCustomObject]@{
            Url                = 'http://something/2'
            FileName           = 'File2.pdf'
            FilePath           = '*\{0}\{1}' -f 'Supplier A', 'File2.pdf'
            DownloadFolderName = 'Supplier A'
            DownloadedOn       = $null
            Error              = 'Download failed:*'
        }
    )

    $testExcel = @{
        FilePath    = Join-Path $testInputFile.DropFolder 'File.xlsx'
        FileContent = $testData |
        Select-Object 'Url', 'FileName', 'DownloadFolderName'
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
    Mock Start-Job {
        & $realCmdLet.StartJob -Scriptblock {
            $using:testData[0]
        }
    } -ParameterFilter {
        ($ArgumentList[0] -eq $testData[0].Url) -and
        ($ArgumentList[0] -eq $testData[0].DownloadFolderName) -and
        ($ArgumentList[0] -eq $testData[0].FileName)
    }
    Mock Start-Job {
        & $realCmdLet.StartJob -Scriptblock {
            $using:testData[1]
        }
    } -ParameterFilter {
        ($ArgumentList[0] -eq $testData[1].Url) -and
        ($ArgumentList[0] -eq $testData[1].DownloadFolderName) -and
        ($ArgumentList[0] -eq $testData[1].FileName)
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MailTo',
                'MaxConcurrentJobs',
                'DropFolder',
                'ExcelFileWorksheetName'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'MaxConcurrentJobs is not a number' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.MaxConcurrentJobs = 'a'

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'MaxConcurrentJobs' needs to be a number, the value 'a' is not supported*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'DropFolder path not found' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.DropFolder = 'TestDrive:/notFound'

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'DropFolder': Path 'TestDrive:/notFound' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'an Error.html file is saved in the Excel file output folder when' {
    BeforeEach {
        Remove-Item "$($testInputFile.DropFolder)\*" -Recurse -ErrorAction Ignore

        $testInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams
    }
    Context 'the Excel file' {
        It 'is missing the sheet defined in ExcelFileWorksheetName' {
            $testExcel.FileContent |
            Export-Excel -Path $testExcel.FilePath -WorksheetName 'wrong'

            .$testScript @testParams

            $testErrorFile = Get-ChildItem -Path $testInputFile.DropFolder -Filter 'Error.html' -Recurse

            Get-Content -Path $testErrorFile.FullName -Raw |
            Should -BeLike "*Worksheet '$($testInputFile.ExcelFileWorksheetName)' not found*"
        }
        Context 'is missing property' {
            It '<_>' -ForEach @(
                'FileName', 'URL', 'DownloadFolderName'
            ) {
                $testNewExcel = Copy-ObjectHC $testExcel

                $testNewExcel.FileContent[0].$_ = $null

                $testNewExcel.FileContent | Export-Excel -Path $testExcel.FilePath -WorksheetName $testInputFile.ExcelFileWorksheetName

                .$testScript @testParams

                $testErrorFile = Get-ChildItem -Path $testInputFile.DropFolder -Filter 'Error.html' -Recurse

                Get-Content -Path $testErrorFile.FullName -Raw |
                Should -BeLike "*Property '$_' not found*"
            } -Tag test
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        Mock Wait-MaxRunningJobsHC

        $testInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        $testExcel.FileContent | Export-Excel -Path $testExcel.FilePath -WorksheetName $testInputFile.ExcelFileWorksheetName

        .$testScript @testParams

        $testExcelFileOutputFolder = Get-ChildItem -Path "$($testInputFile.DropFolder)\Output" -Filter '* File' -Directory
    }
    It 'create an Excel file specific output folder in the DropFolder' {
        $testExcelFileOutputFolder.FullName | Should -Exist
    }
    It 'Move the original Excel file to the output folder' {
        Get-ChildItem -Path $testInputFile.DropFolder -File |
        Should -BeNullOrEmpty

        "$($testExcelFileOutputFolder.FullName)\File.xlsx" |
        Should -Exist
    }
    It "create the folder 'DownloadedFolderName' in output folder'" {
        $testExcel.FileContent.DownloadedFolderName | ForEach-Object {
            Join-Path $testExcelFileOutputFolder.FullName $_ |
            Should -Exist
        }
    }
    It 'download the files' {
        Should -Invoke Wait-MaxRunningJobsHC -Times $testExcel.FileContent.Count -Exactly -Scope Describe
    }
    It 'when not all files are downloaded Error.html is created in the output folder' {
        $testErrorFile = Get-ChildItem -Path $testExcelFileOutputFolder.FullName -Filter 'Error - *.html' -Recurse

        Get-Content -Path $testErrorFile.FullName -Raw |
        Should -BeLike "*No zip-file created*"
    }
    Context 'export an Excel file to the output folder' {
        BeforeAll {
            $testExportedExcelRows = $testData

            $testExcelLogFile = Get-ChildItem $testExcelFileOutputFolder.FullName -Filter 'Download results.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.Url -eq $testRow.Url
                }
                $actualRow.FileName | Should -Be $testRow.FileName
                $actualRow.FilePath | Should -BeLike $testRow.FilePath
                $actualRow.DownloadedOn | Should -Be $testRow.DownloadedOn
                $actualRow.Error | Should -BeLike $testRow.Error
                $actualRow.DownloadFolderName |
                Should -Be $testRow.DownloadFolderName
            }
        }
    }
    Context 'send an e-mail' {
        It 'to the user with a summary of all Excel files' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '0/2 files downloaded, 2 errors') -and
            ($Message -like "*table*2*Files to download*0*Files successfully downloaded<*2*Files failed to download*")
            }
        }
    } -Skip
}