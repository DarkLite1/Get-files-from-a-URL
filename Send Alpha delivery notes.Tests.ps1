#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {    
    $testInputFile = @{
        MailTo                 = 'bob@contoso.com'
        DropFolder             = (New-Item "TestDrive:/Get files" -ItemType Directory).FullName
        ExcelFileWorksheetName = 'FilesToDownload'
        MaxConcurrentJobs      = 5
    }

    $testExcel = @{
        FilePath    = Join-Path $testInputFile.DropFolder 'File.xlsx'
        FileContent = @(
            [PSCustomObject]@{
                FileName = 'File1.pdf'
                Url      = 'http://something/1'
            }
            [PSCustomObject]@{
                FileName = 'File2.pdf'
                Url      = 'http://something/2'
            }
        )
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
    Mock Invoke-WebRequest
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
                'FileName', 'URL'
            ) {
                $testNewExcel = Copy-ObjectHC $testExcel

                $testNewExcel.FileContent[0].$_ = $null
                
                $testNewExcel.FileContent | Export-Excel -Path $testExcel.FilePath -WorksheetName $testInputFile.ExcelFileWorksheetName

                .$testScript @testParams

                $testErrorFile = Get-ChildItem -Path $testInputFile.DropFolder -Filter 'Error.html' -Recurse

                Get-Content -Path $testErrorFile.FullName -Raw | 
                Should -BeLike "*Property '$_' not found*"
            } 
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        $testExcel.FileContent | Export-Excel -Path $testExcel.FilePath -WorksheetName $testInputFile.ExcelFileWorksheetName

        .$testScript @testParams

        $testExcelFileOutputFolder = Get-ChildItem -Path "$($testInputFile.DropFolder)\Output" -Filter '* File' -Directory
    }
    It 'create an Excel file specific output folder in the DropFolder' {
        $testExcelFileOutputFolder.FullName | Should -Exist
    } -Tag test
    It 'Move the original Excel file to the output folder' {
        Get-ChildItem -Path $testInputFile.DropFolder -File | 
        Should -BeNullOrEmpty

        "$($testExcelFileOutputFolder.FullName)\File.xlsx" | 
        Should -Exist
    } -Tag test


    It 'create a download folder in the log folder' {
        Join-Path $testParams.LogFolder 'PDF Files' | 
        Should -Exist
    }
    It 'download the delivery notes' {
        Should -Invoke Invoke-WebRequest -Times $testExcel.FileContent.Count -Exactly -Scope Describe

        $testExcel.FileContent | ForEach-Object {
            Should -Invoke Invoke-WebRequest -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($Uri -eq $_.Url) -and
                ($OutFile -like "*$($_.Destination)")
            }
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                [PSCustomObject]@{
                    FileName     = 'File1.pdf'
                    Destination  = '*File1.pdf'
                    Url          = 'http://something/1'
                    DownloadedOn = Get-Date
                    Error        = $null
                }
                [PSCustomObject]@{
                    FileName     = 'File2.pdf'
                    Destination  = '*File2.pdf'
                    Url          = 'http://something/2'
                    DownloadedOn = Get-Date
                    Error        = $null
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

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
                $actualRow.DownloadedOn.ToString('yyyyMMdd') | 
                Should -Be $testRow.DownloadedOn.ToString('yyyyMMdd')
                $actualRow.Destination | Should -BeLike $testRow.Destination
                $actualRow.FileName | Should -Be $testRow.FileName
                $actualRow.Url | Should -Be $testRow.Url
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'send an e-mail' {
        It 'with attachment to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '2/2 files downloaded') -and
            ($Attachments -like '*- Log.xlsx') -and
            ($Message -like "*table*2*Files to download*2*Files successfully downloaded<*0*Errors while downloading files*")
            }
        }
    }
}
