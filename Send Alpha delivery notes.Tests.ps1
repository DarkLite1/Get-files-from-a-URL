#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testExcel = @{
        FilePath    = 'TestDrive:/DeliveryNotes.xlsx'
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
    
    $testInputFile = @{
        MailTo        = 'bob@contoso.com'
        DeliveryNotes = @{
            ExcelFiles = @($testExcel.FilePath)
        }
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName         = 'Test (Brecht)'
        ImportFile         = $testOutParams.FilePath
        ExcelWorksheetName = 'Tickets'
        LogFolder          = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin        = 'admin@contoso.com'
    }
    
    $testExcel.FileContent | Export-Excel -Path $testExcel.FilePath -WorksheetName $testParams.ExcelWorksheetName

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
                'MailTo'
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
            It 'DeliveryNotes.<_> not found' -ForEach @(
                'ExcelFiles'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.DeliveryNotes.$_ = $null
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'DeliveryNotes.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'DeliveryNotes.ExcelFiles path not found' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.DeliveryNotes.ExcelFiles = @(
                    'TestDrive:/notFound.xlsx'
                )
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams
                    
                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*$ImportFile*Property 'DeliveryNotes.ExcelFiles': Path 'TestDrive:/notFound.xlsx' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
    Context 'the Excel file' {
        It 'is missing the sheet defined in ExcelWorksheetName' {   
            $testInputFile | ConvertTo-Json -Depth 5 | 
            Out-File @testOutParams

            $testNewParams = $testParams.Clone()
            $testNewParams.ExcelWorksheetName = 'wrong'

            .$testScript @testNewParams
                
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "Excel file '*.xlsx' does not contain worksheet '$($testNewParams.ExcelWorksheetName)'*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        } -Tag test
        Context 'is missing property' {
            It '<_>' -ForEach @(
                'FileName', 'URL'
            ) {
                $testNewExcel = Copy-ObjectHC $testExcel

                $testNewExcel.FilePath = 'TestDrive:/DeliveryNotes2.xlsx'
                $testNewExcel.FileContent = $testNewExcel.FileContent[0] 

                $testNewExcel.FileContent.$_ = $null
                
                $testExportParams = @{
                    Path          = $testNewExcel.FilePath
                    WorksheetName = $testParams.ExcelWorksheetName
                    ClearSheet    = $true
                    Verbose       = $false
                }
                $testNewExcel.FileContent | Export-Excel @testExportParams

                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.DeliveryNotes.ExcelFiles = $testNewExcel.FilePath
    
                $testNewInputFile | ConvertTo-Json -Depth 5 | 
                Out-File @testOutParams

                .$testScript @testParams
                    
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and 
                        ($Message -like "*Excel file '*DeliveryNotes2.xlsx': Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testInputFile | ConvertTo-Json -Depth 5 | 
        Out-File @testOutParams

        .$testScript @testParams
    }

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
