#Requires -Version 5.1
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.Remoting

<#
    .SYNOPSIS
        Download all files defined in an Excel sheet.
        
    .DESCRIPTION
        Each Excel file contains a URL and a FileName field so the script knows
        where to download the files and how to name the files.
        
    .PARAMETER ImportFile
        Contains all the parameters for the script

    .PARAMETER MailTo
        E-mail addresses of where to send the summary e-mail

    .PARAMETER DropFolder
        The folder where the Excel files are located. Each Excel file contains 
        a sheet with the delivery notes to download.

        Mandatory fields in the Excel sheet are:
        - FileName
        - URL

    .PARAMETER ExcelFileWorksheetName
        The name of the Excel worksheet where the download details are stored

    .PARAMETER MaxConcurrentJobs
        Amount of web requests that are made at the same time
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\Alpha\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        $startDate = (Get-ScriptRuntimeHC -Start).ToString('yyyy-MM-dd HHmmss')
        
        $Error.Clear()
        
        #region Logging
        try {
            $joinParams = @{
                Path        = $LogFolder 
                ChildPath   = $startDate 
                ErrorAction = 'Ignore'
            }
            
            $logParams = @{
                LogFolder    = New-Item -Path (Join-Path @joinParams) -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @logParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop -Encoding UTF8 | 
        ConvertFrom-Json
        #endregion

        #region Test .json file properties
        Try {
            if (-not ($MailTo = $file.MailTo)) {
                throw "Property 'MailTo' not found"
            }
            if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
                throw "Property 'MaxConcurrentJobs' not found"
            }
            if (-not ($DropFolder = $file.DropFolder)) {
                throw "Property 'DropFolder' not found"
            }
            if (-not ($ExcelFileWorksheetName = $file.ExcelFileWorksheetName)) {
                throw "Property 'ExcelFileWorksheetName' not found"
            }
            if (-not ($file.MaxConcurrentJobs -is [int])) {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$($file.MaxConcurrentJobs)' is not supported."
            }
            if (-not (Test-Path -LiteralPath $DropFolder -PathType Container)) {
                throw "Property 'DropFolder': Path '$DropFolder' not found"
            }
        }
        Catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Get Excel files in drop folder
        $params = @{
            LiteralPath = $DropFolder
            Filter      = '*.xlsx'
            ErrorAction = 'Stop'
        }
        $dropFolderExcelFiles = Get-ChildItem @params

        if (-not $dropFolderExcelFiles) {
            $M = "No Excel files found in drop folder '$DropFolder'"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            Write-EventLog @EventEndParams

            Exit
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Create general output folder
        $outputFolder = Join-Path -Path $DropFolder -ChildPath 'Output'

        $null = New-Item -Path $outputFolder -ItemType Directory -EA Ignore
        #endregion

        $tasks = @()

        foreach ($file in $dropFolderExcelFiles) {
            #region Test if file is still there
            if (-not (Test-Path -LiteralPath $file.FullName -PathType Leaf)) {
                $M = "Excel file '$($file.FullName)' removed in the meantime"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
                
                Continue
            }
            #endregion

            #region Create Excel specific output folder
            try {
                $excelFileOutputFolder = '{0}\{1} {2}' -f 
                $outputFolder, $startDate, $file.BaseName
                
                $null = New-Item -Path $excelFileOutputFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'

                Write-Verbose "Excel file output folder '$excelFileOutputFolder'"
            }
            Catch {
                throw "Failed creating the Excel output folder '$excelFileOutputFolder': $_"
            }
            #endregion

            #region Copy original Excel file to output folder
            $moveParams = @{
                LiteralPath = $file.FullName
                Destination = '{0}\{1}' -f $excelFileOutputFolder, $file.Name
            }

            Write-Verbose "Move original Excel file '$($moveParams.LiteralPath)' to output folder '$($moveParams.Destination)'"

            Move-Item @moveParams
            #endregion

            try {
                $task = [PSCustomObject]@{
                    Jobs      = @()
                    ExcelFile = @{
                        Item         = $file
                        Content      = @()
                        OutputFolder = $excelFileOutputFolder
                        Error        = $null
                    }
                }

                try {
                    #region Import Excel file
                    try {
                        $M = "Import Excel file '$($task.ExcelFile.Item.FullName)'"
                        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
                        $params = @{
                            Path          = $moveParams.Destination
                            WorksheetName = $ExcelFileWorksheetName
                            ErrorAction   = 'Stop'
                            DataOnly      = $true
                        }
                        $task.ExcelFile.Content += Import-Excel @params |
                        Select-Object -Property * -ExcludeProperty 'Error', 
                        'DownloadedOn'
            
                        $M = "Imported {0} rows from Excel file '{1}'" -f
                        $task.ExcelFile.Content.count, $task.ExcelFile.Item.FullName
                        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
                    }
                    catch {
                        $error.RemoveAt(0)
                        throw "Worksheet '$($params.WorksheetName)' not found"
                    }
                    #endregion

                    #region Test Excel file
                    foreach ($row in $task.ExcelFile.Content) {
                        if (-not ($row.FileName)) {
                            throw "Property 'FileName' not found"
                        }
                        if (-not ($row.URL)) {
                            throw "Property 'URL' not found"
                        }
                    }
                    #endregion
                }
                catch {
                    Write-Warning "Excel input file error: $_"
                    $task.ExcelFile.Error = $_
    
                    #region Create Error.html file                    
                    "
                    <!DOCTYPE html>
                    <html>
                    <head>
                    <style>
                    .myDiv {
                    border: 5px outset red;
                    background-color: lightblue;    
                    text-align: center;
                    }
                    </style>
                    </head>
                    <body>

                    <h1>Error detected in the Excel sheet</h1>

                    <div class=`"myDiv`">
                    <h2>$_</h2>
                    </div>

                    <p>Please fix this error and try again.</p>

                    </body>
                    </html>
                    " | Out-File -LiteralPath "$($task.ExcelFile.OutputFolder)\Error.html" -Encoding utf8
                    #endregion
                    
                    $error.RemoveAt(0)
                    Continue
                }

                #region Create download folder
                $downloadFolder = New-Item -Path $task.ExcelFile.OutputFolder -Name 'PDF files' -ItemType Directory
                #endregion

                #region Create Excel objects
                foreach ($row in $excelFile) {
                    $row | Add-Member -NotePropertyMembers @{
                        Destination  = Join-Path $downloadFolder $row.FileName
                        DownloadedOn = $null
                        Error        = $null
                    }
                }
                #endregion

                #region Download files
                $M = "Download $($excelFile.count) delivery notes"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M


                $jobs = @()

                foreach ($row in $excelFile) {
                    Write-Verbose "Download file '$($row.Url)' to '$($row.Destination)'"
                
                    $jobs += Start-Job -ScriptBlock {
                        try {
                            $result = $using:row

                            $invokeParams = @{
                                Uri         = $result.Url 
                                OutFile     = $result.Destination 
                                TimeoutSec  = 10 
                                ErrorAction = 'Stop'
                            }
                            Invoke-WebRequest @invokeParams
                        
                            $result.DownloadedOn = Get-Date   
                        }
                        catch {
                            $statusCode = $_.Exception.Response.StatusCode.value__

                            if ($statusCode) {
                                $errorMessage = switch ($statusCode) {
                                    '404' { 
                                        'Status code: 404 Not found'; break
                                    }
                                    Default {
                                        "Status code: $_"
                                    }
                                }
                            }
                            else {
                                $errorMessage = $_
                            }
                    
                            $result.Error = "Download failed: $errorMessage"
                            $Error.RemoveAt(0)
                        }
                        finally {
                            $result
                        }
                    }

                    #region Wait for max running jobs
                    $waitParams = @{
                        Name       = $jobs | Where-Object { $_ }
                        MaxThreads = $MaxConcurrentJobs
                    }
                    Wait-MaxRunningJobsHC @waitParams
                    #endregion
                }
                #endregion

                #region Wait for jobs to finish
                $M = "Wait for all $($jobs.count) jobs to finish"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
     
                $null = $jobs | Wait-Job
                #endregion
     
                #region Get job results and job errors   
                $jobResults = $jobs | Receive-Job
                #endregion
            }
            catch {
    
            }
            finally {
                $tasks += $task
            }
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    try {
        $mailParams = @{ }

        #region Export results to Excel
        $excelParams = @{
            Path               = $logFile + ' - Log.xlsx'
            NoNumberConversion = '*'
            WorksheetName      = 'Overview'
            TableName          = 'Overview'
            AutoSize           = $true
            FreezeTopRow       = $true
        }

        $M = "Export $($jobResults.count) rows to Excel file '$($excelParams.Path)'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        
        $jobResults | 
        Select-Object -Property * -ExcludeProperty 'PSShowComputerName', 
        'RunspaceId', 'PSComputerName' |
        Export-Excel @excelParams

        $mailParams.Attachments = $excelParams.Path
        #endregion

        #region Send mail to user

        #region Error counters
        $counter = @{
            RowsInExcel     = (
                $excelFile | Measure-Object
            ).Count
            DownloadedFiles = (
                $jobResults.Where({ $_.DownloadedOn }) | Measure-Object
            ).Count
            DownloadErrors  = (
                $jobResults.Where({ $_.Error }) | Measure-Object
            ).Count
            SystemErrors    = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0}/{1} file{2} downloaded' -f 
        $counter.DownloadedFiles,
        $counter.RowsInExcel,
        $(
            if ($counter.RowsInExcel -ne 1) {
                's'
            }
        )

        if (
            $totalErrorCount = $counter.DownloadErrors + $counter.SystemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create error html lists
        $SystemErrorsHtmlList = if ($counter.SystemErrors) {
            "<p>Detected <b>{0} non terminating error{1}</b>:{2}</p>" -f $counter.SystemErrors, 
            $(
                if ($counter.SystemErrors -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } | 
                ConvertTo-HtmlListHC
            )
        }
        #endregion
        
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $SystemErrorsHtmlList
                <p>Summary:</p>
                <table>
                    <tr>
                        <td>$($counter.RowsInExcel)</td>
                        <td>Files to download</td>
                    </tr>
                    <tr>
                        <td>$($counter.DownloadedFiles)</td>
                        <td>Files successfully downloaded</td>
                    </tr>
                    <tr>
                        <td>$($counter.DownloadErrors)</td>
                        <td>Errors while downloading files</td>
                    </tr>
                </table>"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message += 
            "<p><i>* Check the attachment for details</i></p>"
        }
   
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}