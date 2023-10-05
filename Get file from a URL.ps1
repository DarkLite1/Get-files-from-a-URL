#Requires -Version 5.1
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.Remoting

<#
    .SYNOPSIS
        Download files from a URL.
        
    .DESCRIPTION
        A folder is scanned for Excel files. Each Excel file contains a 
        worksheet with the column Url, where to download the file from, the
        column FileName, how to name the downloaded file and the column
        DownloadFolderName, where the file will be downloaded.

        Upon execution of this script a download folder is created in the 
        folder where the Excel files are stored. This download folder will 
        contain the downloaded files. For each download folder there will be 
        a zip file.

        A summary mail is sent to the user with an overview of all Excel files.
        
    .PARAMETER ImportFile
        Contains all the parameters for the script

    .PARAMETER MailTo
        E-mail addresses of where to send the summary e-mail

    .PARAMETER DropFolder
        The folder where the Excel files are located. Each Excel file contains 
        a sheet with a row for each file to download.

        Mandatory columns in the Excel sheet are:
        - URL
        - FileName
        - DownloadFolderName

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

        #region Test 7 zip installed
        $7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"
        
        if (-not (Test-Path -Path $7zipPath -PathType 'Leaf')) {
            throw "7 zip file '$7zipPath' not found"
        }
        
        Set-Alias Start-SevenZip $7zipPath
        #endregion

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
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

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
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            Write-EventLog @EventEndParams; Exit
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

        $inputFiles = @()

        foreach ($file in $dropFolderExcelFiles) {
            try {
                $inputFile = [PSCustomObject]@{
                    ExcelFile = @{
                        Item         = $file
                        Content      = @()
                        OutputFolder = $null
                        Error        = $null
                    }
                    Tasks     = @()
                    Error     = $null
                }

                #region Test if file is still present
                if (-not (Test-Path -LiteralPath $inputFile.ExcelFile.Item.FullName -PathType 'Leaf')) {
                    throw "Excel file '$($inputFile.ExcelFile.Item.FullName)' was removed during execution"
                }
                #endregion

                #region Create Excel specific output folder
                try {
                    $params = @{
                        Path        = '{0}\{1} {2}' -f 
                        $outputFolder, $startDate, $inputFile.ExcelFile.Item.BaseName
                        ItemType    = 'Directory' 
                        Force       = $true
                        ErrorAction = 'Stop'
                    }
                    $inputFile.ExcelFile.OutputFolder = (New-Item @params).FullName

                    Write-Verbose "Excel file output folder '$($inputFile.ExcelFile.OutputFolder)'"
                }
                Catch {
                    throw "Failed creating the Excel output folder '$($inputFile.ExcelFile.OutputFolder)': $_"
                }
                #endregion

                try {
                    #region Move original Excel file to output folder
                    try {
                        $moveParams = @{
                            LiteralPath = $inputFile.ExcelFile.Item.FullName
                            Destination = Join-Path $inputFile.ExcelFile.OutputFolder $inputFile.ExcelFile.Item.Name
                            ErrorAction = 'Stop'
                        }

                        Write-Verbose "Move original Excel file '$($moveParams.LiteralPath)' to output folder '$($moveParams.Destination)'"

                        Move-Item @moveParams
                    }
                    catch {
                        $M = $_
                        $error.RemoveAt(0)
                        throw "Failed moving the file '$($inputFile.ExcelFile.Item.FullName)' to folder '$($inputFile.ExcelFile.OutputFolder)': $M"
                    }
                    #endregion
            
                    #region Import Excel file
                    try {
                        $M = "Import Excel file '$($inputFile.ExcelFile.Item.FullName)'"
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
                        $params = @{
                            Path          = $moveParams.Destination
                            WorksheetName = $ExcelFileWorksheetName
                            ErrorAction   = 'Stop'
                            DataOnly      = $true
                        }
                        $inputFile.ExcelFile.Content += Import-Excel @params |
                        Select-Object -Property 'Url', 'FileName', 
                        'DownloadFolderName'
            
                        $M = "Imported {0} rows from Excel file '{1}'" -f
                        $inputFile.ExcelFile.Content.count, $inputFile.ExcelFile.Item.FullName
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                    }
                    catch {
                        $error.RemoveAt(0)
                        throw "Worksheet '$($params.WorksheetName)' not found"
                    }
                    #endregion

                    #region Test Excel file
                    foreach ($row in $inputFile.ExcelFile.Content) {
                        if (-not ($row.FileName)) {
                            throw "Property 'FileName' not found"
                        }
                        if (-not ($row.URL)) {
                            throw "Property 'URL' not found"
                        }
                        if (-not ($row.DownloadFolderName)) {
                            throw "Property 'DownloadFolderName' not found"
                        }
                    }
                    #endregion
                }
                catch {
                    Write-Warning "Excel input file error: $_"
                    $inputFile.ExcelFile.Error = $_
    
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
                    " | Out-File -LiteralPath "$($inputFile.ExcelFile.OutputFolder)\Error.html" -Encoding utf8
                    #endregion
                    
                    $error.RemoveAt(0)
                    Continue
                }

                foreach (
                    $collection in
                    ($inputFile.ExcelFile.Content | 
                    Group-Object -Property 'DownloadFolderName')
                ) {
                    try {
                        $task = [PSCustomObject]@{
                            ItemsToDownload = $collection.Group
                            DownloadFolder  = @{
                                Name = $collection.Name
                                Item = $null
                            }
                            Job             = @{
                                Object = @()
                                Result = @()
                            }
                            FilePath        = @{
                                DownloadResults = Join-Path $inputFile.ExcelFile.OutputFolder 'Download results.xlsx'
                                ZipFile         = Join-Path $inputFile.ExcelFile.OutputFolder "$($collection.Name).zip"
                            }
                            Error           = $null
                        }

                        #region Create download folder
                        try {
                            $params = @{
                                Path        = Join-Path $inputFile.ExcelFile.OutputFolder $task.DownloadFolder.Name
                                ItemType    = 'Directory'
                                ErrorAction = 'Stop'
                            }
                            $task.DownloadFolder.Item = New-Item @params
                        }
                        catch {
                            throw "Failed creating download folder '$($params.Path)': $_"
                        }
                        #endregion

                        #region Download files
                        $M = "Download $($task.ItemsToDownload.count) files to '$($task.DownloadFolder.Item.FullName)'"
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

                        foreach ($row in $task.ItemsToDownload) {
                            Write-Verbose "Download file '$($row.FileName)' from '$($row.Url)'"
                
                            $task.Job.Object += Start-Job -ScriptBlock {
                                Param (
                                    [Parameter(Mandatory)]
                                    [String]$Url,
                                    [Parameter(Mandatory)]
                                    [String]$DownloadFolder,
                                    [Parameter(Mandatory)]
                                    [String]$FileName
                                )
                            
                                try {
                                    $result = [PSCustomObject]@{
                                        Url          = $Url
                                        FileName     = $FileName
                                        Destination  = $null
                                        DownloadedOn = $null
                                        Error        = $null
                                    }

                                    $result.Destination = Join-Path -Path $DownloadFolder -ChildPath $FileName

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
                            } -ArgumentList $row.Url, $task.DownloadFolder.Item.FullName, $row.FileName

                            #region Wait for max running jobs
                            $waitParams = @{
                                Name       = $task.Job.Object | Where-Object { $_ }
                                MaxThreads = $MaxConcurrentJobs
                            }
                            Wait-MaxRunningJobsHC @waitParams
                            #endregion
                        }
                        #endregion

                        #region Wait for jobs to finish
                        $M = "Wait for all $($task.Job.Object.count) jobs to finish"
                        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
     
                        $null = $task.Job.Object | Wait-Job
                        #endregion
     
                        #region Get job results and job errors   
                        $task.Job.Result += $task.Job.Object | Receive-Job
                        #endregion

                        #region Export results to Excel
                        if ($task.Job.Result) {                  
                            $excelParams = @{
                                Path               = $task.FilePath.DownloadResults
                                NoNumberConversion = '*'
                                WorksheetName      = 'Overview'
                                TableName          = 'Overview'
                                AutoSize           = $true
                                FreezeTopRow       = $true
                            }

                            $M = "Export $($task.Job.Result.count) rows to Excel file '$($excelParams.Path)'"
                            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
        
                            $task.Job.Result | Select-Object -Property 'Url', 
                            'FileName', 'Destination', 'DownloadedOn' , 'Error' |
                            Export-Excel @excelParams
                        }
                        #endregion

                        #region Create zip file
                        if (
                        ($task.ItemsToDownload.Count) -eq 
                        ($task.Job.Result.Count) -eq 
                        ($task.Job.Result.where({ $_.DownloadedOn }).count)
                        ) {
                            try {
                                $M = "Create zip file with $($task.Job.Result.count) files in zip file '$($task.FilePath.ZipFile)'"
                                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
    
                                $Source = $task.DownloadFolder.Item.FullName
                                $Target = $task.FilePath.ZipFile
                                Start-SevenZip a -mx=9 $Target $Source

                                if ($LASTEXITCODE -ne 0) {
                                    throw "7 zip failed with last exit code: $LASTEXITCODE"
                                }
                            }
                            catch {
                                $M = $_
                                $Error.RemoveAt(0)
                                throw "Failed creating zip file: $M"
                            }
                        }
                        else {
                            $M = 'Not all files downloaded, no zip file created'
                            Write-Verbose $M; Write-EventLog @EventWarnParams -Message $M

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
                    <h2>No zip-file created because not all files could be downloaded.</h2>
                    </div>

                    <p>Please check the Excel file for more information.</p>

                    </body>
                    </html>
                    " | Out-File -LiteralPath "$($inputFile.ExcelFile.OutputFolder)\Error.html" -Encoding utf8
                            #endregion
                        }
                        #endregion   
                    }
                    catch {
                        $M = $_
                        Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M

                        $task.Error = $_
                        $error.RemoveAt(0)
                    }
                    finally {
                        $inputFile.Tasks += $task   
                    }
                }
            }
            catch {
                $M = $_
                Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M
                    
                $inputFile.Error = $_
                $error.RemoveAt(0)
            }
            finally {
                $inputFiles += $inputFile
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
        if (-not $inputFiles) {
            Write-Verbose "No tasks found, exit script"
            Write-EventLog @EventEndParams; Exit
        }

        # $M = "Wait for all $($inputFile.Job.Object.count) jobs to finish"
        # Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        $mailParams = @{ }
        $htmlTableTasks = @()

        #region Count totals
        $totalCounter = @{
            All          = @{
                Errors          = 0
                RowsInExcel     = 0
                DownloadedFiles = 0
            }
            SystemErrors = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
            
        $totalCounter.All.Errors += $totalCounter.SystemErrors
        #endregion

        foreach ($inputFile in $inputFiles) {
            #region Count task results
            $counter = @{
                RowsInExcel     = (
                    $inputFile.ExcelFile.Content | Measure-Object
                ).Count
                DownloadedFiles = (
                    $inputFile.Tasks.Job.Result.Where({ $_.DownloadedOn }) | Measure-Object
                ).Count
                Errors          = @{
                    InExcelFile      = (
                        $inputFile.ExcelFile.Error | Measure-Object
                    ).Count
                    InTasks          = (
                        $inputFile.Tasks.Where({ $_.Error }) | Measure-Object
                    ).Count
                    DownloadingFiles = (
                        $inputFile.Tasks.Job.Result.Where({ $_.Error }) | Measure-Object
                    ).Count
                    Other            = (
                        $inputFile.Error | Measure-Object
                    ).Count
                }
            }
            
            $totalCounter.All.RowsInExcel += $counter.RowsInExcel
            $totalCounter.All.DownloadedFiles += $counter.DownloadedFiles
            $totalCounter.All.Errors += (
                $counter.Errors.InExcelFile + 
                $counter.Errors.InTasks + 
                $counter.Errors.DownloadingFiles +
                $counter.Errors.Other
            )
            #endregion
                
            #region Create HTML table
            $htmlTableTasks += "
                <table>
                <tr>
                    <th colspan=`"2`">$($inputFile.ExcelFile.Item.Name)</th>
                </tr>
                <tr>
                    <td>Details</td>
                    <td>
                        <a href=`"$($inputFile.ExcelFile.OutputFolder)`">Output folder</a>
                    </td>
                </tr>
                <tr>
                    <td>$($counter.RowsInExcel)</td>
                    <td>Files to download</td>
                </tr>
                <tr>
                    <td>$($counter.DownloadedFiles)</td>
                    <td>Files successfully downloaded</td>
                </tr>
                $(
                    if ($counter.Errors.InExcelFile) {
                        "<tr>
                            <td style=``"background-color: red``">$($counter.Errors.InExcelFile)</td>
                            <td style=``"background-color: red``">Error{0} in the Excel file</td>
                        </tr>" -f $(if ($counter.Errors.InExcelFile -ne 1) {'s'})
                    }
                )
                $(
                    if ($counter.Errors.DownloadingFiles) {
                        "<tr>
                            <td style=``"background-color: red``">$($counter.Errors.DownloadingFiles)</td>
                            <td style=``"background-color: red``">File{0} failed to download</td>
                        </tr>" -f $(if ($counter.Errors.DownloadingFiles -ne 1) {'s'})
                    }
                )
                $(
                    if ($counter.Errors.Other) {
                        "<tr>
                            <td style=``"background-color: red``">$($counter.Errors.Other)</td>
                            <td style=``"background-color: red``">Error{0} found:<br>{1}</td>
                        </tr>" -f $(
                            if ($counter.Errors.Other -ne 1) {'s'}
                        ),
                        (
                            '- ' + $($inputFile.Error -join '<br> - ')
                        )
                    }
                )
                $(
                    if($inputFile.Tasks) {
                        "<tr>
                            <th colspan=``"2``">Downloads per folder</th>
                        </tr>"
                    }
                )
                $(
                    foreach ($task in $inputFile.Tasks) {
                        $errorCount = $task.Job.Result.Where(
                            {$_.Error}).Count

                        $template = if ($errorCount) {
                            "<tr>
                            <td style=``"background-color: red``">{0}/{1}</td>
                            <td style=``"background-color: red``">{2}{3}</td>
                            </tr>"     
                        } else {
                            "<tr>
                                <td>{0}/{1}</td>
                                <td>{2}{3}</td>
                            </tr>" 
                        }

                        $template -f 
                        $(
                            $task.Job.Result.Where({$_.DownloadedOn}).Count
                        ),
                        $(
                            ($task.ItemsToDownload | Measure-Object).Count
                        ),
                        $(
                            $task.DownloadFolder.Name
                        ),
                        $(
                            if ($errorCount) {
                                ' ({0} error{1})' -f 
                                $errorCount, $(if ($errorCount -ne 1) {'s'})
                            }
                        )
                    }
                )
            </table>
            "
            #endregion
        }

        $htmlTableTasks = $htmlTableTasks -join '<br>'

        #region Send summary mail to user

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'
        $mailParams.Subject = '{0}/{1} file{2} downloaded' -f 
        $totalCounter.All.DownloadedFiles,
        $totalCounter.All.RowsInExcel,
        $(
            if ($totalCounter.All.RowsInExcel -ne 1) {
                's'
            }
        )

        if (
            $totalErrorCount = $totalCounter.All.Errors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create error html lists
        $systemErrorsHtmlList = if ($totalCounter.SystemErrors) {
            "<p>Detected <b>{0} system error{1}</b>:{2}</p>" -f $totalCounter.SystemErrors, 
            $(
                if ($totalCounter.SystemErrors -ne 1) { 's' }
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
                $systemErrorsHtmlList
                <p>Summary:</p>
                $htmlTableTasks"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
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