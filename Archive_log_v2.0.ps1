<#
    .SYNOPSIS
        This script is created for archive files and remove older zip file based on certain retention policy
    .CHANE LOG
        2019 - V.1.0 - Creatd by CapGemini 
        2020 - V.1.1 - Modification by Rahhul TRIVEDI - BOOST
        2020 - V.1.2 - BOOST Removed LastWriteTime to overcome MIM
        2020 - V.2.0 - BOOST Patched Delete logic, Compression logic, retructured the functions
                       BOOST Added logger moudle and HTML Module
                       BOOST Added Parameter options
                       BOOST Added Error Handeling
#>

Param(  
[Parameter(Mandatory=$false)]  
[string]$ArchiveConf
)  
##Provide Multiple recpts in "User1@total.com","user2@total.com" format string[] array
$recipients = "Rahul.Trivedi@external.total.com"

Function Total-Log
{
    ##This function is written by Rahul TRIVEDI - BOOST for logging and trapping against the Archival issue
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [Alias('LogPath')]
        [string]$Path='C:\Logs\PowerShellLog.log',
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Error","Warn","Info")]
        [string]$Level="Info",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoClobber
    )

    Begin
    {
        $VerbosePreference = 'Continue'
    }
    Process
    {
       
        if ((Test-Path $Path) -AND $NoClobber) {
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name."
            Return
            }

       
        elseif (!(Test-Path $Path)) {
            Write-Verbose "Creating $Path."
            $NewLogFile = New-Item $Path -Force -ItemType File
            }

        else {
            # Nothing to see here yet.
            }

        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Error $Message
                $LevelText = 'ERROR'
                }
            'Warn' {
                Write-Warning $Message
                $LevelText = 'WARNING'
                }
            'Info' {
                Write-Verbose $Message
                $LevelText = 'INFO'
                }
            }
        
        # Write log entry to $Path
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append
    }
    End
    {
    }
}

Function Create-HTMLTable{ 
    param([array]$Array)      
    $arrHTML = $Array | Select @{Name='Source Location';expression={$_.SourceFolder}}, 
        @{Name='File Status';expression={$_.FilesStatus}},
        @{Name='Compression Status';expression={$_.zipStatus}},
        @{Name='Archive Rentention Status';expression={$_.RetentionStatus}}|ConvertTo-Html -As Table
    $arrHTML[-1] = $arrHTML[-1].ToString().Replace('</body></html>'," ")    
    Return $arrHTML[5..2000]
}


##CapgemeniCode With BOOST Patch
$Style = @"      
    <style>
      body {
        font-family: "Helvetica Neue", Helvetica, sans-serif;
        font-size: 8pt;
        color: #4C607B;
        }
table {
  table-layout: fixed;
  align:center;
  ont-size: "10px";
  font-family: "Helvetica Neue", Helvetica, sans-serif
}

thead th {
  padding: 2px 8px;
  background: #4b97cd;
  color: white;
  text-align: center;
}

th {
  padding: 2px 8px;
  background: Gray;
  color: white;
  text-align: center;
  border: 1px solid gray;
  font-size: "12px";

}

td {
  padding: 2px 8px;
  background: WhiteSmoke;
  text-align: left;
  color: #273746;
  border: 1px solid gray;
}
    </style>
      
"@
# Creating head style and header title
$output = $null
$output = @()
$output += '<html><head></head><body>'
$output += 
#Import hmtl style file
$Style 
$output += "<h3 style='color: #0B2161'>Archive Report On "+$env:COMPUTERNAME+"</h3><h4 style='color: #999999'>"+ (Get-Date).ToString('dddd dd, MMMM yyyy  hh:mm tt')+"</h4>"
$output += '<strong><font color="red">WARNING: </font></strong>'
$output += "Please review attached logs.</br>"
$output += '</br>'

$SCRIPT_PATH = $MyInvocation.MyCommand.Definition
$SCRIPT_NAME = $MyInvocation.MyCommand.Name
$EXECUTION_PATH = $SCRIPT_PATH.Replace($SCRIPT_NAME, "")
$LOGFILE=($MyInvocation.MyCommand.Definition).Replace('ps1','log')
$ZIPPED_LOGFILE = $LOGFILE.replace('.log','.zip')
$DRIVE_NAME=(get-location).Drive.Name 
If(Test-path $ZIPPED_LOGFILE) {Remove-item $ZIPPED_LOGFILE}

if(test-path $LOGFILE){
  $BKP_STMP = [datetime]::now.ToString('yyyy-MM-dd')
  try{
    Move-Item -Path $LOGFILE -Destination $LOGFILE-$BKP_STMP -Force
  }catch{
    Write-Error $_ -ErrorAction Stop
  }
}else{
  New-Item -itemType File -Path $EXECUTION_PATH -Name ($SCRIPT_NAME.Replace('ps1','log'))
}

if($PSBoundParameters.ContainsKey('ArchiveConf')){
   Total-Log -Message "Loading $ArchiveConf" -Level Info -Path $LOGFILE
}else{
    $ArchiveConf = ".\Desktop\Archive_logs_conf.csv"
    Total-Log -Message "Args: -ArchiveConf is not passed setting up default $ArchiveConf" -Level Warn -Path $LOGFILE
}
$resultArray = @()

try{
    $configuration = import-csv $ArchiveConf
    $CurrentDate = Get-Date
    Total-Log -Message "Executing Archival Process......" -Level Info -Path $logfile
    $sourcefolders = (($configuration | Select-String SourcePath) -split " : ")[1] -split ","
    $configuration | foreach {
        $filewritetime = $null
        $success = 'OK'
        Total-Log -Message "Processing $($_.UNCPath)" -Level info -Path $logfile
        $sourcefolder = $_.UNCPath
        [int]$archivedays = $_.Archivedays
        [int]$retentionpolicy = $_.retentionpolicy
        [int]$archivedaysvalue = "-$archivedays"
        [int]$retentionpolicyvalue = "-$retentionpolicy"
        
        $foldername = (get-date).ToString("ddMMyyyyHHmmss")
        $ArchiveDate = $CurrentDate.AddDays($archivedaysvalue)
        $destpath = "$sourcefolder\$foldername"
        $obj = New-Object psobject -Property @{
            SourceFolder = $sourcefolder
            FilesStatus = "Could Not Determine" 
            RetentionStatus = "Could Not determine"
            zipStatus="Could Not Determine"
            }       
        
        
        try{
            $filewritetime = Get-ChildItem -File -Recurse -Path $sourcefolder -Exclude "*.zip" -ErrorAction Stop| where {$_.LastWriteTime -lt $ArchiveDate} | select *
            if($filewritetime -eq $null){
                Total-Log -Message "[Collecting-Files] $sourcefolder has no files older than $ArchiveDate" -Level Warn -Path $logfile 
                $obj.FilesStatus = "0 Files Older than $ArchiveDate"
                $obj.zipStatus = "Skipped.."
                                     
            }else{                
                md "$destpath"-Force -ErrorAction Stop
                Total-Log -Message "[Folder-Creation] $destpath created for moving files older than $ArchiveDate" -Level info -Path $logfile
                if(test-path $destpath){
                    $filewritetime | foreach {
                        $path = $_.FullName
                        Total-Log -Message "Moving $path >>>> $destpath" -Level Info -Path $logfile 
                        Move-Item $path $destpath -Force -ErrorAction Stop
                        }
                    $total_size=[string]$($filewritetime|Measure-Object -Property length -sum).Sum                    
                    $obj.FilesStatus =[string]$($filewritetime|measure).Count + " Files moved to reclaim {0:N2} MB of space" -f ($total_size/1MB)
                    try{
                        $zipfilename = "$sourcefolder\$foldername.zip"
                        Compress-Archive -Path "$destpath" -CompressionLevel Optimal -DestinationPath $zipfilename -ErrorAction Stop
                        Total-Log -Message "[Compression-Folder] Compression Started for $destpath to $zipfilename" -Level Info -path $logfile
                        $zipSize = $(Get-ChildItem $zipfilename -ErrorAction Stop|Select-Object Length).length
                        $obj.zipStatus = "$zipfilename [{0:N2} MB] created, Reclaimed {1:N2} %" -f (($zipsize/1MB),(($total_size.Length - $zipSize.Length) / $total_size.Length) * 100)    
                          
                    }catch{ 
                        $status = 'KO'
                        Total-Log -Message "[Compression-Folder] $_" -Level Error -Path $logfile
                        $obj.zipStatus = "[KO] $_"
                    }
                }else{
                    Total-Log -Message "[Folder-Creation] $destpath does not exist" -Level Info -Path $logfile
                    $obj.FilesStatus = "[KO] $destpath does not exist, file movement failed."                             
                }
          }
        }catch{
            Total-Log -Message "[Collecting-Files] $_" -Level error -Path $logfile
            $obj.Filesstatus = $_
        }
        
        if($retentionpolicyvalue -ne 0){
            Total-Log -Message "[Archive-Retention] $sourcefolder policy found to remove archive with $retentionpolicy Days older, Initiating Archive Deletion" -level info -path $logfile
            $ArchiveretentionDate = $CurrentDate.AddDays($retentionpolicyvalue)
            try{
                $archiveFiles = Get-ChildItem $sourcefolder -Filter *.zip -ErrorAction Stop| where {$_.LastWriteTime -le $ArchiveretentionDate}
                $totalArchive=$archiveFiles| Measure
                Total-Log -Message "[Archive-Retention] $($totalArchive.count) Zip Files with $retentionpolicy Days older" -level info -path $logfile
                foreach($archiveFile in $archiveFiles){ 
                    Total-Log -Message "[Archive-Retention] Removing $archivefile" -Level info -Path $logfile
                    Remove-Item –path $archiveFile.FullName -ErrorAction Stop
                }
                $obj.RetentionStatus = "$($totalArchive.count) Zip Files older than $retentionpolicy days, removed"
            }catch{
                Total-Log -Message "[Archive-Retention] $_" -level ERROR -path $logfile
                $obj.RetentionStatus=$_
            }
        }else{
            Total-Log -Message "[Archive-Retention] $sourcefolder Retention Policy found as Never Delete $retentionpolicy, Skiping Archive Deletion" -level Info -Path $logfile
            $obj.RetentionStatus = "Skipped as retention Policy is Never Delete"
        }
        if($success -eq 'OK'){
            if (Test-Path $destpath){
                try{ 
                    Remove-Item -Path "$destpath" -Recurse -Force
                    Total-Log -Message "[Removing-Folder] Removed $destpath" -Level info -Path $logfile             
                }catch{ 
                    Total-Log -Message "[Removing-Folder] $_" -level Error -Path $logfile
                    $obj.FileStatus += "[KO] $_"
                }
            }
        }
        $resultArray += $obj
        }#foreachend

    $output += Create-HTMLTable $resultArray
    $output += '</p>'
    $output += '</body></html>'
    $output += '<div><p style="font-weight:bold; color:#999;">- BOOST Automation</div>'    
    $output =  $output | Out-String
    Compress-Archive -Path $logfile -Update -DestinationPath $ZIPPED_LOGFILE
    $endTime = Get-Date
    Total-log -Message "The archive process completed $($endTime.ToString('yyyy-MM-ddTHH:mm:ss'))" -Level info -Path $LOGFILE
    #send-mailmessage -from "Rahul.Trivedi@external.Total.com" -to $recipients -subject "[Archival Status] $env:COMPUTERNAME | Archive Status" -BodyAsHtml $output -Attachment $ZIPPED_LOGFILE -smtpServer emeamaicli-el01.main.glb.corp.local
    If(Test-path $ZIPPED_LOGFILE) {Remove-item $ZIPPED_LOGFILE}
}catch{
    Total-Log -Message $_ -Level Error -Path $logfile
    #send-mailmessage -from "Rahul.Trivedi@external.Total.com" -to $recipients -subject "[Failure] $env:COMPUTERNAME | Archive Error" -BodyAsHtml "Hello, <BR><BR> There is an error occurred during archival operation, kindly review the attached logs. <BR><BR> - Rahul TRIVEDI <BR>BOOST Automation" -Attachment $logfile -smtpServer emeamaicli-el01.main.glb.corp.local
}