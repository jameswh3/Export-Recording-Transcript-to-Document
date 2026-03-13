#Requires -Modules PnP.PowerShell
#Requires -Version 7.0

#region CONFIGURATION
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } `
             elseif ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path } `
             else { $PWD.Path }

$envFile = Join-Path $scriptDir ".env"
if (Test-Path $envFile) {
    Get-Content $envFile | ForEach-Object {
        if ($_ -match '^\s*([^#][^=]+?)\s*=\s*(.*)\s*$') {
            Set-Variable -Name $Matches[1].Trim() -Value $Matches[2].Trim()
        }
    }
} else {
    Write-Warning ".env file not found at '$envFile'. Falling back to values defined in script."
}

$clientId          = if ($CLIENT_ID)          { $CLIENT_ID }          else { $clientId }
$siteUrl           = if ($SITE_URL)           { $SITE_URL }           else { $siteUrl }
$documentLibrary   = if ($DOCUMENT_LIBRARY)   { $DOCUMENT_LIBRARY }   else { "Videos" }
$sharePointFolder  = if ($SHAREPOINT_FOLDER)  { $SHAREPOINT_FOLDER }  else { $null }
$destinationFolder = if ($DESTINATION_FOLDER) { $DESTINATION_FOLDER } else { "C:\Temp" }
$streamEndpoint    = if ($STREAM_ENDPOINT)    { $STREAM_ENDPOINT }    else { "/_layouts/15/stream.aspx" }
$vttSegmentSize    = if ($VTT_SEGMENT_SIZE)   { [int]$VTT_SEGMENT_SIZE } else { 30 }
#endregion

#region HELPER FUNCTIONS

function Convert-ToSeconds {
    param([string]$Time)
    $parts = $Time -split "[:.]"
    return [int]$parts[0] * 3600 + [int]$parts[1] * 60 + [int]$parts[2] + [int]$parts[3] / 1000
}

function Get-WebVTTContent {
    param(
        [string]$VTTFilePath,
        [int]$SegmentSize,
        $Speakers
    )
    $lines = Get-Content -Path $VTTFilePath

    $sentences         = @()
    $currentSentence   = ""
    $currentStartTime  = 0
    $currentEndTime    = ""
    $currentSpeakers   = ""
    $timecodePattern   = "(\d{2}:\d{2}:\d{2}\.\d{3}) --> (\d{2}:\d{2}:\d{2}\.\d{3})"

    foreach ($line in $lines) {
        if ($line -match "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}-\d+" `
            -or $line -match "WebVTT" -or $line -eq "") { continue }

        if ($line -match $timecodePattern) {
            if ($currentSentence -ne "") { $currentEndTime = Convert-ToSeconds $Matches[2] }
            $currentStartTime = Convert-ToSeconds $Matches[1]
            $currentEndTime   = Convert-ToSeconds $Matches[2]
        } elseif ($line -match '[\.\!\?]') {
            if ($line -match '<v\s+([^>]+)>') { $currentSpeakers += $Matches[1] }
            $currentSentence += " " + $line
            $sentences += [PSCustomObject]@{
                Sentence  = ($currentSentence -replace '<[^>]+>', '').Trim()
                StartTime = $currentStartTime
                EndTime   = $currentEndTime
                Speakers  = $Speakers ? $Speakers : $currentSpeakers
            }
            $currentSentence = ""; $currentSpeakers = ""
        } else {
            if ($line -match '<v\s+([^>]+)>') { $currentSpeakers += $Matches[1] }
            $currentSentence += " " + $line
        }
    }

    if ($currentSentence -ne "") {
        $sentences += [PSCustomObject]@{
            Sentence  = ($currentSentence -replace '<[^>]+>', '').Trim()
            StartTime = $currentStartTime
            EndTime   = $currentEndTime
            Speakers  = $Speakers ? $Speakers : $currentSpeakers
        }
    }

    if (-not $SegmentSize) { return $sentences }

    $groupedSentences    = @()
    $currentGroup        = ""
    $currentGroupSpeakers = @()
    $currentStart        = 0
    $currentEnd          = $SegmentSize

    foreach ($sentence in $sentences) {
        if ($sentence.StartTime -ge $currentStart -and $sentence.StartTime -lt $currentEnd) {
            $currentGroup        += " " + $sentence.Sentence
            $currentGroupSpeakers += $sentence.Speakers
        } else {
            if ($currentGroup -ne "") {
                $groupedSentences += [PSCustomObject]@{
                    Sentence  = $currentGroup.Trim()
                    StartTime = $currentStart
                    EndTime   = $currentEnd
                    Speakers  = ($currentGroupSpeakers | Select-Object -Unique)
                }
            }
            $currentStart         = [int]([math]::Floor($sentence.StartTime / $SegmentSize) * $SegmentSize)
            $currentEnd           = [int]($currentStart + $SegmentSize)
            $currentGroup         = $sentence.Sentence
            $currentGroupSpeakers = @()
        }
    }
    if ($currentGroup -ne "") {
        $groupedSentences += [PSCustomObject]@{
            Sentence  = $currentGroup.Trim()
            StartTime = $currentStart
            EndTime   = $currentEnd
            Speakers  = ($currentGroupSpeakers | Select-Object -Unique)
        }
    }
    return $groupedSentences
}

# Downloads the VTT transcript for a single SharePoint file.
# Returns an array of local VTT file paths (one per transcript track).
function Get-FileTranscript {
    param(
        [string]$SiteUrl,
        [string]$DriveId,
        [string]$DestinationFolder,
        [string]$BearerToken,
        [PnP.PowerShell.Commands.Base.PnPConnection]$PnPConnection,
        $File   # PnP file object with UniqueId and Name
    )

    $itemId   = $File.UniqueId
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name) `
                    -replace "-Meeting Recording$", ""

    $transcriptsUrl = "$SiteUrl/_api/v2.1/drives/$DriveId/items/$itemId/media/transcripts"
    $response = Invoke-PnPSPRestMethod -Method Get -Url $transcriptsUrl -Connection $PnPConnection

    if (-not $response.value -or $response.value.Count -eq 0) {
        Write-Warning "  No transcripts found for '$($File.Name)'."
        return @()
    }

    $headers    = @{ "Authorization" = "Bearer $BearerToken" }
    $localPaths = @()
    $i = 1
    foreach ($transcript in $response.value) {
        $outPath = Join-Path $DestinationFolder "$baseName - $i.vtt"
        Invoke-WebRequest -Uri $transcript.temporaryDownloadUrl -OutFile $outPath -Headers $headers
        $localPaths += $outPath
        $i++
    }
    return $localPaths
}

# Writes a Word document for one file's transcript segments.
function New-TranscriptWordDoc {
    param(
        [string]$FileBaseName,
        [string]$FileName,
        [string]$MeetingSubject,
        [string]$FileUrl,
        [string]$SiteUrl,
        [string]$StreamEndpoint,
        [string]$OutputPath,
        $Segments
    )

    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $word.ActiveWindow.WindowState = 2  # wdWindowStateMinimize
    $sel = $word.Selection

    $sel.Style = "Heading 1"
    $sel.TypeText($FileBaseName)
    $sel.TypeParagraph()

    $sel.Style = "Heading 2"
    $sel.TypeText("Overview")
    $sel.TypeParagraph()

    $sel.Style = "Normal"
    $sel.TypeText("Subject:   $MeetingSubject")
    $sel.TypeParagraph()

    $sel.Style = "Normal"
    $sel.TypeText("File Name: $FileName")
    $sel.TypeParagraph()

    $sel.Style = "Heading 2"
    $sel.TypeText("Transcript")
    $sel.TypeParagraph()

    $encodedFileRef = [URI]::EscapeUriString($FileUrl).Replace("/", "%2F")

    foreach ($segment in $Segments) {
        $startSecs = $segment.StartTime
        $endSecs   = $segment.EndTime

        # Timestamp heading
        $sel.Style = "Heading 2"
        $sel.TypeText("$startSecs - $endSecs")
        $sel.TypeParagraph()

        # Speakers (bold)
        $speakerLine = ($segment.Speakers | Select-Object -Unique) -join ", "
        if ($speakerLine) {
            $sel.Style = "Normal"
            $sel.Font.Bold = $true
            $sel.TypeText("Speakers: $speakerLine")
            $sel.Font.Bold = $false
            $sel.TypeParagraph()
        }

        # Transcript text
        $sel.Style = "Normal"
        $sel.TypeText($segment.Sentence.Trim())
        $sel.TypeParagraph()

        # Hyperlink to this segment in Stream
        $sel.Style = "Normal"
        $playbackOptions = [URI]::EscapeUriString("&nav={""playbackOptions"":{""startTimeInSeconds"":$startSecs}}")
        $fullUrl = "$SiteUrl$StreamEndpoint" + "?id=$encodedFileRef" + "$playbackOptions"
        $doc.Hyperlinks.Add($sel.Range, $fullUrl, [Type]::Missing, [Type]::Missing, "View recording segment") | Out-Null
        $sel.MoveRight(1) | Out-Null   # move cursor past the hyperlink field
        $sel.TypeParagraph()
    }

    $outputFile = Join-Path $OutputPath "$FileBaseName.docx"
    $doc.SaveAs2($outputFile)
    $doc.Close()
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

    Write-Host "  Saved: $outputFile"
}

#endregion

#region MAIN

# Connect to SharePoint
Write-Host "Connecting to SharePoint..."
$connection = Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive -ForceAuthentication -ReturnConnection

# One-time setup: drive ID (needed for transcript API) and bearer token
$site    = Get-PnPSite    -Connection $connection -Includes Id
$web     = Get-PnPWeb     -Connection $connection -Includes Id
$library = Get-PnPList    $documentLibrary -Connection $connection -Includes Id

$bytes   = $site.Id.ToByteArray() + $web.Id.ToByteArray() + $library.Id.ToByteArray()
$driveId = "b!" + ([Convert]::ToBase64String($bytes)).Replace('/', '_').Replace('+', '-')

$token = Get-PnPAccessToken -ResourceTypeName SharePoint -Connection $connection

# Enumerate files in the target SharePoint folder
$folderUrl = $documentLibrary
if ($sharePointFolder) { $folderUrl = "$folderUrl/$sharePointFolder" }
Write-Host "Retrieving files from '$folderUrl'..."
$spFiles = Get-PnPFileInFolder -FolderSiteRelativeUrl $folderUrl -Connection $connection `
               -Includes UniqueId, ServerRelativeUrl, Name

# Per-file pipeline: transcript download → VTT parse → Word doc
foreach ($spFile in $spFiles) {
    Write-Host "Processing: $($spFile.Name)"

    # Step 1: Download VTT transcript for this file
    $vttPaths = Get-FileTranscript -SiteUrl $siteUrl -DriveId $driveId `
        -DestinationFolder $destinationFolder -BearerToken $token `
        -PnPConnection $connection -File $spFile

    if (-not $vttPaths) { continue }

    $fileBaseName = [System.IO.Path]::GetFileNameWithoutExtension($spFile.Name) `
                        -replace "-Meeting Recording$", ""

    foreach ($vttPath in $vttPaths) {
        # Step 2: Parse VTT into timed segments
        $segments = Get-WebVTTContent -VTTFilePath $vttPath -SegmentSize $vttSegmentSize

        # Step 3: Write Word doc immediately
        New-TranscriptWordDoc `
            -FileBaseName    $fileBaseName `
            -FileName        $spFile.Name `
            -MeetingSubject  $fileBaseName `
            -FileUrl         $spFile.ServerRelativeUrl `
            -SiteUrl         $siteUrl `
            -StreamEndpoint  $streamEndpoint `
            -OutputPath      $destinationFolder `
            -Segments        $segments
    }
}

Write-Host "Done."
#endregion
