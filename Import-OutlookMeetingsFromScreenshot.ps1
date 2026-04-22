[CmdletBinding()]
param(
    [string[]]$ScreenshotPaths,

    [switch]$FromClipboard,

    [string]$ClipboardImagePath,

    [string]$ConfigPath,

    [datetime]$WeekStart,

    [string]$ExportJsonPath,

    [string]$DebugOutputDir,

    [ValidateSet("Auto", "OpenAI", "Grid", "Line")]
    [string]$ParseMode = "Auto",

    [int]$TimeSnapMinutes = 30,

    [switch]$WriteObsidian,

    [switch]$CreateDailyNotes
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Initialize-WinRt {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime

    [void][Windows.Storage.StorageFile, Windows, ContentType = WindowsRuntime]
    [void][Windows.Storage.Streams.IRandomAccessStream, Windows, ContentType = WindowsRuntime]
    [void][Windows.Graphics.Imaging.BitmapDecoder, Windows, ContentType = WindowsRuntime]
    [void][Windows.Graphics.Imaging.SoftwareBitmap, Windows, ContentType = WindowsRuntime]
    [void][Windows.Media.Ocr.OcrEngine, Windows, ContentType = WindowsRuntime]
    [void][Windows.Media.Ocr.OcrResult, Windows, ContentType = WindowsRuntime]

    $script:AsTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq "AsTask" -and
            $_.IsGenericMethod -and
            $_.GetParameters().Count -eq 1
        } |
        Select-Object -First 1

    $script:AsTaskNonGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() |
        Where-Object {
            $_.Name -eq "AsTask" -and
            -not $_.IsGenericMethod -and
            $_.GetParameters().Count -eq 1
        } |
        Select-Object -First 1

    if (-not $script:AsTaskGeneric -or -not $script:AsTaskNonGeneric) {
        throw "Failed to initialize WinRT task helpers."
    }
}

function Wait-WinRt {
    param(
        [Parameter(Mandatory = $true)]
        $Operation,

        [Type]$ResultType
    )

    if ($null -eq $ResultType) {
        $task = $script:AsTaskNonGeneric.Invoke($null, @($Operation))
    }
    else {
        $task = $script:AsTaskGeneric.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    }

    return $task.GetAwaiter().GetResult()
}

function Get-EnvironmentVariableValue {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    $targets = @(
        [System.EnvironmentVariableTarget]::Process,
        [System.EnvironmentVariableTarget]::User,
        [System.EnvironmentVariableTarget]::Machine
    )

    foreach ($name in $Names) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        foreach ($target in $targets) {
            $value = [Environment]::GetEnvironmentVariable($name, $target)
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value.Trim()
            }
        }
    }

    return $null
}

function Invoke-InStaRunspace {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [object[]]$ArgumentList = @()
    )

    if ([Threading.Thread]::CurrentThread.GetApartmentState() -eq [Threading.ApartmentState]::STA) {
        return & $ScriptBlock @ArgumentList
    }

    $runspace = [Runspaces.RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions = [Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()

    try {
        $powerShell = [PowerShell]::Create()
        $powerShell.Runspace = $runspace
        [void]$powerShell.AddScript($ScriptBlock.ToString())
        foreach ($argument in $ArgumentList) {
            [void]$powerShell.AddArgument($argument)
        }

        $result = $powerShell.Invoke()
        if ($powerShell.Streams.Error.Count -gt 0) {
            throw $powerShell.Streams.Error[0].Exception
        }

        return $result
    }
    finally {
        if ($powerShell) {
            $powerShell.Dispose()
        }
        $runspace.Dispose()
    }
}

function ConvertTo-NormalizedText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormKC)
    $normalized = $normalized -replace "[\u2010\u2011\u2012\u2013\u2014\u2015\u30FC]", "-"
    $normalized = $normalized -replace "[\uFF1A\uFE55\u2236\uFE13]", ":"
    $normalized = $normalized -replace "[\u301C\uFF5E]", "~"
    $normalized = $normalized -replace "\s+", " "
    $normalized = [regex]::Replace($normalized, "(?<=[\u3040-\u30FF\u4E00-\u9FFF])\s+(?=[\u3040-\u30FF\u4E00-\u9FFF])", "")
    $normalized = [regex]::Replace($normalized, "(?<=[0-9])\s*:\s*(?=[0-9])", ":")
    $normalized = [regex]::Replace($normalized, "(?<=[0-9])\s*/\s*(?=[0-9])", "/")
    $normalized = [regex]::Replace($normalized, "(?<=[0-9])\s*-\s*(?=[0-9])", "-")
    return $normalized.Trim()
}

function ConvertTo-TimeText {
    param(
        [AllowNull()]
        [string]$Token
    )

    if ([string]::IsNullOrWhiteSpace($Token)) {
        return $null
    }

    $value = ConvertTo-NormalizedText -Text $Token
    $match = [regex]::Match($value, "^(?<hour>\d{1,2})(?::(?<minute>\d{2}))?$")
    if (-not $match.Success) {
        return $null
    }

    $hour = [int]$match.Groups["hour"].Value
    $minute = if ($match.Groups["minute"].Success) { [int]$match.Groups["minute"].Value } else { 0 }

    if ($hour -gt 23 -or $minute -gt 59) {
        return $null
    }

    return "{0:D2}:{1:D2}" -f $hour, $minute
}

function ConvertTo-Minutes {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TimeText
    )

    $parts = $TimeText.Split(":")
    return ([int]$parts[0] * 60) + [int]$parts[1]
}

function ConvertTo-TimeStringFromMinutes {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Minutes
    )

    $safeMinutes = [math]::Max(0, $Minutes)
    $hours = [int][math]::Floor($safeMinutes / 60)
    $mins = [int]($safeMinutes % 60)
    return "{0:D2}:{1:D2}" -f $hours, $mins
}

function Snap-Minutes {
    param(
        [int]$Minutes,
        [int]$Step,
        [ValidateSet("Floor", "Ceiling", "Nearest")]
        [string]$Mode
    )

    if ($Step -le 0) {
        return $Minutes
    }

    switch ($Mode) {
        "Floor" {
            return [int]([math]::Floor($Minutes / $Step) * $Step)
        }
        "Ceiling" {
            return [int]([math]::Ceiling($Minutes / $Step) * $Step)
        }
        default {
            return [int]([math]::Round($Minutes / $Step) * $Step)
        }
    }
}

function Get-DateCandidateByDayOfMonth {
    param(
        [int]$DayOfMonth,
        [datetime]$ReferenceDate
    )

    if ($DayOfMonth -lt 1 -or $DayOfMonth -gt 31) {
        return $null
    }

    $baseDate = if ($ReferenceDate -ne [datetime]::MinValue) { $ReferenceDate.Date } else { (Get-Date).Date }
    $candidates = New-Object System.Collections.Generic.List[datetime]

    foreach ($offset in -1, 0, 1) {
        $candidateMonth = $baseDate.AddMonths($offset)
        try {
            $candidate = [datetime]::new($candidateMonth.Year, $candidateMonth.Month, $DayOfMonth)
            $candidates.Add($candidate)
        }
        catch {
        }
    }

    if ($candidates.Count -eq 0) {
        return $null
    }

    return $candidates |
        Sort-Object { [math]::Abs(($_ - $baseDate).Days) } |
        Select-Object -First 1
}

function Get-DateCandidate {
    param(
        [int]$Year,
        [int]$Month,
        [int]$Day,
        [datetime]$ReferenceDate
    )

    try {
        $candidate = [datetime]::new($Year, $Month, $Day)
    }
    catch {
        return $null
    }

    if ($ReferenceDate -ne [datetime]::MinValue) {
        if (($candidate - $ReferenceDate.Date).Days -gt 180) {
            $candidate = $candidate.AddYears(-1)
        }
        elseif (($candidate - $ReferenceDate.Date).Days -lt -180) {
            $candidate = $candidate.AddYears(1)
        }
    }

    return $candidate
}

function Try-ParseDateFromText {
    param(
        [string]$Text,
        [datetime]$ReferenceDate
    )

    $normalized = ConvertTo-NormalizedText -Text $Text
    if (-not $normalized) {
        return $null
    }

    $baseDate = if ($ReferenceDate -ne [datetime]::MinValue) { $ReferenceDate.Date } else { (Get-Date).Date }
    $patterns = @(
        "(?<matched>(?<year>\d{4})/(?<month>\d{1,2})/(?<day>\d{1,2}))",
        "(?<matched>(?<month>\d{1,2})/(?<day>\d{1,2}))"
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($normalized, $pattern)
        if (-not $match.Success) {
            continue
        }

        if ($match.Groups["year"].Success) {
            $date = Get-DateCandidate -Year ([int]$match.Groups["year"].Value) -Month ([int]$match.Groups["month"].Value) -Day ([int]$match.Groups["day"].Value) -ReferenceDate $baseDate
        }
        else {
            $date = Get-DateCandidate -Year $baseDate.Year -Month ([int]$match.Groups["month"].Value) -Day ([int]$match.Groups["day"].Value) -ReferenceDate $baseDate
        }

        if ($date) {
            return [pscustomobject]@{
                Date        = $date
                MatchedText = $match.Groups["matched"].Value
            }
        }
    }

    return $null
}

function Get-OcrResult {
    param(
        [string]$ImagePath
    )

    $resolvedPath = (Resolve-Path -LiteralPath $ImagePath).Path
    $file = Wait-WinRt -Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($resolvedPath)) -ResultType ([Windows.Storage.StorageFile])
    $stream = Wait-WinRt -Operation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) -ResultType ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Wait-WinRt -Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) -ResultType ([Windows.Graphics.Imaging.BitmapDecoder])
    $softwareBitmap = Wait-WinRt -Operation ($decoder.GetSoftwareBitmapAsync()) -ResultType ([Windows.Graphics.Imaging.SoftwareBitmap])

    if ($softwareBitmap.BitmapPixelFormat -ne [Windows.Graphics.Imaging.BitmapPixelFormat]::Gray8 -or
        $softwareBitmap.BitmapAlphaMode -ne [Windows.Graphics.Imaging.BitmapAlphaMode]::Ignore) {
        $softwareBitmap = [Windows.Graphics.Imaging.SoftwareBitmap]::Convert(
            $softwareBitmap,
            [Windows.Graphics.Imaging.BitmapPixelFormat]::Gray8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Ignore
        )
    }

    $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (-not $ocrEngine) {
        throw "Windows OCR engine is not available for the current user profile languages."
    }

    $ocrResult = Wait-WinRt -Operation ($ocrEngine.RecognizeAsync($softwareBitmap)) -ResultType ([Windows.Media.Ocr.OcrResult])
    $lines = foreach ($line in $ocrResult.Lines) {
        $wordRects = foreach ($word in $line.Words) {
            [pscustomobject]@{
                X      = [double]$word.BoundingRect.X
                Y      = [double]$word.BoundingRect.Y
                Width  = [double]$word.BoundingRect.Width
                Height = [double]$word.BoundingRect.Height
                Right  = [double]($word.BoundingRect.X + $word.BoundingRect.Width)
                Bottom = [double]($word.BoundingRect.Y + $word.BoundingRect.Height)
            }
        }

        if (-not $wordRects) {
            continue
        }

        $minX = ($wordRects | Measure-Object X -Minimum).Minimum
        $minY = ($wordRects | Measure-Object Y -Minimum).Minimum
        $maxRight = ($wordRects | Measure-Object Right -Maximum).Maximum
        $maxBottom = ($wordRects | Measure-Object Bottom -Maximum).Maximum

        [pscustomobject]@{
            Text    = ConvertTo-NormalizedText -Text $line.Text
            X       = [double]$minX
            Y       = [double]$minY
            Width   = [double]($maxRight - $minX)
            Height  = [double]($maxBottom - $minY)
            Right   = [double]$maxRight
            Bottom  = [double]$maxBottom
            CenterX = [double]($minX + (($maxRight - $minX) / 2))
            CenterY = [double]($minY + (($maxBottom - $minY) / 2))
        }
    }

    [pscustomobject]@{
        ImagePath = $resolvedPath
        Text      = ConvertTo-NormalizedText -Text $ocrResult.Text
        Lines     = $lines | Sort-Object Y, X
    }
}

function Get-TimeAnchors {
    param(
        $OcrResult
    )

    $anchors = foreach ($line in $OcrResult.Lines) {
        if ($line.X -gt 80) {
            continue
        }

        $timeText = ConvertTo-TimeText -Token $line.Text
        if (-not $timeText) {
            continue
        }

        [pscustomobject]@{
            Text    = $timeText
            Minutes = ConvertTo-Minutes -TimeText $timeText
            X       = $line.X
            Y       = $line.Y
            CenterY = $line.CenterY
        }
    }

    return $anchors |
        Sort-Object Minutes, CenterY |
        Group-Object Minutes |
        ForEach-Object { $_.Group | Sort-Object X | Select-Object -First 1 } |
        Sort-Object Minutes
}

function Get-TimeScale {
    param(
        [object[]]$TimeAnchors
    )

    if (-not $TimeAnchors -or $TimeAnchors.Count -lt 2) {
        return $null
    }

    $slopes = New-Object System.Collections.Generic.List[double]
    for ($i = 1; $i -lt $TimeAnchors.Count; $i++) {
        $previous = $TimeAnchors[$i - 1]
        $current = $TimeAnchors[$i]
        $deltaY = $current.CenterY - $previous.CenterY
        if ([math]::Abs($deltaY) -lt 0.01) {
            continue
        }

        $slopes.Add(($current.Minutes - $previous.Minutes) / $deltaY)
    }

    if ($slopes.Count -eq 0) {
        return $null
    }

    $minutesPerPixel = ($slopes | Measure-Object -Average).Average
    $reference = $TimeAnchors | Select-Object -First 1
    $maxMinutes = ($TimeAnchors | Measure-Object Minutes -Maximum).Maximum

    [pscustomobject]@{
        MinutesPerPixel = [double]$minutesPerPixel
        ReferenceY      = [double]$reference.CenterY
        ReferenceMinute = [int]$reference.Minutes
        MinMinute       = [int]$reference.Minutes
        MaxMinute       = [int]$maxMinutes
        MeetingTopY     = [double]($reference.Y - 6)
    }
}

function Get-WeekdayIndexFromText {
    param(
        [AllowNull()]
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $value = (ConvertTo-NormalizedText -Text $Text).ToLowerInvariant()

    switch -Regex ($value) {
        "^(mon|monday|[\u6708])" { return 0 }
        "^(tue|tues|tuesday|[\u706B])" { return 1 }
        "^(wed|wednesday|[\u6C34])" { return 2 }
        "^(thu|thurs|thursday|[\u6728])" { return 3 }
        "^(fri|friday|[\u91D1])" { return 4 }
        "^(sat|saturday|[\u571F])" { return 5 }
        "^(sun|sunday|[\u65E5])" { return 6 }
        default { return $null }
    }
}

function ConvertFrom-YToMinute {
    param(
        $Scale,
        [double]$Y
    )

    $minute = $Scale.ReferenceMinute + (($Y - $Scale.ReferenceY) * $Scale.MinutesPerPixel)
    return [int][math]::Round($minute)
}

function Get-DayHeaderCandidates {
    param(
        $OcrResult
    )

    $candidates = foreach ($line in $OcrResult.Lines) {
        if ($line.Y -gt 120 -or $line.X -lt 40) {
            continue
        }

        $dayMatch = [regex]::Match($line.Text, "^(?<day>\d{1,2})(?:\D.*)?$")
        $weekdayIndex = Get-WeekdayIndexFromText -Text $line.Text

        if (-not $dayMatch.Success -and $null -eq $weekdayIndex) {
            continue
        }

        [pscustomobject]@{
            DayOfMonth = if ($dayMatch.Success) { [int]$dayMatch.Groups["day"].Value } else { $null }
            Weekday    = $weekdayIndex
            Text       = $line.Text
            X          = $line.X
            Width      = $line.Width
            CenterX    = $line.CenterX
        }
    }

    $grouped = @{}
    foreach ($candidate in ($candidates | Sort-Object CenterX)) {
        $key = [string][math]::Round($candidate.CenterX / 25)
        if (-not $grouped.ContainsKey($key)) {
            $grouped[$key] = New-Object System.Collections.Generic.List[object]
        }
        $grouped[$key].Add($candidate)
    }

    $reduced = foreach ($key in ($grouped.Keys | Sort-Object {[int]$_})) {
        $grouped[$key] | Sort-Object Width -Descending | Select-Object -First 1
    }

    return $reduced | Sort-Object CenterX
}

function Get-MondayBasedDayIndex {
    param(
        [datetime]$Date
    )

    switch ($Date.DayOfWeek) {
        "Monday" { return 0 }
        "Tuesday" { return 1 }
        "Wednesday" { return 2 }
        "Thursday" { return 3 }
        "Friday" { return 4 }
        "Saturday" { return 5 }
        "Sunday" { return 6 }
        default { return $null }
    }
}

function Resolve-ConsecutiveHeaderDates {
    param(
        [object[]]$Headers,
        [datetime]$ReferenceDate
    )

    if (-not $Headers -or $Headers.Count -eq 0) {
        return @()
    }

    $baseDate = if ($ReferenceDate -ne [datetime]::MinValue) {
        $ReferenceDate.Date
    }
    else {
        (Get-Date).Date
    }

    $searchStart = $baseDate.AddDays(-400)
    $searchEnd = $baseDate.AddDays(400)
    $bestStart = $null
    $bestDistance = [double]::PositiveInfinity

    for ($candidate = $searchStart; $candidate -le $searchEnd; $candidate = $candidate.AddDays(1)) {
        $matches = $true

        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $header = $Headers[$i]
            $headerDate = $candidate.AddDays($i)

            if ($null -ne $header.DayOfMonth -and $headerDate.Day -ne $header.DayOfMonth) {
                $matches = $false
                break
            }

            if ($null -ne $header.Weekday) {
                $weekday = Get-MondayBasedDayIndex -Date $headerDate
                if ($weekday -ne $header.Weekday) {
                    $matches = $false
                    break
                }
            }
        }

        if (-not $matches) {
            continue
        }

        $midpoint = $candidate.AddDays([math]::Floor($Headers.Count / 2))
        $distance = [math]::Abs(($midpoint - $baseDate).TotalDays)
        if ($distance -lt $bestDistance) {
            $bestDistance = $distance
            $bestStart = $candidate
        }
    }

    if (-not $bestStart) {
        return @()
    }

    $dates = New-Object System.Collections.Generic.List[datetime]
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $dates.Add($bestStart.AddDays($i).Date)
    }

    return $dates
}

function Build-DayColumns {
    param(
        $OcrResult,
        [datetime]$ReferenceDate
    )

    $headers = Get-DayHeaderCandidates -OcrResult $OcrResult
    if (-not $headers -or $headers.Count -lt 2) {
        return @()
    }

    $resolvedDates = Resolve-ConsecutiveHeaderDates -Headers $headers -ReferenceDate $ReferenceDate
    if (-not $resolvedDates -or $resolvedDates.Count -lt $headers.Count) {
        return @()
    }

    $columns = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $current = $headers[$i]
        $left = if ($i -eq 0) {
            $current.CenterX - (($headers[$i + 1].CenterX - $current.CenterX) / 2)
        }
        else {
            ($headers[$i - 1].CenterX + $current.CenterX) / 2
        }

        $right = if ($i -eq ($headers.Count - 1)) {
            $current.CenterX + (($current.CenterX - $headers[$i - 1].CenterX) / 2)
        }
        else {
            ($current.CenterX + $headers[$i + 1].CenterX) / 2
        }

        $date = $resolvedDates[$i]

        $columns.Add([pscustomobject]@{
            Index      = $i
            Left       = [double][math]::Max(0, $left)
            Right      = [double]$right
            HeaderText = $current.Text
            Date       = $date.Date
            DateText   = $date.ToString("yyyy-MM-dd")
        })
    }

    return $columns
}

function Get-GridContext {
    param(
        $OcrResult,
        [datetime]$ReferenceDate
    )

    $timeAnchors = Get-TimeAnchors -OcrResult $OcrResult
    $scale = Get-TimeScale -TimeAnchors $timeAnchors
    $columns = Build-DayColumns -OcrResult $OcrResult -ReferenceDate $ReferenceDate

    if (-not $scale -or -not $columns -or $columns.Count -lt 2) {
        return $null
    }

    [pscustomobject]@{
        TimeAnchors = $timeAnchors
        Scale       = $scale
        Columns     = $columns
        LeftAxisMax = (($timeAnchors | Measure-Object X -Maximum).Maximum + 55)
    }
}

function Get-XOverlapRatio {
    param(
        [double]$Left1,
        [double]$Right1,
        [double]$Left2,
        [double]$Right2
    )

    $overlap = [math]::Min($Right1, $Right2) - [math]::Max($Left1, $Left2)
    if ($overlap -le 0) {
        return 0.0
    }

    $width = [math]::Min(($Right1 - $Left1), ($Right2 - $Left2))
    if ($width -le 0) {
        return 0.0
    }

    return $overlap / $width
}

function New-Cluster {
    param(
        $Line
    )

    [pscustomobject]@{
        Lines        = New-Object System.Collections.Generic.List[object]
        Left         = [double]$Line.X
        Right        = [double]$Line.Right
        Top          = [double]$Line.Y
        Bottom       = [double]$Line.Bottom
        AvgHeight    = [double]$Line.Height
        LastBottom   = [double]$Line.Bottom
        LastCenterX  = [double]$Line.CenterX
    }
}

function Add-LineToCluster {
    param(
        $Cluster,
        $Line
    )

    $Cluster.Lines.Add($Line)
    $Cluster.Left = [math]::Min($Cluster.Left, $Line.X)
    $Cluster.Right = [math]::Max($Cluster.Right, $Line.Right)
    $Cluster.Top = [math]::Min($Cluster.Top, $Line.Y)
    $Cluster.Bottom = [math]::Max($Cluster.Bottom, $Line.Bottom)
    $Cluster.AvgHeight = (($Cluster.AvgHeight * ($Cluster.Lines.Count - 1)) + $Line.Height) / $Cluster.Lines.Count
    $Cluster.LastBottom = $Line.Bottom
    $Cluster.LastCenterX = $Line.CenterX
}

function Group-LinesIntoClusters {
    param(
        [object[]]$Lines
    )

    $clusters = New-Object System.Collections.Generic.List[object]
    $current = $null

    foreach ($line in $Lines | Sort-Object Y, X) {
        if (-not $current) {
            $current = New-Cluster -Line $line
            Add-LineToCluster -Cluster $current -Line $line
            continue
        }

        $gap = $line.Y - $current.LastBottom
        $xOverlap = Get-XOverlapRatio -Left1 $current.Left -Right1 $current.Right -Left2 $line.X -Right2 $line.Right
        $sameBand = [math]::Abs($line.CenterX - $current.LastCenterX) -le 35 -or $xOverlap -ge 0.35
        $gapLimit = [math]::Max(12, $current.AvgHeight * 0.9)

        if ($gap -le $gapLimit -and $sameBand) {
            Add-LineToCluster -Cluster $current -Line $line
        }
        else {
            $clusters.Add($current)
            $current = New-Cluster -Line $line
            Add-LineToCluster -Cluster $current -Line $line
        }
    }

    if ($current) {
        $clusters.Add($current)
    }

    return $clusters
}

function Clean-MeetingTitle {
    param(
        [string]$Title
    )

    $value = ConvertTo-NormalizedText -Text $Title
    $value = $value -replace "(?i)\bMicrosoft Teams\b.*$", ""
    $value = $value -replace "(?i)\bTeams Meeting\b.*$", ""
    $value = $value -replace "\s+\($", ""
    return $value.Trim(" ", ";", ",", "-", "/", "|")
}

function Get-MeetingTitleFromCluster {
    param(
        $Cluster
    )

    $ordered = $Cluster.Lines | Sort-Object Y, X
    $kept = New-Object System.Collections.Generic.List[string]

    foreach ($line in $ordered) {
        $text = ConvertTo-NormalizedText -Text $line.Text
        if (-not $text) {
            continue
        }

        if ($text -match "(?i)\bMicrosoft Teams\b" -and $kept.Count -gt 0) {
            break
        }

        if ($kept.Count -ge 2) {
            break
        }

        $kept.Add($text)
    }

    if ($kept.Count -eq 0) {
        return "(title not recognized)"
    }

    $title = Clean-MeetingTitle -Title ($kept -join " ")
    if (-not $title) {
        return "(title not recognized)"
    }

    return $title
}

function New-MeetingRecord {
    param(
        [datetime]$Date,
        [string]$Start,
        [string]$End,
        [string]$Title,
        [string]$SourceImage,
        [string]$SourceText,
        [string]$CategoryKey,
        [string]$CategoryPrefix,
        [string]$CategoryProject
    )

    [pscustomobject]@{
        Date        = $Date.Date
        DateText    = $Date.ToString("yyyy-MM-dd")
        Start       = $Start
        End         = $End
        SortStart   = if ($Start) { $Start } else { "99:99" }
        Title       = if ([string]::IsNullOrWhiteSpace($Title)) { "(title not recognized)" } else { $Title.Trim() }
        SourceImage = $SourceImage
        SourceText  = $SourceText
        CategoryKey = $CategoryKey
        CategoryPrefix = $CategoryPrefix
        CategoryProject = $CategoryProject
    }
}

function Get-DefaultMeetingCategoryRules {
    return @(
        [pscustomobject]@{
            ColorKey    = "yellow"
            Prefix      = "[ProjectA]"
            Project     = "Project A"
            Description = "Yellow meetings"
        },
        [pscustomobject]@{
            ColorKey    = "green"
            Prefix      = "[ProjectB]"
            Project     = "Project B"
            Description = "Green meetings"
        },
        [pscustomobject]@{
            ColorKey    = "red"
            Prefix      = "[Internal]"
            Project     = "Internal"
            Description = "Red meetings"
        },
        [pscustomobject]@{
            ColorKey    = "lightblue"
            Prefix      = "[Other]"
            Project     = "Other / Offhours"
            Description = "Light blue meetings"
        },
        [pscustomobject]@{
            ColorKey    = "unknown"
            Prefix      = ""
            Project     = "Uncategorized"
            Description = "Color could not be determined"
        }
    )
}

function Normalize-MeetingCategoryKey {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return "unknown"
    }

    $normalized = (ConvertTo-NormalizedText -Text $Value).ToLowerInvariant()
    $normalized = $normalized -replace "[\s_-]+", ""

    switch ($normalized) {
        "yellow" { return "yellow" }
        "green" { return "green" }
        "red" { return "red" }
        "lightblue" { return "lightblue" }
        "blue" { return "lightblue" }
        "skyblue" { return "lightblue" }
        "default" { return "lightblue" }
        "unknown" { return "unknown" }
        default { return $normalized }
    }
}

function Get-MeetingCategoryRulesNoteContent {
    param(
        [object[]]$Rules
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("---")
    $lines.Add("type: meeting-category-rules")
    $lines.Add("---")
    $lines.Add("# Meeting Category Rules")
    $lines.Add("")
    $lines.Add("Edit this note to control how Outlook meeting colors map to prefixes.")
    $lines.Add("The script reads this file on every run.")
    $lines.Add("")
    $lines.Add("| colorKey | prefix | project | description |")
    $lines.Add("| --- | --- | --- | --- |")

    foreach ($rule in $Rules) {
        $lines.Add(("| {0} | {1} | {2} | {3} |" -f $rule.ColorKey, $rule.Prefix, $rule.Project, $rule.Description))
    }

    $lines.Add("")
    $lines.Add("Use colorKey values such as yellow, green, red, lightblue, and unknown.")
    return ($lines -join "`r`n") + "`r`n"
}

function Resolve-MeetingCategoryRulesPath {
    param(
        $Config
    )

    $relativePath = if ($Config.PSObject.Properties.Match("meetingCategoryRulesPath").Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Config.meetingCategoryRulesPath)) {
        [string]$Config.meetingCategoryRulesPath
    }
    else {
        "91_MeetingSchedule/00_MeetingCategoryRules.md"
    }

    if ([IO.Path]::IsPathRooted($relativePath)) {
        return [IO.Path]::GetFullPath($relativePath)
    }

    return [IO.Path]::GetFullPath((Join-Path $Config.vaultRoot $relativePath))
}

function Ensure-MeetingCategoryRulesNote {
    param(
        $Config
    )

    $path = Resolve-MeetingCategoryRulesPath -Config $Config
    if (-not (Test-Path -LiteralPath $path)) {
        $content = Get-MeetingCategoryRulesNoteContent -Rules (Get-DefaultMeetingCategoryRules)
        Write-Utf8File -Path $path -Content $content
    }

    return $path
}

function Parse-MeetingCategoryRules {
    param(
        [string]$Content
    )

    $rules = New-Object System.Collections.Generic.List[object]
    foreach ($line in ($Content -split "\r?\n")) {
        if ($line -notmatch "^\|") {
            continue
        }

        if ($line -match "^\|\s*---" -or $line -match "^\|\s*colorKey\s*\|") {
            continue
        }

        $match = [regex]::Match($line, "^\|\s*(?<color>[^|]+?)\s*\|\s*(?<prefix>[^|]*?)\s*\|\s*(?<project>[^|]*?)\s*\|\s*(?<description>[^|]*?)\s*\|$")
        if (-not $match.Success) {
            continue
        }

        $colorKey = Normalize-MeetingCategoryKey -Value $match.Groups["color"].Value
        if ([string]::IsNullOrWhiteSpace($colorKey)) {
            continue
        }

        $rules.Add([pscustomobject]@{
            ColorKey    = $colorKey
            Prefix      = $match.Groups["prefix"].Value.Trim()
            Project     = $match.Groups["project"].Value.Trim()
            Description = $match.Groups["description"].Value.Trim()
        })
    }

    return $rules
}

function Get-MeetingCategoryRuleSet {
    param(
        $Config
    )

    $path = Ensure-MeetingCategoryRulesNote -Config $Config
    $content = Read-Utf8File -Path $path
    $rules = Parse-MeetingCategoryRules -Content $content

    if (-not $rules -or $rules.Count -eq 0) {
        $rules = Get-DefaultMeetingCategoryRules
    }

    $ruleMap = @{}
    foreach ($rule in $rules) {
        $ruleMap[(Normalize-MeetingCategoryKey -Value $rule.ColorKey)] = $rule
    }

    if (-not $ruleMap.ContainsKey("unknown")) {
        $unknownRule = [pscustomobject]@{
            ColorKey    = "unknown"
            Prefix      = ""
            Project     = "Uncategorized"
            Description = "Color could not be determined"
        }
        $rules += $unknownRule
        $ruleMap["unknown"] = $unknownRule
    }

    return [pscustomobject]@{
        Path            = $path
        Rules           = @($rules)
        RuleMap         = $ruleMap
        AllowedColorKeys = @($rules | ForEach-Object { Normalize-MeetingCategoryKey -Value $_.ColorKey } | Select-Object -Unique)
    }
}

function Get-MeetingCategoryRulesPromptText {
    param(
        [object[]]$Rules
    )

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Color classification rules:")
    foreach ($rule in $Rules) {
        $lines.Add(("- colorKey={0}; prefix={1}; project={2}; description={3}" -f $rule.ColorKey, $rule.Prefix, $rule.Project, $rule.Description))
    }

    return $lines -join "`n"
}

function Apply-MeetingCategoryPrefix {
    param(
        [string]$Title,
        $Rule
    )

    $cleanTitle = Clean-MeetingTitle -Title $Title
    if (-not $Rule) {
        return $cleanTitle
    }

    $prefix = ConvertTo-NormalizedText -Text ([string]$Rule.Prefix)
    if ([string]::IsNullOrWhiteSpace($prefix)) {
        return $cleanTitle
    }

    $body = $cleanTitle -replace "^\[[^\]]+\]\s*", ""
    if ($body.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $body
    }

    if ($prefix.EndsWith("]")) {
        return $prefix + $body
    }

    return ($prefix + " " + $body).Trim()
}

function Get-ImageMimeType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )

    switch ([IO.Path]::GetExtension($ImagePath).ToLowerInvariant()) {
        ".png" { return "image/png" }
        ".jpg" { return "image/jpeg" }
        ".jpeg" { return "image/jpeg" }
        ".webp" { return "image/webp" }
        ".gif" { return "image/gif" }
        default {
            throw "Unsupported image type for OpenAI vision input: $ImagePath"
        }
    }
}

function ConvertTo-ImageDataUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )

    $mimeType = Get-ImageMimeType -ImagePath $ImagePath
    $bytes = [IO.File]::ReadAllBytes($ImagePath)
    $base64 = [Convert]::ToBase64String($bytes)
    return "data:{0};base64,{1}" -f $mimeType, $base64
}

function Get-ExceptionResponseBody {
    param(
        $Exception
    )

    if (-not $Exception -or -not $Exception.Response) {
        return $null
    }

    $stream = $null
    $reader = $null

    try {
        $stream = $Exception.Response.GetResponseStream()
        if (-not $stream) {
            return $null
        }

        $reader = New-Object System.IO.StreamReader($stream)
        return $reader.ReadToEnd()
    }
    catch {
        return $null
    }
    finally {
        if ($reader) {
            $reader.Dispose()
        }
        elseif ($stream) {
            $stream.Dispose()
        }
    }
}

function Get-OpenAiApiKey {
    param(
        $Config
    )

    $environmentVariableName = "OPENAI_API_KEY"
    if ($Config -and $Config.PSObject.Properties.Match("openAiApiKeyEnvVar").Count -gt 0) {
        $candidateName = [string]$Config.openAiApiKeyEnvVar
        if (-not [string]::IsNullOrWhiteSpace($candidateName)) {
            $environmentVariableName = $candidateName.Trim()
        }
    }

    $targets = @(
        [System.EnvironmentVariableTarget]::Process,
        [System.EnvironmentVariableTarget]::User,
        [System.EnvironmentVariableTarget]::Machine
    )

    foreach ($target in $targets) {
        $value = [Environment]::GetEnvironmentVariable($environmentVariableName, $target)
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }

    throw ("OpenAI API key was not found in the {0} environment variable." -f $environmentVariableName)
}

function Get-OpenAiOutputText {
    param(
        $Response
    )

    if ($Response -and $Response.PSObject.Properties.Match("output_text").Count -gt 0) {
        $outputText = [string]$Response.output_text
        if (-not [string]::IsNullOrWhiteSpace($outputText)) {
            return $outputText
        }
    }

    $textParts = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($Response.output)) {
        if ($null -eq $item -or $item.type -ne "message") {
            continue
        }

        foreach ($content in @($item.content)) {
            if ($null -ne $content -and $content.type -eq "output_text" -and -not [string]::IsNullOrWhiteSpace([string]$content.text)) {
                $textParts.Add([string]$content.text)
            }
        }
    }

    if ($textParts.Count -eq 0) {
        return $null
    }

    return ($textParts -join "`n").Trim()
}

function Get-OpenAiRefusalText {
    param(
        $Response
    )

    $refusalParts = New-Object System.Collections.Generic.List[string]
    foreach ($item in @($Response.output)) {
        if ($null -eq $item -or $item.type -ne "message") {
            continue
        }

        foreach ($content in @($item.content)) {
            if ($null -eq $content -or $content.type -ne "refusal") {
                continue
            }

            $text = if ($content.PSObject.Properties.Match("refusal").Count -gt 0) {
                [string]$content.refusal
            }
            else {
                [string]$content.text
            }

            if (-not [string]::IsNullOrWhiteSpace($text)) {
                $refusalParts.Add($text)
            }
        }
    }

    if ($refusalParts.Count -eq 0) {
        return $null
    }

    return ($refusalParts -join " ").Trim()
}

function Get-OpenAiMeetingExtractionPrompt {
    param(
        [datetime]$ReferenceDate,
        [object[]]$CategoryRules,
        [int]$StepMinutes
    )

    return @"
Reference date: $($ReferenceDate.ToString("yyyy-MM-dd")) (Asia/Tokyo).
The image is an Outlook weekly calendar screenshot.
Extract timed meetings that are clearly visible from every visible day column in the displayed week.
Include the entire displayed week, not only meetings up to the reference date. Include future days later in the same visible week as well.
Infer each meeting's absolute date by matching the visible weekday and day-of-month headers to the calendar week closest to the reference date.
Ignore all-day banners, holidays, paid leave, travel, private or personal appointments, and availability shading.
Keep only the real meeting title. Remove boilerplate such as Microsoft Teams labels, organizer names, attendee names, room names, and recurrence metadata when they are not part of the title.
If a title wraps across multiple lines, merge only the title lines.
Classify each meeting by the fill color of its block, not by the title text.
Use one `colorKey` from the rules below.
$([string](Get-MeetingCategoryRulesPromptText -Rules $CategoryRules))
The calendar uses only $StepMinutes-minute increments. All start and end times must be on that grid.
For a 30-minute grid, valid times are only HH:00 and HH:30. Never return HH:15 or HH:45.
If the exact time is slightly unclear, choose the nearest valid $StepMinutes-minute boundary from the visible box.
Estimate start and end from the visible box position and height when needed, but do not invent meetings that are not clearly present.
Do not omit a visible meeting only because its date is after the reference date.
If any field is too uncertain, omit that meeting instead of guessing.
Return only the JSON schema output.
"@
}

function ConvertTo-MeetingDate {
    param(
        [string]$Value,
        [datetime]$ReferenceDate
    )

    $normalized = ConvertTo-NormalizedText -Text $Value
    if (-not $normalized) {
        throw "Meeting date is empty."
    }

    $fullDateFormats = @(
        "yyyy-MM-dd",
        "yyyy-M-d",
        "yyyy/MM/dd",
        "yyyy/M/d"
    )

    foreach ($format in $fullDateFormats) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($normalized, $format, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed.Date
        }
    }

    $monthDayMatch = [regex]::Match($normalized, "^(?<month>\d{1,2})[/-](?<day>\d{1,2})$")
    if ($monthDayMatch.Success) {
        return (Get-Date -Year $ReferenceDate.Year -Month ([int]$monthDayMatch.Groups["month"].Value) -Day ([int]$monthDayMatch.Groups["day"].Value)).Date
    }

    $fallback = [datetime]::MinValue
    if ([datetime]::TryParse($normalized, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeLocal, [ref]$fallback)) {
        return $fallback.Date
    }

    throw "Unsupported meeting date format: $Value"
}

function ConvertTo-MeetingRecordsFromOpenAiPayload {
    param(
        $Payload,
        [datetime]$ReferenceDate,
        [string]$SourceImage,
        $CategoryRuleMap,
        [int]$StepMinutes
    )

    $records = New-Object System.Collections.Generic.List[object]
    $items = @()

    if ($Payload -and $Payload.PSObject.Properties.Match("meetings").Count -gt 0 -and $Payload.meetings) {
        $items = @($Payload.meetings)
    }

    foreach ($item in $items) {
        if ($null -eq $item) {
            continue
        }

        try {
            $date = ConvertTo-MeetingDate -Value ([string]$item.date) -ReferenceDate $ReferenceDate
            $start = ConvertTo-TimeText -Token ([string]$item.start)
            $end = ConvertTo-TimeText -Token ([string]$item.end)
            $categoryKey = Normalize-MeetingCategoryKey -Value ([string]$item.colorKey)
            $rule = if ($CategoryRuleMap.ContainsKey($categoryKey)) { $CategoryRuleMap[$categoryKey] } else { $null }
            $title = Apply-MeetingCategoryPrefix -Title ([string]$item.title) -Rule $rule

            if (-not $start -or -not $end -or -not $title) {
                continue
            }

            $snapStep = if ($StepMinutes -gt 0) { $StepMinutes } else { 30 }
            $startMinute = Snap-Minutes -Minutes (ConvertTo-Minutes -TimeText $start) -Step $snapStep -Mode Floor
            $endMinute = Snap-Minutes -Minutes (ConvertTo-Minutes -TimeText $end) -Step $snapStep -Mode Ceiling
            if ($endMinute -le $startMinute) {
                $endMinute = $startMinute + $snapStep
            }

            $start = ConvertTo-TimeStringFromMinutes -Minutes $startMinute
            $end = ConvertTo-TimeStringFromMinutes -Minutes $endMinute

            if ((ConvertTo-Minutes -TimeText $end) -le (ConvertTo-Minutes -TimeText $start)) {
                continue
            }

            $records.Add((New-MeetingRecord `
                -Date $date `
                -Start $start `
                -End $end `
                -Title $title `
                -SourceImage $SourceImage `
                -SourceText ($item | ConvertTo-Json -Compress -Depth 6) `
                -CategoryKey $categoryKey `
                -CategoryPrefix $(if ($rule) { [string]$rule.Prefix } else { "" }) `
                -CategoryProject $(if ($rule) { [string]$rule.Project } else { "" })))
        }
        catch {
            continue
        }
    }

    return $records
}

function Invoke-OpenAiMeetingExtraction {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath,

        [datetime]$ReferenceDate,

        $Config,
        $CategoryRuleSet,
        [int]$StepMinutes
    )

    $apiKey = Get-OpenAiApiKey -Config $Config
    $apiUrl = if ($Config.PSObject.Properties.Match("openAiApiUrl").Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Config.openAiApiUrl)) {
        [string]$Config.openAiApiUrl
    }
    else {
        "https://api.openai.com/v1/responses"
    }

    $model = if ($Config.PSObject.Properties.Match("openAiModel").Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Config.openAiModel)) {
        [string]$Config.openAiModel
    }
    else {
        "gpt-5.4-mini"
    }

    $imageDetail = if ($Config.PSObject.Properties.Match("openAiImageDetail").Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$Config.openAiImageDetail)) {
        [string]$Config.openAiImageDetail
    }
    else {
        "original"
    }

    $timeoutSeconds = if ($Config.PSObject.Properties.Match("openAiTimeoutSeconds").Count -gt 0 -and [int]$Config.openAiTimeoutSeconds -gt 0) {
        [int]$Config.openAiTimeoutSeconds
    }
    else {
        120
    }

    $maxOutputTokens = if ($Config.PSObject.Properties.Match("openAiMaxOutputTokens").Count -gt 0 -and [int]$Config.openAiMaxOutputTokens -gt 0) {
        [int]$Config.openAiMaxOutputTokens
    }
    else {
        4000
    }

    $requestBody = @{
        model = $model
        input = @(
            @{
                role = "system"
                content = "You extract structured meeting schedules from Outlook weekly calendar screenshots."
            },
            @{
                role = "user"
                content = @(
                    @{
                        type = "input_text"
                        text = Get-OpenAiMeetingExtractionPrompt -ReferenceDate $ReferenceDate -CategoryRules $CategoryRuleSet.Rules -StepMinutes $StepMinutes
                    },
                    @{
                        type = "input_image"
                        image_url = ConvertTo-ImageDataUrl -ImagePath $ImagePath
                        detail = $imageDetail
                    }
                )
            }
        )
        text = @{
            format = @{
                type = "json_schema"
                name = "outlook_meeting_schedule"
                schema = @{
                    type = "object"
                    properties = @{
                        meetings = @{
                            type = "array"
                            items = @{
                                type = "object"
                                properties = @{
                                    date = @{
                                        type = "string"
                                        description = "Absolute date in YYYY-MM-DD."
                                    }
                                    start = @{
                                        type = "string"
                                        description = "Start time in HH:MM 24-hour format."
                                    }
                                    end = @{
                                        type = "string"
                                        description = "End time in HH:MM 24-hour format."
                                    }
                                    title = @{
                                        type = "string"
                                        description = "Meeting title only."
                                    }
                                    colorKey = @{
                                        type = "string"
                                        description = "Meeting block fill color category."
                                        enum = @($CategoryRuleSet.AllowedColorKeys)
                                    }
                                }
                                required = @("date", "start", "end", "title", "colorKey")
                                additionalProperties = $false
                            }
                        }
                    }
                    required = @("meetings")
                    additionalProperties = $false
                }
                strict = $true
            }
        }
        max_output_tokens = $maxOutputTokens
    }

    $headers = @{
        Authorization = "Bearer $apiKey"
    }

    try {
        $response = Invoke-RestMethod `
            -Method Post `
            -Uri $apiUrl `
            -Headers $headers `
            -ContentType "application/json" `
            -Body ($requestBody | ConvertTo-Json -Depth 20 -Compress) `
            -TimeoutSec $timeoutSeconds
    }
    catch {
        $responseBody = Get-ExceptionResponseBody -Exception $_.Exception
        if ($responseBody) {
            throw "OpenAI API request failed: $responseBody"
        }

        throw
    }

    $outputText = Get-OpenAiOutputText -Response $response
    if ([string]::IsNullOrWhiteSpace($outputText)) {
        $refusalText = Get-OpenAiRefusalText -Response $response
        if ($refusalText) {
            throw "OpenAI refused the request: $refusalText"
        }

        throw "OpenAI returned no structured text output."
    }

    try {
        $payload = $outputText | ConvertFrom-Json
    }
    catch {
        throw "OpenAI returned invalid JSON: $outputText"
    }

    return [pscustomobject]@{
        Mode       = "OpenAI"
        Meetings   = ConvertTo-MeetingRecordsFromOpenAiPayload -Payload $payload -ReferenceDate $ReferenceDate -SourceImage $ImagePath -CategoryRuleMap $CategoryRuleSet.RuleMap -StepMinutes $StepMinutes
        Context    = [pscustomobject]@{
            provider      = "openai"
            model         = $model
            imageDetail   = $imageDetail
            referenceDate = $ReferenceDate.ToString("yyyy-MM-dd")
            categoryRulesPath = $CategoryRuleSet.Path
            timeSnapMinutes = $StepMinutes
        }
        RawResponse = $response
        RawText     = $outputText
    }
}

function Parse-GridMeetings {
    param(
        $OcrResult,
        $GridContext,
        [int]$StepMinutes
    )

    $meetings = New-Object System.Collections.Generic.List[object]
    $meetingLines = foreach ($line in $OcrResult.Lines) {
        if ($line.CenterY -lt $GridContext.Scale.MeetingTopY) {
            continue
        }

        if ($line.CenterX -lt $GridContext.LeftAxisMax) {
            continue
        }

        if (ConvertTo-TimeText -Token $line.Text) {
            continue
        }

        $column = $GridContext.Columns | Where-Object {
            $line.CenterX -ge $_.Left -and $line.CenterX -le $_.Right
        } | Select-Object -First 1

        if (-not $column) {
            continue
        }

        [pscustomobject]@{
            Column    = $column
            Line      = $line
            ColumnKey = $column.DateText
        }
    }

    foreach ($group in $meetingLines | Group-Object ColumnKey) {
        $column = $group.Group[0].Column
        $lines = $group.Group | ForEach-Object { $_.Line }
        $clusters = Group-LinesIntoClusters -Lines $lines

        foreach ($cluster in $clusters) {
            $rawTitle = Get-MeetingTitleFromCluster -Cluster $cluster
            if (-not $rawTitle -or $rawTitle -eq "(title not recognized)") {
                continue
            }

            $startMinute = ConvertFrom-YToMinute -Scale $GridContext.Scale -Y ($cluster.Top - 4)
            $endMinute = ConvertFrom-YToMinute -Scale $GridContext.Scale -Y ($cluster.Bottom + 8)

            $startMinute = Snap-Minutes -Minutes $startMinute -Step $StepMinutes -Mode Floor
            $endMinute = Snap-Minutes -Minutes $endMinute -Step $StepMinutes -Mode Ceiling

            $startMinute = [math]::Max($GridContext.Scale.MinMinute, $startMinute)
            $endMinute = [math]::Max($startMinute + $StepMinutes, $endMinute)

            $meeting = New-MeetingRecord `
                -Date $column.Date `
                -Start (ConvertTo-TimeStringFromMinutes -Minutes $startMinute) `
                -End (ConvertTo-TimeStringFromMinutes -Minutes $endMinute) `
                -Title $rawTitle `
                -SourceImage $OcrResult.ImagePath `
                -SourceText (($cluster.Lines | Sort-Object Y, X | ForEach-Object { $_.Text }) -join " | ")

            $meetings.Add($meeting)
        }
    }

    return $meetings
}

function Parse-LineMeetings {
    param(
        $OcrResult,
        [datetime]$ReferenceDate,
        [int]$StepMinutes
    )

    $meetings = New-Object System.Collections.Generic.List[object]
    $currentDate = $null
    $pending = $null

    foreach ($line in $OcrResult.Lines | Sort-Object Y, X) {
        $text = ConvertTo-NormalizedText -Text $line.Text
        if (-not $text) {
            continue
        }

        $dateMatch = Try-ParseDateFromText -Text $text -ReferenceDate $ReferenceDate
        $workingText = $text
        if ($dateMatch) {
            $currentDate = $dateMatch.Date
            $workingText = ($workingText -replace [regex]::Escape($dateMatch.MatchedText), "").Trim(" ", "-", ":", "|")
        }

        $timeMatch = [regex]::Match($workingText, "(?<start>\d{1,2}:\d{2})\s*(?:-|~)\s*(?<end>\d{1,2}:\d{2})(?<rest>.*)$")
        if ($timeMatch.Success -and $currentDate) {
            $title = Clean-MeetingTitle -Title $timeMatch.Groups["rest"].Value
            $meetings.Add((New-MeetingRecord `
                -Date $currentDate `
                -Start (ConvertTo-TimeText -Token $timeMatch.Groups["start"].Value) `
                -End (ConvertTo-TimeText -Token $timeMatch.Groups["end"].Value) `
                -Title $title `
                -SourceImage $OcrResult.ImagePath `
                -SourceText $text))
            $pending = $null
            continue
        }

        if ($pending) {
            $pending.Title = Clean-MeetingTitle -Title ($pending.Title + " " + $workingText)
            continue
        }

        $singleTimeMatch = [regex]::Match($workingText, "^(?<start>\d{1,2}:\d{2})(?<rest>.*)$")
        if ($singleTimeMatch.Success -and $currentDate) {
            $pending = [pscustomobject]@{
                Date  = $currentDate
                Start = ConvertTo-TimeText -Token $singleTimeMatch.Groups["start"].Value
                Title = Clean-MeetingTitle -Title $singleTimeMatch.Groups["rest"].Value
                Raw   = $text
            }
            continue
        }
    }

    if ($pending) {
        $startMinute = ConvertTo-Minutes -TimeText $pending.Start
        $endMinute = $startMinute + $StepMinutes
        $meetings.Add((New-MeetingRecord `
            -Date $pending.Date `
            -Start $pending.Start `
            -End (ConvertTo-TimeStringFromMinutes -Minutes $endMinute) `
            -Title $pending.Title `
            -SourceImage $OcrResult.ImagePath `
            -SourceText $pending.Raw))
    }

    return $meetings
}

function Get-WeekdayLabel {
    param(
        [datetime]$Date
    )

    switch ($Date.DayOfWeek) {
        "Monday" { return "Mon" }
        "Tuesday" { return "Tue" }
        "Wednesday" { return "Wed" }
        "Thursday" { return "Thu" }
        "Friday" { return "Fri" }
        "Saturday" { return "Sat" }
        "Sunday" { return "Sun" }
        default { return $Date.DayOfWeek.ToString() }
    }
}

function Get-IsoWeekInfo {
    param(
        [datetime]$Date
    )

    $calendar = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
    $week = $calendar.GetWeekOfYear($Date, [System.Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
    $year = $Date.Year

    if ($Date.Month -eq 1 -and $week -ge 52) {
        $year -= 1
    }
    elseif ($Date.Month -eq 12 -and $week -eq 1) {
        $year += 1
    }

    [pscustomobject]@{
        Year = $year
        Week = $week
    }
}

function Resolve-TemplatePath {
    param(
        [string]$VaultRoot,
        [string]$Pattern,
        [datetime]$Date
    )

    if ([IO.Path]::IsPathRooted($Pattern)) {
        return [IO.Path]::GetFullPath($Pattern)
    }

    $weekInfo = Get-IsoWeekInfo -Date $Date
    $map = [ordered]@{
        "{date}"     = $Date.ToString("yyyy-MM-dd")
        "{yyyy}"     = $Date.ToString("yyyy")
        "{MM}"       = $Date.ToString("MM")
        "{dd}"       = $Date.ToString("dd")
        "{week}"     = ("{0:D2}" -f $weekInfo.Week)
        "{weekYear}" = $weekInfo.Year
    }

    $relativePath = $Pattern
    foreach ($key in $map.Keys) {
        $relativePath = $relativePath.Replace($key, [string]$map[$key])
    }

    return [IO.Path]::GetFullPath((Join-Path $VaultRoot $relativePath))
}

function Get-DefaultConfigValues {
    param(
        [string]$VaultRoot
    )

    return [ordered]@{
        vaultRoot                = $VaultRoot
        weeklyNotePattern        = "91_MeetingSchedule/{weekYear}-W{week}.md"
        dailyNotePattern         = "01_Daily/{date}.md"
        dailyTemplatePath        = "90_Templates/Daily template.md"
        dailyInsertBeforeHeading = "## " + [string]([char]0x660E) + [char]0x65E5 + [char]0x3084 + [char]0x308B + [char]0x3053 + [char]0x3068
        dailyMeetingHeading      = "## " + [string]([char]0x4F1A) + [char]0x8B70
        dailyStartMarker         = "<!-- outlook-screenshot:start -->"
        dailyEndMarker           = "<!-- outlook-screenshot:end -->"
        noMeetingsLine           = "- " + [string]([char]0x4E88) + [char]0x5B9A + [char]0x306A + [char]0x3057
        meetingCategoryRulesPath = "91_MeetingSchedule/00_MeetingCategoryRules.md"
        extractionMode           = "openai"
        openAiApiUrl             = "https://api.openai.com/v1/responses"
        openAiApiKeyEnvVar       = "OPENAI_API_KEY"
        openAiModel              = "gpt-5.4-mini"
        openAiImageDetail        = "original"
        openAiTimeoutSeconds     = 120
        openAiMaxOutputTokens    = 4000
    }
}

function Resolve-DefaultVaultRoot {
    $value = Get-EnvironmentVariableValue -Names @("OUTLOOK_SCREENSHOT_VAULT_ROOT")
    if ([string]::IsNullOrWhiteSpace($value)) {
        return ""
    }

    return [IO.Path]::GetFullPath($value)
}

function Resolve-ConfigPath {
    param(
        [string]$ConfiguredPath
    )

    if (-not [string]::IsNullOrWhiteSpace($ConfiguredPath)) {
        return [IO.Path]::GetFullPath($ConfiguredPath)
    }

    $configuredByEnvironment = Get-EnvironmentVariableValue -Names @("OUTLOOK_SCREENSHOT_CONFIG_PATH")
    if (-not [string]::IsNullOrWhiteSpace($configuredByEnvironment)) {
        return [IO.Path]::GetFullPath($configuredByEnvironment)
    }

    return Join-Path $PSScriptRoot "outlook-screenshot.json"
}

function Ensure-ParentDirectory {
    param(
        [string]$Path
    )

    $parent = [IO.Path]::GetDirectoryName($Path)
    if ($parent -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}

function Save-ClipboardImageToFile {
    param(
        [string]$Path,
        [string]$BaseDirectory
    )

    if ($Path) {
        $targetPath = [IO.Path]::GetFullPath($Path)
    }
    else {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $directory = if ($BaseDirectory) {
            $BaseDirectory
        }
        else {
            Join-Path $PSScriptRoot "output\clipboard"
        }

        $targetPath = Join-Path $directory ("clipboard-{0}.png" -f $timestamp)
    }

    Ensure-ParentDirectory -Path $targetPath

    $clipboardScript = {
        param($TargetPath)

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        if (-not [System.Windows.Forms.Clipboard]::ContainsImage()) {
            throw "Clipboard does not contain an image."
        }

        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if (-not $image) {
            throw "Failed to read image from clipboard."
        }

        try {
            $image.Save($TargetPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $image.Dispose()
        }

        return $TargetPath
    }

    $savedPath = Invoke-InStaRunspace -ScriptBlock $clipboardScript -ArgumentList @($targetPath)
    return [string]($savedPath | Select-Object -Last 1)
}

function Read-Utf8File {
    param(
        [string]$Path
    )

    return [IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Content
    )

    Ensure-ParentDirectory -Path $Path
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, $Content, $utf8)
}

function Build-MarkedBlock {
    param(
        [string]$StartMarker,
        [string]$EndMarker,
        [string]$Content
    )

    return (@(
        $StartMarker
        $Content.TrimEnd()
        $EndMarker
    ) -join "`r`n").TrimEnd()
}

function Set-BlockBeforeHeading {
    param(
        [string]$Path,
        [string]$BeforeHeadingText,
        [string]$StartMarker,
        [string]$EndMarker,
        [string]$BodyContent,
        [string]$FallbackTitle
    )

    $blockText = Build-MarkedBlock -StartMarker $StartMarker -EndMarker $EndMarker -Content $BodyContent

    if (Test-Path -LiteralPath $Path) {
        $existing = Read-Utf8File -Path $Path
        if ($existing.Contains($StartMarker) -and $existing.Contains($EndMarker)) {
            $pattern = [regex]::Escape($StartMarker) + ".*?" + [regex]::Escape($EndMarker)
            $updated = [regex]::Replace(
                $existing,
                $pattern,
                [System.Text.RegularExpressions.MatchEvaluator]{ param($match) $blockText },
                [System.Text.RegularExpressions.RegexOptions]::Singleline
            )
        }
        else {
            $beforePattern = "(?m)^" + [regex]::Escape($BeforeHeadingText) + "\s*$"
            $match = [regex]::Match($existing, $beforePattern)
            if ($match.Success) {
                $insert = $blockText + "`r`n`r`n"
                $updated = $existing.Insert($match.Index, $insert)
            }
            else {
                $updated = $existing.TrimEnd() + "`r`n`r`n" + $blockText + "`r`n"
            }
        }
    }
    else {
        $updated = "# $FallbackTitle`r`n`r`n$blockText`r`n"
    }

    Write-Utf8File -Path $Path -Content $updated
}

function Write-WeeklyMeetingNote {
    param(
        [string]$Path,
        [datetime]$RepresentativeDate,
        [object[]]$Meetings
    )

    $weekInfo = Get-IsoWeekInfo -Date $RepresentativeDate
    $content = @(
        "---"
        "type: meeting-schedule"
        "week: {0}-W{1:D2}" -f $weekInfo.Year, $weekInfo.Week
        "---"
        "# {0}-W{1:D2} Meeting Schedule" -f $weekInfo.Year, $weekInfo.Week
        ""
        (Format-WeeklyMarkdown -Meetings $Meetings)
        ""
    ) -join "`r`n"

    Write-Utf8File -Path $Path -Content $content
}

function Get-DailyMeetingSectionBody {
    param(
        [string]$Heading,
        [string]$MeetingMarkdown
    )

    return (@(
        $Heading
        $MeetingMarkdown.TrimEnd()
    ) -join "`r`n").TrimEnd()
}

function Get-DailyTemplateMeetingSectionBody {
    param(
        [string]$Heading,
        [string]$WeeklyFolder,
        [string]$NoMeetingsLine
    )

    $lines = @(
        $Heading,
        "<%*",
        "const noteDate = tp.file.title.trim();",
        'const weekFile = `${tp.date.now("GGGG-[W]WW", 0, noteDate, "YYYY-MM-DD")}.md`;',
        ('const schedulePath = "{0}/${{weekFile}}";' -f $WeeklyFolder),
        ('const fallback = "{0}";' -f $NoMeetingsLine),
        'const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");',
        "const file = app.vault.getAbstractFileByPath(schedulePath);",
        "if (!file) {",
        "  tR += fallback;",
        "} else {",
        "  const content = await app.vault.cachedRead(file);",
        '  const sectionPattern = new RegExp(`## ${escapeRegExp(noteDate)} \\([^\\n]+\\)\\r?\\n([\\s\\S]*?)(?=\\r?\\n## |$)`);',
        "  const match = content.match(sectionPattern);",
        "  tR += match && match[1].trim() ? match[1].trim() : fallback;",
        "}",
        "%>"
    )

    return ($lines -join "`r`n").TrimEnd()
}

function Load-Config {
    param(
        [string]$Path
    )

    $defaultVaultRoot = Resolve-DefaultVaultRoot
    $defaults = Get-DefaultConfigValues -VaultRoot $defaultVaultRoot
    $loaded = $null

    if (Test-Path -LiteralPath $Path) {
        $loaded = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    else {
        $loaded = [pscustomobject]@{}
    }

    foreach ($key in $defaults.Keys) {
        if ($loaded.PSObject.Properties.Match($key).Count -eq 0) {
            $loaded | Add-Member -NotePropertyName $key -NotePropertyValue $defaults[$key]
        }
    }

    if ($loaded.PSObject.Properties.Match("configPath").Count -eq 0) {
        $loaded | Add-Member -NotePropertyName configPath -NotePropertyValue $Path
    }
    else {
        $loaded.configPath = $Path
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$loaded.vaultRoot) -and -not [IO.Path]::IsPathRooted([string]$loaded.vaultRoot)) {
        $configDirectory = Split-Path -Path $Path -Parent
        $loaded.vaultRoot = [IO.Path]::GetFullPath((Join-Path $configDirectory ([string]$loaded.vaultRoot)))
    }

    return $loaded
}

function Assert-ObsidianConfig {
    param(
        $Config
    )

    if ([string]::IsNullOrWhiteSpace([string]$Config.vaultRoot)) {
        throw "vaultRoot is not configured. Set it in outlook-screenshot.json or OUTLOOK_SCREENSHOT_VAULT_ROOT before writing to Obsidian."
    }
}

function Save-ConfigIfMissing {
    param(
        $Config
    )

    if (Test-Path -LiteralPath $Config.configPath) {
        return
    }

    $persisted = [ordered]@{
        vaultRoot                = $Config.vaultRoot
        weeklyNotePattern        = $Config.weeklyNotePattern
        dailyNotePattern         = $Config.dailyNotePattern
        dailyTemplatePath        = $Config.dailyTemplatePath
        dailyInsertBeforeHeading = $Config.dailyInsertBeforeHeading
        dailyMeetingHeading      = $Config.dailyMeetingHeading
        dailyStartMarker         = $Config.dailyStartMarker
        dailyEndMarker           = $Config.dailyEndMarker
        noMeetingsLine           = $Config.noMeetingsLine
        meetingCategoryRulesPath = $Config.meetingCategoryRulesPath
        extractionMode           = $Config.extractionMode
        openAiApiUrl             = $Config.openAiApiUrl
        openAiApiKeyEnvVar       = $Config.openAiApiKeyEnvVar
        openAiModel              = $Config.openAiModel
        openAiImageDetail        = $Config.openAiImageDetail
        openAiTimeoutSeconds     = $Config.openAiTimeoutSeconds
        openAiMaxOutputTokens    = $Config.openAiMaxOutputTokens
    }

    Write-JsonFile -Path $Config.configPath -Data $persisted
}

function Resolve-ExtractionMode {
    param(
        $Config,
        [string]$RequestedMode
    )

    if ($RequestedMode -eq "OpenAI") {
        return "OpenAI"
    }

    if ($RequestedMode -in @("Grid", "Line")) {
        return "OCR"
    }

    $configuredMode = ""
    if ($Config -and $Config.PSObject.Properties.Match("extractionMode").Count -gt 0) {
        $configuredMode = [string]$Config.extractionMode
    }

    switch -Regex ($configuredMode.Trim()) {
        "^(?i:ocr)$" { return "OCR" }
        default { return "OpenAI" }
    }
}

function Format-MeetingLine {
    param(
        $Meeting
    )

    if ($Meeting.End) {
        return "- {0}-{1} {2}" -f $Meeting.Start, $Meeting.End, $Meeting.Title
    }

    return "- {0} {1}" -f $Meeting.Start, $Meeting.Title
}

function Format-WeeklyMarkdown {
    param(
        [object[]]$Meetings
    )

    if (-not $Meetings -or $Meetings.Count -eq 0) {
        return "- " + [string]([char]0x4E88) + [char]0x5B9A + [char]0x306A + [char]0x3057
    }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($group in $Meetings | Group-Object DateText | Sort-Object Name) {
        $date = [datetime]$group.Group[0].Date
        $lines.Add(("## {0} ({1})" -f $date.ToString("yyyy-MM-dd"), (Get-WeekdayLabel -Date $date)))
        foreach ($meeting in $group.Group | Sort-Object SortStart, Title) {
            $lines.Add((Format-MeetingLine -Meeting $meeting))
        }
        $lines.Add("")
    }

    return ($lines -join "`r`n").TrimEnd()
}

function Format-DailyMarkdown {
    param(
        [object[]]$Meetings
    )

    if (-not $Meetings -or $Meetings.Count -eq 0) {
        return "- " + [string]([char]0x4E88) + [char]0x5B9A + [char]0x306A + [char]0x3057
    }

    return (($Meetings | Sort-Object SortStart, Title | ForEach-Object {
        Format-MeetingLine -Meeting $_
    }) -join "`r`n")
}

function Write-JsonFile {
    param(
        [string]$Path,
        $Data
    )

    Ensure-ParentDirectory -Path $Path
    $Data | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Parse-MeetingsFromOcr {
    param(
        $OcrResult,
        [datetime]$ReferenceDate,
        [string]$Mode,
        [int]$StepMinutes
    )

    $selectedMode = $Mode
    if ($selectedMode -eq "Auto" -or $selectedMode -eq "Grid") {
        $gridContext = Get-GridContext -OcrResult $OcrResult -ReferenceDate $ReferenceDate
        if ($gridContext) {
            return [pscustomobject]@{
                Mode     = "Grid"
                Meetings = Parse-GridMeetings -OcrResult $OcrResult -GridContext $gridContext -StepMinutes $StepMinutes
                Context  = $gridContext
            }
        }
        elseif ($selectedMode -eq "Grid") {
            throw "Grid parsing was requested, but the script could not detect time anchors and day columns."
        }
    }

    return [pscustomobject]@{
        Mode     = "Line"
        Meetings = Parse-LineMeetings -OcrResult $OcrResult -ReferenceDate $ReferenceDate -StepMinutes $StepMinutes
        Context  = $null
    }
}

$resolvedConfigPath = Resolve-ConfigPath -ConfiguredPath $ConfigPath
$config = Load-Config -Path $resolvedConfigPath
$extractionMode = Resolve-ExtractionMode -Config $config -RequestedMode $ParseMode

Assert-ObsidianConfig -Config $config
$categoryRuleSet = Get-MeetingCategoryRuleSet -Config $config

if ($extractionMode -eq "OCR") {
    Initialize-WinRt
}

$useClipboardInput = if ($PSBoundParameters.ContainsKey("FromClipboard")) {
    [bool]$FromClipboard
}
else {
    $true
}

$shouldWriteObsidian = if ($PSBoundParameters.ContainsKey("WriteObsidian")) {
    [bool]$WriteObsidian
}
else {
    $true
}

$inputScreenshotPaths = New-Object System.Collections.Generic.List[string]
if ($ScreenshotPaths) {
    foreach ($path in $ScreenshotPaths) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $inputScreenshotPaths.Add($path)
        }
    }
}

if ($useClipboardInput -and $inputScreenshotPaths.Count -eq 0) {
    $clipboardBaseDirectory = if ($DebugOutputDir) {
        Join-Path $DebugOutputDir "clipboard"
    }
    else {
        Join-Path $PSScriptRoot "output\clipboard"
    }

    $savedClipboardImage = Save-ClipboardImageToFile -Path $ClipboardImagePath -BaseDirectory $clipboardBaseDirectory
    Write-Host ("Saved clipboard image: {0}" -f $savedClipboardImage)
    $inputScreenshotPaths.Add($savedClipboardImage)
}

if ($inputScreenshotPaths.Count -eq 0) {
    throw "No input image was found. Copy an Outlook screenshot to the clipboard or pass -ScreenshotPaths."
}

$referenceDate = if ($PSBoundParameters.ContainsKey("WeekStart")) {
    $WeekStart.Date
}
else {
    (Get-Date).Date
}

$allMeetings = New-Object System.Collections.Generic.List[object]
$debugMetadata = New-Object System.Collections.Generic.List[object]

foreach ($screenshotPath in $inputScreenshotPaths) {
    Write-Host ("Processing screenshot: {0}" -f $screenshotPath)
    if ($extractionMode -eq "OpenAI") {
        $parseResult = Invoke-OpenAiMeetingExtraction -ImagePath $screenshotPath -ReferenceDate $referenceDate -Config $config -CategoryRuleSet $categoryRuleSet -StepMinutes $TimeSnapMinutes
    }
    else {
        $ocrResult = Get-OcrResult -ImagePath $screenshotPath
        $parseResult = Parse-MeetingsFromOcr -OcrResult $ocrResult -ReferenceDate $referenceDate -Mode $ParseMode -StepMinutes $TimeSnapMinutes
    }

    Write-Host ("Detected parse mode: {0}" -f $parseResult.Mode)

    if ($DebugOutputDir) {
        Ensure-ParentDirectory -Path (Join-Path $DebugOutputDir "placeholder.txt")
        $baseName = [IO.Path]::GetFileNameWithoutExtension($screenshotPath)
        if ($extractionMode -eq "OpenAI") {
            Write-Utf8File -Path (Join-Path $DebugOutputDir ($baseName + ".openai-output.json")) -Content $parseResult.RawText
            Write-JsonFile -Path (Join-Path $DebugOutputDir ($baseName + ".openai-response.json")) -Data $parseResult.RawResponse
        }
        else {
            Set-Content -LiteralPath (Join-Path $DebugOutputDir ($baseName + ".ocr.txt")) -Value $ocrResult.Text -Encoding UTF8
            Write-JsonFile -Path (Join-Path $DebugOutputDir ($baseName + ".ocr-lines.json")) -Data $ocrResult.Lines
        }
        if ($parseResult.Context) {
            Write-JsonFile -Path (Join-Path $DebugOutputDir ($baseName + ".parse-context.json")) -Data $parseResult.Context
        }
    }

    $debugMetadata.Add([pscustomobject]@{
        imagePath    = $screenshotPath
        extractionMode = $extractionMode
        parseMode    = $parseResult.Mode
        meetingCount = ($parseResult.Meetings | Measure-Object).Count
    })

    foreach ($meeting in $parseResult.Meetings) {
        $allMeetings.Add($meeting)
    }
}

$dedupeMap = @{}
foreach ($meeting in ($allMeetings | Sort-Object DateText, SortStart, Title)) {
    $key = "{0}|{1}|{2}|{3}" -f $meeting.DateText, $meeting.Start, $meeting.End, $meeting.Title
    if (-not $dedupeMap.ContainsKey($key)) {
        $dedupeMap[$key] = $meeting
    }
}

$dedupedMeetings = $dedupeMap.Values | Sort-Object DateText, SortStart, Title

if (-not $ExportJsonPath) {
    $ExportJsonPath = Join-Path $PSScriptRoot "output\meeting-cache.json"
}

$jsonReady = foreach ($meeting in $dedupedMeetings) {
    [pscustomobject]@{
        date        = $meeting.DateText
        start       = $meeting.Start
        end         = $meeting.End
        title       = $meeting.Title
        categoryKey = $meeting.CategoryKey
        categoryPrefix = $meeting.CategoryPrefix
        categoryProject = $meeting.CategoryProject
        sourceImage = $meeting.SourceImage
        sourceText  = $meeting.SourceText
    }
}

Write-JsonFile -Path $ExportJsonPath -Data $jsonReady

if ($DebugOutputDir) {
    Write-JsonFile -Path (Join-Path $DebugOutputDir "run-summary.json") -Data $debugMetadata
}

Write-Host ""
Write-Host "Detected meetings:"
Write-Host (Format-WeeklyMarkdown -Meetings $dedupedMeetings)
Write-Host ""
Write-Host ("JSON exported to: {0}" -f $ExportJsonPath)

if ($shouldWriteObsidian) {
    Save-ConfigIfMissing -Config $config

    $meetingsByWeek = @{}
    foreach ($meeting in $dedupedMeetings) {
        $week = Get-IsoWeekInfo -Date ([datetime]$meeting.Date)
        $weekKey = "{0}-W{1:D2}" -f $week.Year, $week.Week
        if (-not $meetingsByWeek.ContainsKey($weekKey)) {
            $meetingsByWeek[$weekKey] = New-Object System.Collections.Generic.List[object]
        }
        $meetingsByWeek[$weekKey].Add($meeting)
    }

    foreach ($weekKey in ($meetingsByWeek.Keys | Sort-Object)) {
        $weekMeetings = $meetingsByWeek[$weekKey]
        $representativeDate = [datetime]$weekMeetings[0].Date
        $weeklyPath = Resolve-TemplatePath -VaultRoot $config.vaultRoot -Pattern $config.weeklyNotePattern -Date $representativeDate
        Write-WeeklyMeetingNote -Path $weeklyPath -RepresentativeDate $representativeDate -Meetings $weekMeetings
        Write-Host ("Updated weekly note: {0}" -f $weeklyPath)
    }

    $dailyTemplateFullPath = [IO.Path]::GetFullPath((Join-Path $config.vaultRoot $config.dailyTemplatePath))
    $weeklyFolder = Split-Path -Path ($config.weeklyNotePattern -replace '/', '\') -Parent
    $weeklyFolderForObsidian = ($weeklyFolder -replace '\\', '/').Trim('/')
    $templateBody = Get-DailyTemplateMeetingSectionBody -Heading $config.dailyMeetingHeading -WeeklyFolder $weeklyFolderForObsidian -NoMeetingsLine $config.noMeetingsLine
    Set-BlockBeforeHeading `
        -Path $dailyTemplateFullPath `
        -BeforeHeadingText $config.dailyInsertBeforeHeading `
        -StartMarker $config.dailyStartMarker `
        -EndMarker $config.dailyEndMarker `
        -BodyContent $templateBody `
        -FallbackTitle "Daily template"
    Write-Host ("Updated daily template: {0}" -f $dailyTemplateFullPath)

    foreach ($dayGroup in $dedupedMeetings | Group-Object DateText) {
        $date = [datetime]$dayGroup.Group[0].Date
        $dailyPath = Resolve-TemplatePath -VaultRoot $config.vaultRoot -Pattern $config.dailyNotePattern -Date $date
        if ((Test-Path -LiteralPath $dailyPath) -or $CreateDailyNotes) {
            $dailyContent = Get-DailyMeetingSectionBody -Heading $config.dailyMeetingHeading -MeetingMarkdown (Format-DailyMarkdown -Meetings $dayGroup.Group)
            Set-BlockBeforeHeading `
                -Path $dailyPath `
                -BeforeHeadingText $config.dailyInsertBeforeHeading `
                -StartMarker $config.dailyStartMarker `
                -EndMarker $config.dailyEndMarker `
                -BodyContent $dailyContent `
                -FallbackTitle ($date.ToString("yyyy-MM-dd"))
            Write-Host ("Updated daily note: {0}" -f $dailyPath)
        }
    }
}
