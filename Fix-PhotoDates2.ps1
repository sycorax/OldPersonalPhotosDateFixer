<#
.SYNOPSIS
    Fixes photo and video timestamps using dates from filenames or their containing folders.

.DESCRIPTION
    The script takes no parameters, it asks everything when it starts: which folder (default .\PozeDB),
    what to do, and where to write the list. Subfolders are always included and files are always
    handled newest first, by the date found in the name.

    The name of the file decides the day, except for WhatsApp Image files: their name date is when
    WhatsApp saved the image, not when it was taken, so those files use the date in the nearest dated
    folder instead. The time of day comes from the name when it carries one,
    unless the metadata is within two minutes of it, in which case the metadata keeps its own second.
    A name that gives only a day takes the time from the metadata when that already sits on that day,
    and from the file system date otherwise, so midnight is never invented. Both the metadata and the
    file system date are then brought to that date, each one only if it is off.

    A difference of a whole number of hours is not counted as a difference at all, neither between the
    file system date and the date the file should have, nor between the time in the name and the time
    in the metadata. That is a time zone, one clock wrote local time and the other wrote UTC, and
    neither of the two readings is more right than the other, so nothing is rewritten over it.

    Two files come out of a run, next to each other:
      output.csv  the list of files that need a fix, plus a section listing large file-date differences
      output.log  the counts of the run, what was done to each file, and anything that went wrong

    It then runs in three phases:
      1. Scans for files that are off in the metadata, in the file system date, or in both. A file
         whose name holds a date is measured against that date, a file whose name holds none is
         measured against the date in the name of its folder. The metadata counts as wrong when it
         agrees with neither the file system date nor the date from the name, so a date that is
         merely written with a different rounding is left alone, and also when a date tag holds a
         value that is not a date at all, 0000:00:00 00:00:00 for example, which is repaired. A file
         carrying no date tag at all keeps its bare metadata and only has its file system date
         corrected. It shows the percentage done and an estimate of the time left. Every hit is
         appended to the list right away, so the file can be watched while the scan runs. Esc or Q
         stops the scan and keeps what was found.
      2. Rewrites the list sorted by date.
      3. Walks the list one file at a time, shows the current metadata and the values that would be
         written, and asks for confirmation before touching the file. Each answer is written to the
         log as it happens.

    A date read out of a name is only believed between 1970 and the end of today. Digits that read as
    a date outside that, a serial number or a resolution for instance, are dropped and the file is
    treated as having no date in its name.

    Files whose name holds no date are read as well, and two things are done with them. A metadata or
    file system date that cannot be right, 1950 or later than today, is reported as a warning while
    the scan runs, whether anything can be done about it or not. And the name of the folder is read as
    a rough date. Folder names follow one convention, a dot between the parts of a date and a dash for
    an interval, so 2016.11.05 stands for that one day, 2024.02 for the whole of February, 2016.10-12
    for those three months, 2016.11.05-06 for those two days, 2016.11.28-12.03 for a run of days across
    a month, 2015-2016 for those two years and 2016 for the whole of it. The far end of an interval may
    be written out in full, 2016.11-2017.01 or 2016.12.28-2017.01.03, and spaces around the dash say
    nothing. A folder name that breaks the convention, 20161105 or 2016-11-05 or 2016_11, is reported as
    a warning while the scan runs, and the date in it is read all the same, since 20180421 means that one
    day and nothing else. A file whose dates sit more than 30 days outside what its folder says is
    listed like the rest and offered for a fix, always one file at a time since a folder name is a hint
    and not a timestamp. Metadata holding no date tag at all is left alone for these files, there only
    the file system date is corrected.

    When both dates of such a file are outside the interval, the day written can only be an end of it,
    and for a folder naming a whole year or wider that would put every picture in it on the first of
    January and lose the order they were taken in. So a file like that is put in the list with its
    reason and nothing is written for it, no question asked, whenever the interval is wider than a
    month. A folder that names one day, one month or a couple of weeks still writes.

    WhatsApp Image filenames are treated as having no date, regardless of the date and time in their
    name. They are judged by the nearest dated folder, so the date WhatsApp saved the image is never
    written into its metadata or file system dates. The csv also has a separate information section
    for files judged by a folder whose file system date is more than 30 days outside that folder's
    date interval.

    The name of a folder says which day, never which hour, so a date written from one moves the day and
    nothing else: metadata reading 2012:10:22 19:32:51 in a folder called 2011 is written as
    2011:01:01 19:32:51, keeping the second the camera recorded, and the file system date keeps its own
    time of day the same way. The two then differ in the time of day, each holding on to the one it had.

    Unless the metadata reads as a camera clock counting from its own reset, which takes two things: a
    date in the first days of a January of a year the folder says nothing about, a GoPro coming back at
    2012:01:01 after the battery is out, and a time of day more than four hours from the time of day of
    the file. Both together mean the clock knows neither the day nor the hour: 2018:01:01 04:30:40 next
    to a file written at 14:54:52 is only how long the camera had been running, so the file system date
    is taken whole, 2019:06:01 14:54:52. A reset clock reading 2013:01:05 21:09:36 next to a file written
    at 22:27:23 still keeps its 21:09:36, an hour and a bit is a time zone and a drifting clock, and that
    hour is what puts the photos of an evening back in the order they were taken.

    Uses exiftool at C:\Prog\exiftool\exiftool.exe, change the $ExifTool setting below if it moves.

.EXAMPLE
    .\Fix-PhotoDates.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Settings, change these if needed
# ---------------------------------------------------------------------------

$ExifTool       = "C:\Prog\exiftool\exiftool.exe"
$DefaultFolder  = ".\PozeDB"
$DefaultOutput  = "output.csv"

# Excel cuts a csv into columns on the list separator of the Windows locale, a comma in English and
# a semicolon in most European ones, so the separator of the current culture is used. Put a "," or a
# ";" here instead if the file has to be read by something with a fixed idea about it.
$CsvDelimiter   = (Get-Culture).TextInfo.ListSeparator

# How many files are read per exiftool call. Starting the tool is what costs time, not reading one file,
# so a bigger number is faster, but progress and the stop key only react once per group.
$ChunkSize      = 200

# Two dates this many seconds apart or less count as the same date, they are one date written with a
# different rounding or with a time zone read slightly differently. Everything below compares dates
# with this margin, so nothing is listed and nothing is written over a difference this small.
$ConflictSeconds = 80

# The file system date is brought to the date from the name whenever the two differ. Set this to
# $false to leave file system dates alone and only ever touch the metadata.
$FixFileSystemDate = $true

# A file system date sitting a whole number of hours away from the date the file should have is not a
# mistake, it is a time zone: the camera wrote one zone into the picture and Windows wrote another one
# into the file, or summer time came in between. A difference that close to a whole hour, up to this
# many hours, is left as it is and the file is not listed for it. Three hours covers a zone two or
# three hours off UTC with summer time on top. Put 1 here to excuse only the plain one hour case, or 0
# to switch this off and have every difference corrected.
#
# Whole hours means whole hours, give or take the $ConflictSeconds above: one hour and ten seconds off
# is a time zone, one hour and a half off is not, so a real mistake of half an hour is still found.
#
# The same goes for the time in the name against the time in the metadata. A whole hour between those
# two is one moment written in two zones, the phone naming the file in one and storing the other, so
# the metadata keeps its own reading of it instead of being moved to a zone just as arbitrary.
$TimeZoneHoursMax = 3

# How many files were left alone because their file system date was off by a whole number of hours,
# and how many kept their metadata because it was a whole number of hours from the time in the name.
$script:TimeZoneLeft = 0
$script:TimeZoneKept = 0

# A name carrying a time usually got it from the metadata in the first place, and phones round that
# second differently in the two places: WP_20141109_16_01_07_Pro.jpg holds 16:01:03 in its metadata.
# A name time this close to the metadata date counts as confirming it, so the metadata is kept as it
# stands and the photo is not rewritten over a few seconds. Further apart than this, the name wins.
$NameTimeSlackSeconds = 120

# A file with no date tag at all in its metadata has nothing wrong there, only nothing at all, so
# its metadata is left as it is and just its file system date is corrected. Set this to $true to
# write the date from the name into it as well.
#
# This is not about a tag holding 0000:00:00 00:00:00, which many phone videos carry. That is a
# written value that is not a date, so it is always corrected, whatever this setting says.
$IncludeFilesWithoutExifDate = $false

# Files whose name holds no date are read too, so that a date which cannot be right is reported and
# so that the name of the folder can be used as a rough date for them. Set this to $false to leave
# them unread, which roughly halves the time the scan takes, since they are usually about half of
# what is found.
$CheckFilesWithoutDateInName = $true

# What the name of a folder says is an interval, not a moment: 2016.11.05 is one day, 2024.02 is the
# whole of February, 2016 is a whole year. A date sitting inside that interval is not off at all, and
# only a date this many days outside it is treated as wrong. The margin is wide on purpose, a folder
# name is a hint about when the pictures in it were taken and not a timestamp.
$FolderDayTolerance = 30

# For the same reason a date taken from a folder name is always confirmed one file at a time, even in
# the mode that fixes everything without asking. Set this to $false to have those written like the
# rest. Answering a to a file like this switches the asking off for the remaining ones as well.
$AlwaysAskForFolderDates = $true

# A folder named after a whole year says too little to place a photo in time. When neither of the two
# dates of a file sits inside the interval, the only day left to write is the edge of it, the first of
# January for a year, and a folder holding a week of a trip would have all of its pictures on that one
# day: the days would be gone and with them the order the pictures were taken in. So a file like that
# is put in the list with its reason and nothing is offered for it, no question and no write, as long
# as the interval is wider than this many days. A month wide folder still writes, the edge of it is at
# most a few weeks out. Put a very large number here to have every folder write, whatever it says.
#
# This only holds when both dates have to be thrown away. Whenever one of the two does sit inside the
# interval it is a real timestamp, not a guess, and the other one is brought to it however wide the
# folder is, because that is copying a date and not inventing one.
$FolderMaxIntervalDays = 31

# How many files were listed for the reason above and left alone.
$script:FolderReported = 0

# A camera that loses power forgets the date and starts again at the first of January of whatever year
# its firmware was built around, 2012 on an older GoPro and 2018 on a newer one, and from there it
# counts real time. So a date in the first days of a January that has nothing to do with the folder the
# file sits in is not a date the camera knew, it is a clock counting from its own reset, and the time of
# day in it is worth no more than the day: GH010025.MP4 in a folder called 2019.06 holds
# 2018:01:01 04:30:40 while the file itself was written at 14:54:52, and 04:30:40 is simply how long
# the camera had been running.
#
# A metadata date like that is not kept for its time of day and is not used to pick the day either, the
# file system date is preferred to it. This many days into January counts as the reset, so a camera
# switched on a few times after losing its clock is still recognised. Put 0 here to switch this off and
# have every metadata date keep its own time of day.
$ResetClockDays = 7

# A clock that lost its date has not always lost the time of day with it. A camera keeps counting real
# time while it is switched off, so one reset a week before the picture reads 2013:01:05 and still says
# an evening hour that is only a time zone away from the truth, and that hour is worth keeping: it is
# what puts sixteen photos of an evening in the order they were taken instead of all on the second the
# folder was copied. What is not worth keeping is an hour that is half a day out, 04:30:40 against a
# file written at 14:54:52, because then the clock is counting from its own reset and nothing else.
#
# So the time of day of a reset clock is thrown away only when it is further than this many hours from
# the time of day of the file, measured the short way round midnight. Four hours covers a time zone and
# the drift of a cheap clock. Put 0 here to throw away the time of day of every reset clock.
$LostClockHours = 4

# How many files had a metadata date that reads as a clock which lost its date.
$script:LostClocks = 0

# How many warnings about impossible dates are printed while scanning. The rest only go to the log,
# so a folder full of them cannot bury the progress lines. All of them are counted.
$MaxWarningsShown = 20

# How many warnings were given, reported at the end of the scan.
$script:DateWarnings = 0

# How many files with no date in the name were listed because of what their folder is called.
$script:FolderFixes = 0

# The interval a folder stands for is worked out once per folder, not once per file in it.
$script:FolderDateCache = @{}

# And a folder whose name breaks the naming convention is complained about once, not once per file.
$script:FolderWarned = @{}

# A pattern can line up with digits that are not a date at all, a serial number or a resolution for
# instance, and that is where years like 1954, 1928 or 2084 come from. A date read out of a name is
# only believed inside this range: after the start date and not later than the end of today, since a
# photo cannot have been taken tomorrow. Outside it the match is dropped and the next pattern tried.
$EarliestDate = [datetime] "1970-01-01"
$LatestDate   = (Get-Date).Date.AddDays(1).AddSeconds(-1)

# How many names held a date outside that range, reported at the end of the scan.
$script:ImplausibleNames = 0

# Videos keep their QuickTime dates in UTC by the standard, but many phones write local time there
# anyway. Set this to $true if every video comes out with a conflict that is exactly your time zone
# offset, it makes exiftool convert those dates to local time when reading and back when writing.
$QuickTimeUtc = $false

# The names of the files to read are handed to exiftool in an argument file, that keeps names with
# diacritics intact. Some builds ignore -@ and read the argument file itself, that is detected once
# and the names go on the command line from then on.
$script:UseArgFile = $true

# How many groups had trouble, so the same warning is not printed thousands of times.
$script:ExifTroubleGroups = 0

# The first group that goes wrong is described in full in the log file, once.
$script:ExifDiagDone = $false

# What the last exiftool call was given and what it answered, kept for that description.
$script:LastExifArgs = @()
$script:LastExifRaw  = ""

# The log file, opened once the name of the list file is known.
$script:LogWriter = $null

# Read the output of exiftool as UTF-8, otherwise names with diacritics come back mangled.
try
{
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
}
catch
{
    # Not every host allows this, it only affects how names are printed.
}

$Extensions = @(".jpg", ".jpeg", ".png", ".heic", ".heif", ".tif", ".tiff", ".dng", ".cr2", ".cr3", ".nef", ".arw",
                ".mp4", ".mov", ".avi", ".mts", ".m4v", ".3gp", ".mkv")

$VideoExtensions = @(".mp4", ".mov", ".avi", ".mts", ".m4v", ".3gp", ".mkv")

# Each pattern must expose named groups y, mo, d and optionally h, mi, s. They are tried in this
# order and the first one that gives a real date wins, so the ones that carry a time come first.
# A pattern that gives an impossible date, 25 as the hour for example, is passed over.
#
# Every one of them starts at the front of a number, never in the middle of one: the identifier in
# 67874694_2197506223692157_995977790383718400_n.jpg holds the digits 19750622 in the middle of it and
# they are not the 22nd of June 1975, they are a part of a number that Facebook made up. The two forms
# written with no separator at all are also required to end where the number ends, so eight digits are
# read as a day only when eight digits is all there is, and a longer run of them stays a number. A
# time in a name is allowed to carry more digits after the seconds though, those are the thousandths
# some phones write, and dropping them is right where dropping the day would not be.

# The pieces the patterns are built from, named so that no line of them grows past the width of the
# rest of the script: the front and the back of a number, a day, a time of day and what may sit
# between the two.
$NdOpen = '(?<!\d)'
$NdShut = '(?!\d)'
$NdDay  = '(?<y>(19|20)\d{2})[\-\._]?(?<mo>\d{2})[\-\._]?(?<d>\d{2})'
$NdTime = '(?<h>\d{2})[\-\._: ]?(?<mi>\d{2})[\-\._: ]?(?<s>\d{2})'
$NdAt   = '[ _\-T\.]+(?:[Aa][Tt][ _\-]+)?'

$DatePatterns = @(
    # A day and a time, with any of the usual separators or none inside each of the two, and the word
    # "at" allowed between them. This one covers IMG_20240428_150842.jpg, WP_20141115_18_36_44_Pro.jpg,
    # WhatsApp Video 2026-06-14 at 09.05.15.mp4, 2024-04-28 15.08.42.jpg and 2024_04_28T15:08:42.jpg.
    "$NdOpen$NdDay$NdAt$NdTime",
    # 14 digits in a row, no separator anywhere, as in IMG20240428150842.jpg.
    "$NdOpen" + '(?<y>(19|20)\d{2})(?<mo>\d{2})(?<d>\d{2})(?<h>\d{2})(?<mi>\d{2})(?<s>\d{2})',
    "$NdOpen" + '(?<y>(19|20)\d{2})[\-\._](?<mo>\d{2})[\-\._](?<d>\d{2})' + "$NdShut",
    "$NdOpen" + '(?<y>(19|20)\d{2})(?<mo>\d{2})(?<d>\d{2})' + "$NdShut"
)

# The name of a folder says when the pictures in it were taken, and it says it by one convention: a
# dot between the parts of a date, a dash for an interval and nothing else. So:
#
#   2016                    the whole year
#   2016.11                 the whole of November
#   2016.11.05              that one day
#   2016.10-12              October to December
#   2016.11.05-06           those two days
#   2016.11.28-12.03        from one day to another across a month
#   2016.11.28-2017.01.03   the same across a year, the far end written out in full
#   2016.11-2017.01         from one month to another across a year
#   2015-2016               those two years
#
# Each pattern gives an interval and they are tried in this order, the longest form first, so a name is
# never cut short: 2016.11.05-2016.11.08 must not read as the fifth alone. Spaces around the dash are
# allowed, they say nothing. A pattern exposes y and, depending on its kind, mo, d, y2, mo2 and d2.
#
# The pieces the patterns are built from are named so that each pattern still fits on one line: the
# year and the year an interval runs to, the dash between the two ends, and the two edges that keep a
# pattern from biting into a longer run of digits, so that 20161105 is never read as the year 2016.
$FdOpen  = '(?<!\d)'
$FdShut  = '(?!\d)'
$FdYear  = '(?<y>(19|20)\d{2})'
$FdYear2 = '(?<y2>(19|20)\d{2})'
$FdTo    = '\s*-\s*'
$FolderDatePatterns = @(
    [pscustomobject] @{ Kind = "daysfull";   Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})\.(?<d>\d{2})$FdTo$FdYear2\.(?<mo2>\d{2})\.(?<d2>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "daysacross"; Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})\.(?<d>\d{2})$FdTo(?<mo2>\d{2})\.(?<d2>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "days";       Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})\.(?<d>\d{2})$FdTo(?<d2>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "day";        Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})\.(?<d>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "monthsfull"; Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})$FdTo$FdYear2\.(?<mo2>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "months";     Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})$FdTo(?<mo2>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "month";      Pattern = "$FdOpen$FdYear\.(?<mo>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "years";      Pattern = "$FdOpen$FdYear$FdTo$FdYear2$FdShut" }
)

# The year on its own, which is both a name in the convention and what is left of a name that breaks
# it, so it is tried after all of the above and after the check below.
$FolderYearPattern = "$FdOpen$FdYear$FdShut"

# A date written some other way, with any separator or none at all: 20180421, 2016-11-05, 2016_11,
# 2018 04. These are tried only after every pattern above has failed, and a name read by one of them
# is warned about. It is still read though: 20180421 ASTRO means the 21st of April 2018 and nothing
# else, and throwing that away to fall back on whatever folder sits above it would put the pictures in
# the wrong month while the warning says only that the name is written oddly.
#
# The day comes first here too, so 2016-11-05 is the fifth of November and not November to May. The
# month needs a separator, otherwise six digits of a serial number would read as one.
$FolderLoosePatterns = @(
    [pscustomobject] @{ Kind = "day";   Pattern = "$FdOpen$FdYear[-._ ]?(?<mo>\d{2})[-._ ]?(?<d>\d{2})$FdShut" },
    [pscustomobject] @{ Kind = "month"; Pattern = "$FdOpen$FdYear[-._ ](?<mo>\d{2})$FdShut" }
)

# A folder name holding a date that is not written the way the convention says. These are only looked
# for when none of the patterns above fit, and the first one that matches names the problem in the
# warning. The date itself is still read, by the loose patterns above or, failing those, as the year.
$FolderNameProblems = @(
    [pscustomobject] @{ Pattern = '(?<!\d)(19|20)\d{4,}(?!\d)';            Problem = "a date in one run of digits, the dots are missing" },
    [pscustomobject] @{ Pattern = '(?<!\d)(19|20)\d{2}-\d{1,2}(?!\d)';     Problem = "a dash where a dot belongs, a dash stands for an interval" },
    [pscustomobject] @{ Pattern = '(?<!\d)(19|20)\d{2}[_ /]\d{1,2}(?!\d)'; Problem = "a date written with something other than a dot" },
    [pscustomobject] @{ Pattern = '(?<!\d)(19|20)\d{2}\.\d{1,2}(?!\d)';    Problem = "a month or an interval that cannot be read" }
)

# ---------------------------------------------------------------------------
# Small input helpers
# ---------------------------------------------------------------------------

function Read-Text
{
    param([string] $Prompt, [string] $Default)

    $answer = Read-Host "$Prompt [$Default]"

    if ([string]::IsNullOrWhiteSpace($answer))
    {
        return $Default
    }

    return $answer.Trim().Trim('"')
}

function Read-YesNo
{
    param([string] $Prompt, [bool] $Default)

    if ($Default)
    {
        $hint = "Y/n"
    }
    else
    {
        $hint = "y/N"
    }

    while ($true)
    {
        $answer = (Read-Host "$Prompt [$hint]").Trim().ToLower()

        if ($answer -eq "")
        {
            return $Default
        }

        if ($answer -in @("y", "yes"))
        {
            return $true
        }

        if ($answer -in @("n", "no"))
        {
            return $false
        }
    }
}

function Read-Choice
{
    param([string] $Prompt, [string[]] $Valid, [string] $Default)

    while ($true)
    {
        $answer = (Read-Host "$Prompt [$Default]").Trim().ToLower()

        if ($answer -eq "")
        {
            return $Default
        }

        if ($Valid -contains $answer)
        {
            return $answer
        }
    }
}

function Write-Log
{
    param([string] $Text)

    if ($null -ne $script:LogWriter)
    {
        # The stamp goes in a variable of its own, a call in brackets as the first value of -f
        # is taken for the whole right side of the operator and the second slot stays empty.
        $stamp = Get-Date -Format "HH:mm:ss"
        $script:LogWriter.WriteLine("$stamp  $Text")
    }
}

function Write-Both
{
    param([string] $Text, [string] $Color = "Gray")

    Write-Host $Text -ForegroundColor $Color
    Write-Log $Text
}

function Get-Shortened
{
    param([string] $Text, [int] $Length)

    if ([string]::IsNullOrEmpty($Text))
    {
        return "(empty)"
    }

    if ($Text.Length -le $Length)
    {
        return $Text
    }

    return ($Text.Substring(0, $Length) + " ...[" + $Text.Length + " characters in total]")
}

function Test-StopKey
{
    try
    {
        if (-not [Console]::KeyAvailable)
        {
            return $false
        }
    }
    catch
    {
        # No interactive console, for example when the output is redirected.
        return $false
    }

    $key = [Console]::ReadKey($true)

    if ($key.Key -eq [ConsoleKey]::Escape -or $key.Key -eq [ConsoleKey]::Q)
    {
        return $true
    }

    return $false
}

function Format-Duration
{
    param([double] $Seconds)

    if ($Seconds -lt 60)
    {
        return ("{0:N0} s" -f $Seconds)
    }

    $span = [timespan]::FromSeconds([Math]::Round($Seconds))
    return ("{0:D2}:{1:D2} min" -f ($span.Hours * 60 + $span.Minutes), $span.Seconds)
}

# ---------------------------------------------------------------------------
# exiftool and date helpers
# ---------------------------------------------------------------------------

function Test-ExifTool
{
    param([string] $Tool)

    try
    {
        $version = & $Tool -ver 2>$null
    }
    catch
    {
        $version = $null
    }

    if (-not $version)
    {
        throw "exiftool did not answer at '$Tool'. Change the `$ExifTool setting at the top of the script to the right path."
    }

    return $version.Trim()
}

function Get-DateFromName
{
    param([string] $Name)

    # WhatsApp dates are save dates, not capture dates, so these files must be judged by their folder.
    if ($Name -match '^WhatsApp Image\b')
    {
        return $null
    }

    # Set when a pattern read a date that cannot be one, so the name can be counted once at the end
    # however many patterns line up with the same digits.
    $refused = $false

    foreach ($pattern in $DatePatterns)
    {
        $match = [regex]::Match($Name, $pattern)

        while ($match.Success)
        {
            # Every place in the name where the pattern fits is tried, not only the first one, so a
            # name carrying a serial number before the real date still gives up that date. The next
            # match is taken now because nothing below may jump over this line.
            $current = $match
            $match   = $match.NextMatch()

            $hasTime = $current.Groups['h'].Success

            $hour   = if ($hasTime)                      { $current.Groups['h'].Value }  else { "00" }
            $minute = if ($current.Groups['mi'].Success) { $current.Groups['mi'].Value } else { "00" }
            $second = if ($current.Groups['s'].Success)  { $current.Groups['s'].Value }  else { "00" }

            $text   = "{0}-{1}-{2} {3}:{4}:{5}" -f $current.Groups['y'].Value, $current.Groups['mo'].Value,
                                                   $current.Groups['d'].Value, $hour, $minute, $second
            $parsed = [datetime]::MinValue

            if ([datetime]::TryParseExact($text, "yyyy-MM-dd HH:mm:ss", $null, [System.Globalization.DateTimeStyles]::None, [ref] $parsed))
            {
                if ($parsed -lt $EarliestDate -or $parsed -gt $LatestDate)
                {
                    # Digits that read as a date but cannot be one, so this is not a date either.
                    # The rest of the name and the other patterns still get their turn.
                    $refused = $true
                }
                else
                {
                    # HasTime says whether the time comes from the name or is just midnight standing
                    # in for a time the name never held. Midnight must not overwrite a good time.
                    return [pscustomobject] @{ Date = $parsed; HasTime = $hasTime }
                }
            }
        }
    }

    if ($refused)
    {
        $script:ImplausibleNames++
    }

    return $null
}

function Get-DayOrNothing
{
    param([string] $Year, [string] $Month, [string] $Day)

    # The one day three groups of digits spell, or nothing when they spell no day at all, the 31st of
    # February for one, or a month numbered 13.
    $stamp  = "{0}-{1}-{2}" -f $Year, $Month, $Day
    $parsed = [datetime]::MinValue

    if (-not [datetime]::TryParseExact($stamp, "yyyy-MM-dd", $null, [System.Globalization.DateTimeStyles]::None, [ref] $parsed))
    {
        return $null
    }

    return $parsed
}

function Get-IntervalFromMatch
{
    param([string] $Kind, $Match)

    # The interval one match of one folder pattern stands for, or nothing when the digits it caught are
    # no date, are the wrong way round or fall outside the years a photo can carry. The two ends and
    # the text shown for them are all it gives, the caller says what to do with them.
    $year  = [int] $Match.Groups['y'].Value
    $first = $null
    $last  = $null
    $shown = ""

    if ($Kind -eq "day" -or $Kind -eq "days" -or $Kind -eq "daysacross" -or $Kind -eq "daysfull")
    {
        $first = Get-DayOrNothing -Year $Match.Groups['y'].Value -Month $Match.Groups['mo'].Value -Day $Match.Groups['d'].Value

        if ($null -eq $first)
        {
            return $null
        }

        $last = $first

        if ($Kind -eq "days")
        {
            $last = Get-DayOrNothing -Year $Match.Groups['y'].Value -Month $Match.Groups['mo'].Value -Day $Match.Groups['d2'].Value
        }
        elseif ($Kind -eq "daysacross")
        {
            $last = Get-DayOrNothing -Year $Match.Groups['y'].Value -Month $Match.Groups['mo2'].Value -Day $Match.Groups['d2'].Value
        }
        elseif ($Kind -eq "daysfull")
        {
            $last = Get-DayOrNothing -Year $Match.Groups['y2'].Value -Month $Match.Groups['mo2'].Value -Day $Match.Groups['d2'].Value
        }

        if ($null -eq $last -or $last -lt $first)
        {
            return $null
        }

        $shown = $first.ToString("yyyy-MM-dd")

        if ($last -ne $first)
        {
            $shown = "{0} to {1}" -f $first.ToString("yyyy-MM-dd"), $last.ToString("yyyy-MM-dd")
        }
    }
    elseif ($Kind -eq "month" -or $Kind -eq "months" -or $Kind -eq "monthsfull")
    {
        $month    = [int] $Match.Groups['mo'].Value
        $upto     = $month
        $uptoYear = $year

        if ($Kind -eq "months")
        {
            $upto = [int] $Match.Groups['mo2'].Value
        }
        elseif ($Kind -eq "monthsfull")
        {
            $upto     = [int] $Match.Groups['mo2'].Value
            $uptoYear = [int] $Match.Groups['y2'].Value
        }

        if ($month -lt 1 -or $month -gt 12 -or $upto -lt 1 -or $upto -gt 12)
        {
            return $null
        }

        $first = [datetime]::new($year, $month, 1)
        $last  = ([datetime]::new($uptoYear, $upto, 1)).AddMonths(1).AddDays(-1)

        if ($last -lt $first)
        {
            # October to March, an interval running backwards, which is a name written wrong.
            return $null
        }

        $shown = $first.ToString("yyyy-MM")

        if ($last -gt ($first.AddMonths(1).AddDays(-1)))
        {
            $shown = "{0} to {1}" -f $first.ToString("yyyy-MM"), $last.ToString("yyyy-MM")
        }
    }
    elseif ($Kind -eq "years")
    {
        $upto = [int] $Match.Groups['y2'].Value

        if ($upto -lt $year)
        {
            return $null
        }

        $first = [datetime]::new($year, 1, 1)
        $last  = [datetime]::new($upto, 12, 31)
        $shown = "{0} to {1}" -f $year, $upto
    }
    else
    {
        $first = [datetime]::new($year, 1, 1)
        $last  = [datetime]::new($year, 12, 31)
        $shown = "$year"
    }

    if ($first -lt $EarliestDate -or $first -gt $LatestDate)
    {
        # A year no photo can carry, so those digits are something else, a house number maybe.
        return $null
    }

    # The last second of the last day, so a picture taken in the evening of it still counts as inside.
    # The interval never reaches into the future either, nothing was taken tomorrow.
    $end = $last.Date.AddDays(1).AddSeconds(-1)

    if ($end -gt $LatestDate)
    {
        $end = $LatestDate
    }

    return [pscustomobject] @{ First = $first; Last = $end; Text = $shown; Problem = "" }
}

function Get-DateIntervalFromText
{
    param([string] $Text)

    # The interval one folder name stands for, or nothing when it holds no date at all. First and Last
    # are the two ends of it, Last being the last second of the last day, Text is the date as it was
    # read, which is what the list and the questions show, and Problem says how the name breaks the
    # convention, empty when it keeps it. A name that breaks it is still read for its year, so First
    # and Problem can both be filled in, and First is empty when not even the year survived.
    if ([string]::IsNullOrWhiteSpace($Text))
    {
        return $null
    }

    # A name can be written the way the convention asks and still hold digits that are no date, the
    # 31st of February for one, or an interval running backwards. What is left of such a name is read
    # by a wider pattern, 2016.02.30 as February and 2016.12-10 as December, and the refusal is kept
    # here so that it is still complained about afterwards. The first one met is the one reported.
    $problem = ""

    foreach ($pattern in $FolderDatePatterns)
    {
        $match = [regex]::Match($Text, $pattern.Pattern)

        while ($match.Success)
        {
            # Every place a pattern fits is tried, so a folder called "Album 3 2016.11" gives up its
            # date as well. The next match is taken now because nothing below may jump over this line.
            $current = $match
            $match   = $match.NextMatch()

            $found = Get-IntervalFromMatch -Kind $pattern.Kind -Match $current

            if ($null -ne $found)
            {
                $found.Problem = $problem
                return $found
            }

            if ($problem -eq "")
            {
                $problem = "$($current.Value) is no date, an impossible day, an interval the wrong way round or a year too far off"
            }
        }
    }

    # Nothing in the name is written the way the convention asks. What follows is read all the same,
    # only complained about, so the name is first held against the ways a date is usually written wrong
    # and the first of those that fits is what the warning will say.
    if ($problem -eq "")
    {
        foreach ($bad in $FolderNameProblems)
        {
            if ([regex]::IsMatch($Text, $bad.Pattern))
            {
                $problem = $bad.Problem
                break
            }
        }
    }

    # A date written some other way is still a date and is taken as one. Only when even that fails does
    # the name fall back to its year, which is the last thing about it that cannot be misread.
    foreach ($pattern in $FolderLoosePatterns)
    {
        $match = [regex]::Match($Text, $pattern.Pattern)

        while ($match.Success)
        {
            $current = $match
            $match   = $match.NextMatch()

            $found = Get-IntervalFromMatch -Kind $pattern.Kind -Match $current

            if ($null -ne $found)
            {
                if ($problem -eq "")
                {
                    $problem = "a date not written with dots"
                }

                $found.Problem = $problem
                return $found
            }
        }
    }

    $match = [regex]::Match($Text, $FolderYearPattern)

    while ($match.Success)
    {
        $current = $match
        $match   = $match.NextMatch()

        $found = Get-IntervalFromMatch -Kind "year" -Match $current

        if ($null -ne $found)
        {
            $found.Problem = $problem
            return $found
        }
    }

    if ($problem -ne "")
    {
        # Digits meant as a date, written so that not even the year of them can be picked out, which is
        # what happens when the whole date is one run of digits. Worth a warning, worth nothing else.
        return [pscustomobject] @{ First = $null; Last = $null; Text = ""; Problem = $problem }
    }

    return $null
}

function Get-FolderDateInterval
{
    param([string] $Folder, [string] $Root)

    # The folders are read from the one holding the file upwards and the first name that gives a date
    # wins, so the nearest folder decides and a year folder above it only serves the ones that say
    # nothing themselves. Nothing above the folder being scanned is ever read. A name holding a date
    # that is not written the way the convention asks is complained about, once per folder however many
    # files sit in it, and what could still be read out of it is used.
    if ($script:FolderDateCache.ContainsKey($Folder))
    {
        return $script:FolderDateCache[$Folder]
    }

    $found   = $null
    $current = $Folder

    while (-not [string]::IsNullOrEmpty($current) -and $current.Length -ge $Root.Length)
    {
        $read = Get-DateIntervalFromText -Text ([System.IO.Path]::GetFileName($current))

        if ($null -ne $read)
        {
            if ($read.Problem -ne "" -and -not $script:FolderWarned.ContainsKey($current))
            {
                $script:FolderWarned[$current] = $true

                if ($null -eq $read.First)
                {
                    Write-DateWarning "the folder $current does not keep to the convention, $($read.Problem), so no date is read from it"
                }
                else
                {
                    $spelled = $read.Text.Replace("-", ".")
                    $note    = "the folder $current does not keep to the convention, $($read.Problem)"

                    Write-DateWarning "$note, it is read as $spelled and should say so in the name"
                }
            }

            if ($null -ne $read.First)
            {
                $found = $read
                break
            }
        }

        $parent = [System.IO.Path]::GetDirectoryName($current)

        if ([string]::IsNullOrEmpty($parent) -or $parent -eq $current)
        {
            break
        }

        $current = $parent
    }

    $script:FolderDateCache[$Folder] = $found
    return $found
}

function Get-DaysOutside
{
    param([datetime] $Stamp, $Interval)

    # How far a date sits outside the interval its folder stands for, in days. A date anywhere inside
    # it is not off at all, whichever end of it the date is near.
    if ($Stamp -lt $Interval.First)
    {
        return ($Interval.First - $Stamp).TotalDays
    }

    if ($Stamp -gt $Interval.Last)
    {
        return ($Stamp - $Interval.Last).TotalDays
    }

    return 0
}

function Get-ClampedDate
{
    param([datetime] $Stamp, $Interval)

    # The nearest day inside the interval, keeping the time of day of the date given, so a picture
    # holds on to the hour it was taken at even when its day has to come from the name of a folder.
    if ($Stamp -lt $Interval.First)
    {
        return $Interval.First.Date.Add($Stamp.TimeOfDay)
    }

    if ($Stamp -gt $Interval.Last)
    {
        return $Interval.Last.Date.Add($Stamp.TimeOfDay)
    }

    return $Stamp
}

function Get-TimeOfDayGap
{
    param([datetime] $One, [datetime] $Other)

    # How far apart two times of day are, in hours, measured the short way round midnight: 23:04 and
    # 00:03 are one hour apart and not twenty three. The day each of them belongs to says nothing here,
    # the question is only whether the two clocks agree about the hour.
    $gap = ($One.TimeOfDay - $Other.TimeOfDay).TotalHours

    if ($gap -gt 12)
    {
        $gap -= 24
    }

    if ($gap -lt -12)
    {
        $gap += 24
    }

    return [Math]::Abs($gap)
}

function Test-LostClockDate
{
    param([datetime] $Stamp, [datetime] $FileStamp, $Interval)

    # True when a metadata date is a camera clock counting from its own reset and nothing more, so that
    # not even the time of day in it is worth keeping. Two things have to hold together for that. Only a
    # date that already sits well outside the interval is asked about, the caller sees to that.
    #
    # The date falls in the first days of a January belonging to no year the folder speaks of. A GoPro
    # comes back at 2012:01:01 after the battery is out, a newer one at 2018:01:01, and it counts on from
    # there, so a file recorded a few days later reads 2012:01:05. The year of the folder is the way out
    # of this: a photo really taken on the first of January sits in a folder that says so, 2013.01.01 or
    # 2013, and it is left alone by all of this. See $ResetClockDays.
    #
    # And the time of day it holds is far from the time of day of the file itself. A clock that lost only
    # the date keeps counting real time and still names the right hour give or take a time zone, and that
    # hour is the one thing about it worth keeping: it is what puts the photos of an evening in the order
    # they were taken. An hour half a day out, 04:30:40 against a file written at 14:54:52, is the time
    # since the reset and nothing else. See $LostClockHours.
    if ($ResetClockDays -le 0)
    {
        return $false
    }

    if ($Stamp.Month -ne 1 -or $Stamp.Day -gt $ResetClockDays)
    {
        return $false
    }

    if ($Stamp.Year -eq $Interval.First.Year -or $Stamp.Year -eq $Interval.Last.Year)
    {
        return $false
    }

    return ((Get-TimeOfDayGap -One $Stamp -Other $FileStamp) -gt $LostClockHours)
}

function Test-TimeZoneShift
{
    param([double] $Seconds)

    # True when a difference is close enough to a whole number of hours to be a time zone rather than a
    # mistake. The margin around each whole hour is the same $ConflictSeconds used everywhere else, so
    # a rounding on top of the hour still counts, while a difference of half an hour does not.
    if ($TimeZoneHoursMax -le 0)
    {
        return $false
    }

    $away = [Math]::Abs($Seconds)

    for ($hours = 1; $hours -le $TimeZoneHoursMax; $hours++)
    {
        if ([Math]::Abs($away - ($hours * 3600)) -le $ConflictSeconds)
        {
            return $true
        }
    }

    return $false
}

function Write-DateWarning
{
    param([string] $Text)

    # Every warning goes to the log, only the first few to the screen: one folder of bad videos would
    # otherwise push the progress lines out of the window. The total comes at the end of the scan.
    $script:DateWarnings++
    Write-Log "WARNING $Text"

    if ($script:DateWarnings -le $MaxWarningsShown)
    {
        Write-Host "  warning: $Text" -ForegroundColor Yellow
    }
    elseif ($script:DateWarnings -eq ($MaxWarningsShown + 1))
    {
        Write-Host "  further warnings only go to the log, they are counted at the end of the scan." -ForegroundColor Yellow
    }
}

function New-FolderCandidate
{
    param($File, $Interval, $Info, [bool] $IsVideo, $Read)

    # A file whose name holds no date. Two separate things happen here. A date that cannot be right is
    # reported whatever else is known about the file, that is worth saying even when nothing can be
    # done about it. And when the folder the file sits in is named after a date, that date is what the
    # file is measured against, loosely: only more than $FolderDayTolerance days outside the interval
    # the folder stands for counts as wrong. Nothing is returned when there is nothing to propose.
    $exifDate = $Read.Date
    $fileDate = $File.LastWriteTime
    $name     = $File.Name

    if ($Read.Bad)
    {
        Write-DateWarning -Text "$name : $($Read.BadTag) holds $($Read.BadValue), which is not a date"
    }
    elseif ($null -ne $exifDate -and ($exifDate -lt $EarliestDate -or $exifDate -gt $LatestDate))
    {
        Write-DateWarning -Text "$name : the metadata date $($exifDate.ToString('yyyy-MM-dd HH:mm:ss')) cannot be right"
    }

    if ($fileDate -lt $EarliestDate -or $fileDate -gt $LatestDate)
    {
        Write-DateWarning -Text "$name : the file system date $($fileDate.ToString('yyyy-MM-dd HH:mm:ss')) cannot be right"
    }

    if ($null -eq $Interval)
    {
        # No date in the name of the file and none in the name of any folder above it, so there is
        # nothing to measure against. The warning above is all that can be said about this one.
        return $null
    }

    $fileOff = Get-DaysOutside -Stamp $fileDate -Interval $Interval
    $exifOff = 0

    if ($null -ne $exifDate)
    {
        $exifOff = Get-DaysOutside -Stamp $exifDate -Interval $Interval
    }

    # Whichever of the two dates already sits inside the interval is the best one there is, so it is
    # kept as it stands and the other is brought to it. When neither does, the near end of the
    # interval is taken, with the time of day of the date that is being replaced.
    #
    # A metadata date that reads as a camera clock which lost its date is not one of the two, though. It
    # is not a moment anybody recorded, so the file system date is preferred to it even when that has to
    # be moved to an end of the interval, and its time of day is not kept below either.
    #
    # Only a date that is already too far outside the interval to be believed is asked about, the same
    # $FolderDayTolerance days that make it wrong at all. A photo taken in the small hours of the first
    # of January sits in the folder of the December before it, a few days outside it and right, and a
    # clock is not called broken for landing that close to where its folder says it should be.
    $lostClock = $false

    if ($null -ne $exifDate -and $exifOff -gt $FolderDayTolerance)
    {
        if (Test-LostClockDate -Stamp $exifDate -FileStamp $fileDate -Interval $Interval)
        {
            $lostClock = $true
            $script:LostClocks++
        }
    }

    $clamped = $false

    if ($null -ne $exifDate -and $exifOff -eq 0)
    {
        $target = $exifDate
    }
    elseif ($fileOff -eq 0)
    {
        $target = $fileDate
    }
    elseif ($null -ne $exifDate -and -not $lostClock -and $exifDate -ge $EarliestDate -and $exifDate -le $LatestDate)
    {
        $target  = Get-ClampedDate -Stamp $exifDate -Interval $Interval
        $clamped = $true
    }
    else
    {
        $target  = Get-ClampedDate -Stamp $fileDate -Interval $Interval
        $clamped = $true
    }

    # Both dates had to be thrown away and the day written comes from an end of the interval. When that
    # interval is a wide one the day is a guess, and the same guess for every file in the folder, so the
    # file is only put in the list and nothing is written for it. See $FolderMaxIntervalDays.
    $reportOnly = ($clamped -and ($Interval.Last - $Interval.First).TotalDays -gt $FolderMaxIntervalDays)

    # The name of a folder says which day and says nothing at all about the hour, so the hour a file
    # already carries is the best one there is and it is kept: only the day of it moves. A photo whose
    # metadata reads 2012:10:22 19:32:51 in a folder called 2011 is written 2011:01:01 19:32:51, that
    # 19:32:51 being the second the camera recorded, and not the time of day the file system date
    # happens to hold. The file system date keeps its own time of day for the same reason, it is what
    # $target already carries. So the two can differ in the time of day, each holding on to its own.
    #
    # A tag that is not there, one holding something that is not a date, and a clock that lost its date
    # have no time of day worth keeping and take the target whole. That last one is the counterpart of
    # the case above: GH010025.MP4 holding 2018:01:01 04:30:40 in a folder called 2019.06, next to a file
    # system date of 2019:06:01 14:54:52, is written 2019:06:01 14:54:52 and not 04:30:40, because the
    # camera did not know the time of day either, it was counting from its own reset.
    $metaTarget = $target

    if ($null -ne $exifDate -and -not $lostClock)
    {
        $metaTarget = $target.Date.Add($exifDate.TimeOfDay)

        # A time of day kept across midnight: the file was written at 23:04:26 and the camera called that
        # moment 00:03:33, one hour apart on two clocks and a day apart on the calendar. Keeping the hour
        # on the day of the file would then move the picture almost a whole day, so the day goes with the
        # hour, as long as the interval the folder stands for still holds the answer.
        if (($metaTarget - $target).TotalHours -gt 12 -and $metaTarget.AddDays(-1) -ge $Interval.First)
        {
            $metaTarget = $metaTarget.AddDays(-1)
        }
        elseif (($metaTarget - $target).TotalHours -lt -12 -and $metaTarget.AddDays(1) -le $Interval.Last)
        {
            $metaTarget = $metaTarget.AddDays(1)
        }
    }

    # A date tag holding a value that is not a date is repaired here as everywhere else. A tag that is
    # simply not there is left alone, whatever $IncludeFilesWithoutExifDate says: for a file with no
    # date in its name either, the name of a folder is too thin a reason to write metadata that was
    # never there. Such a file only has its file system date corrected.
    if ($Read.Bad)
    {
        $metaWrong = $true
    }
    elseif ($null -eq $exifDate)
    {
        $metaWrong = $false
    }
    else
    {
        $metaWrong = ($exifOff -gt $FolderDayTolerance)
    }

    $fileWrong = ($FixFileSystemDate -and $fileOff -gt $FolderDayTolerance)

    if (-not $metaWrong -and -not $fileWrong)
    {
        return $null
    }

    $reasons = New-Object System.Collections.Generic.List[string]

    if ($metaWrong)
    {
        if ($Read.Bad)
        {
            $reasons.Add("$($Read.BadTag) holds $($Read.BadValue), which is not a date")
        }
        else
        {
            $stamp = $exifDate.ToString("yyyy-MM-dd HH:mm:ss")
            $reasons.Add("the metadata date $stamp is $([int] $exifOff) day(s) outside the folder $($Interval.Text)")

            if ($lostClock)
            {
                $reasons.Add("it reads as a camera clock that lost its date, so its time of day is not kept either")
            }
        }
    }

    if ($fileWrong)
    {
        $stamp = $fileDate.ToString("yyyy-MM-dd HH:mm:ss")
        $reasons.Add("the file system date $stamp is $([int] $fileOff) day(s) outside the folder $($Interval.Text)")
    }

    if ($reportOnly)
    {
        $reasons.Add("nothing is written, the folder $($Interval.Text) is too wide to say which day this was")
    }

    return [pscustomobject] @{
        File        = $File.FullName
        Name        = $File.Name
        IsVideo     = $IsVideo
        Source      = "folder"
        DateFrom    = "folder " + $Interval.Text
        FolderText  = $Interval.Text
        NameDate    = $null
        NameHasTime = $false
        ExifDate    = $exifDate
        FileDate    = $fileDate
        Target      = $target
        MetaTarget  = $metaTarget
        LostClock   = $lostClock
        MetaWrong   = $metaWrong
        FileWrong   = $fileWrong
        ReportOnly  = $reportOnly
        Info        = $Info
        Reason      = ($reasons -join ", and ")
    }
}

function Get-ExifArguments
{
    param([bool] $UseFast, [bool] $ViaArgFile)

    $arguments = @("-j", "-q")

    if ($ViaArgFile)
    {
        # The argument file is written as UTF-8, this tells exiftool how to read the names in it.
        # On the command line the names arrive in the code page of the console instead, so the
        # option would make exiftool misread them.
        $arguments += @("-charset", "filename=utf8")
    }

    if ($UseFast)
    {
        $arguments += "-fast2"
    }

    if ($QuickTimeUtc)
    {
        $arguments += @("-api", "QuickTimeUTC=1")
    }

    return ($arguments + @("-DateTimeOriginal", "-CreateDate", "-ModifyDate", "-MediaCreateDate",
                           "-TrackCreateDate", "-FileModifyDate"))
}

function Invoke-ExifJson
{
    param([string] $Tool, [string[]] $Arguments, [string] $ErrPath)

    $json = & $Tool @Arguments 2>$ErrPath

    # One string, or Windows PowerShell tries to parse each line on its own and fails.
    $text = ($json -join [Environment]::NewLine).Trim()

    # Kept so a group that goes wrong can be described in the log.
    $script:LastExifArgs = $Arguments
    $script:LastExifRaw  = $text

    if ($text -eq "")
    {
        return @()
    }

    try
    {
        # Windows PowerShell hands over the whole JSON array as one object instead of one object
        # per record, so it is put in a variable first and spread out from there.
        $parsed = $text | ConvertFrom-Json
    }
    catch
    {
        return @()
    }

    if ($null -eq $parsed)
    {
        return @()
    }

    return @($parsed)
}

function Test-ArgFileWasRead
{
    param($Entries, [string] $ArgFile)

    # exiftool answering about the argument file itself means -@ was ignored and the file was
    # taken for a photo. That is one answer per call whatever the group holds.
    $name = [System.IO.Path]::GetFileName($ArgFile).ToLower()

    foreach ($entry in $Entries)
    {
        $property = $entry.PSObject.Properties["SourceFile"]

        if ($null -ne $property -and $property.Value -and ([string] $property.Value).ToLower().Contains($name))
        {
            return $true
        }
    }

    return $false
}

function Get-ExifMessage
{
    param([string] $Text)

    # PowerShell wraps every line a native command writes to its error stream in an error record, so
    # what comes back is the message of exiftool followed by the line of the script that called it,
    # the line underlined, and a couple of CategoryInfo lines. Only what exiftool said is of any use.
    if ([string]::IsNullOrWhiteSpace($Text))
    {
        return ""
    }

    $wanted = New-Object System.Collections.Generic.List[string]

    foreach ($line in ($Text -split '\r?\n'))
    {
        $clean = $line.Trim()

        if ($clean -eq "" -or $clean.StartsWith("+") -or $clean.StartsWith("At "))
        {
            continue
        }

        # "exiftool.exe : Error: ..." keeps only what follows the name of the tool.
        $wanted.Add(($clean -replace '^\S*exiftool(\.exe)?\s*:\s*', ''))
    }

    return (($wanted -join " ").Trim())
}

function Write-ExifDiagnostics
{
    param([object[]] $Files, $Entries, [string] $ErrPath, [string] $Tool, [bool] $UseFast)

    # Everything about one failing group, written once, so the cause can be seen instead of guessed.
    if ($script:UseArgFile)
    {
        $mode = "names in an argument file (-@)"
    }
    else
    {
        $mode = "names on the command line"
    }

    Write-Log ""
    Write-Log "--- exiftool did not answer for a whole group, here is what happened ---"
    Write-Log ("mode          : {0}" -f $mode)
    Write-Log ("fast reading  : {0}" -f $UseFast)
    Write-Log ("files asked   : {0}" -f $Files.Count)

    for ($i = 0; $i -lt [Math]::Min(3, $Files.Count); $i++)
    {
        Write-Log ("  file {0}      : {1}" -f $i, $Files[$i].FullName)
    }

    Write-Log ("arguments     : {0}" -f (Get-Shortened -Text ($script:LastExifArgs -join " ") -Length 1500))
    Write-Log ("raw answer    : {0}" -f (Get-Shortened -Text $script:LastExifRaw -Length 2000))

    $problem = (Get-Content -LiteralPath $ErrPath -Raw -ErrorAction SilentlyContinue)
    Write-Log ("error output  : {0}" -f (Get-Shortened -Text $problem -Length 1000))

    $entryList = @($Entries)
    Write-Log ("answers read  : {0}" -f $entryList.Count)

    for ($i = 0; $i -lt [Math]::Min(5, $entryList.Count); $i++)
    {
        $entry = $entryList[$i]
        $kind  = "null"
        $names = "none"

        if ($null -ne $entry)
        {
            $kind  = $entry.GetType().FullName
            $names = (@($entry.PSObject.Properties | ForEach-Object { $_.Name }) -join ", ")
        }

        $shortNames = Get-Shortened -Text $names -Length 500
        $shortValue = Get-Shortened -Text ([string] $entry) -Length 500

        Write-Log ("  answer {0}    : type {1}" -f $i, $kind)
        Write-Log ("  fields {0}    : {1}" -f $i, $shortNames)
        Write-Log ("  value  {0}    : {1}" -f $i, $shortValue)
    }

    # A control call with a single name on the command line, which is how the script read files before.
    if ($Files.Count -gt 0)
    {
        $single = @("-j", "-DateTimeOriginal", "-CreateDate", "-ModifyDate", "-FileModifyDate", $Files[0].FullName)
        $answer = ""

        try
        {
            $answer = ((& $Tool @single 2>$ErrPath) -join [Environment]::NewLine).Trim()
        }
        catch
        {
            $answer = "the call itself failed: $($_.Exception.Message)"
        }

        Write-Log ("control call  : one name on the command line, for {0}" -f $Files[0].Name)
        Write-Log ("control answer: {0}" -f (Get-Shortened -Text $answer -Length 1500))
    }

    Write-Log "--- end of the description ---"
    Write-Log ""
}

function Write-ExifTrouble
{
    param([int] $Answered, [int] $Wanted, [int] $Missing, [string] $ErrPath, [object[]] $Files, $Entries, [string] $Tool, [bool] $UseFast)

    $script:ExifTroubleGroups++

    $line = "  exiftool answered for {0} of {1} file(s) in this group, {2} could not be matched." -f $Answered, $Wanted, $Missing
    Write-Log $line

    if (-not $script:ExifDiagDone)
    {
        $script:ExifDiagDone = $true
        Write-ExifDiagnostics -Files $Files -Entries $Entries -ErrPath $ErrPath -Tool $Tool -UseFast $UseFast
        Write-Host "  The first group that went wrong is described in full in the log file." -ForegroundColor Yellow
    }

    if ($script:ExifTroubleGroups -gt 3)
    {
        return
    }

    Write-Host $line -ForegroundColor Yellow

    $problem = Get-ExifMessage -Text (Get-Content -LiteralPath $ErrPath -Raw -ErrorAction SilentlyContinue)

    if ($problem -ne "")
    {
        Write-Host ("  exiftool said: {0}" -f $problem) -ForegroundColor Yellow
    }

    if ($script:ExifTroubleGroups -eq 3)
    {
        Write-Host "  more groups like this are only counted, the total comes at the end of the scan." -ForegroundColor Yellow
    }
}

function Read-ExifByCommandLine
{
    param([string] $Tool, [object[]] $Files, [string[]] $Base, [string] $ErrPath)

    # Windows takes about 32000 characters of command line, so the group is cut in slices that
    # stay well under that. exiftool keeps the order given, so the answers stay lined up.
    $limit   = 24000
    $entries = New-Object System.Collections.Generic.List[object]
    $slice   = New-Object System.Collections.Generic.List[string]
    $length  = 0

    foreach ($file in $Files)
    {
        $room = $file.FullName.Length + 3

        if ($slice.Count -gt 0 -and ($length + $room) -gt $limit)
        {
            $entries.AddRange([object[]] @(Invoke-ExifJson -Tool $Tool -Arguments ($Base + $slice.ToArray()) -ErrPath $ErrPath))
            $slice.Clear()
            $length = 0
        }

        $slice.Add($file.FullName)
        $length += $room
    }

    if ($slice.Count -gt 0)
    {
        $entries.AddRange([object[]] @(Invoke-ExifJson -Tool $Tool -Arguments ($Base + $slice.ToArray()) -ErrPath $ErrPath))
    }

    return $entries.ToArray()
}

function Get-ExifInfoBatch
{
    param([string] $Tool, [object[]] $Files, [bool] $UseFast)

    # Keyed by the position in $Files, not by path text. exiftool prints SourceFile with forward
    # slashes and in the console code page, so comparing paths breaks on anything non ASCII.
    $map = @{}

    if ($null -eq $Files -or $Files.Count -eq 0)
    {
        return $map
    }

    $errFile = [System.IO.Path]::GetTempFileName()
    $argFile = $null

    try
    {
        $base    = Get-ExifArguments -UseFast $UseFast -ViaArgFile $script:UseArgFile
        $entries = @()

        if ($script:UseArgFile)
        {
            # One call takes any number of files this way, the command line length does not matter.
            $argFile = [System.IO.Path]::GetTempFileName()
            $paths   = New-Object System.Collections.Generic.List[string]

            foreach ($file in $Files)
            {
                $paths.Add($file.FullName)
            }

            [System.IO.File]::WriteAllLines($argFile, $paths, (New-Object System.Text.UTF8Encoding($false)))

            $entries = @(Invoke-ExifJson -Tool $Tool -Arguments ($base + @("-@", $argFile)) -ErrPath $errFile)

            if ($entries.Count -ne $Files.Count -and (Test-ArgFileWasRead -Entries $entries -ArgFile $argFile))
            {
                $script:UseArgFile = $false

                Write-Host "  This exiftool ignores -@ and reads the argument file itself, so the names go on the command line from now on." -ForegroundColor Yellow

                return (Get-ExifInfoBatch -Tool $Tool -Files $Files -UseFast $UseFast)
            }
        }
        else
        {
            $entries = @(Read-ExifByCommandLine -Tool $Tool -Files $Files -Base $base -ErrPath $errFile)
        }

        if ($entries.Count -eq $Files.Count)
        {
            # The normal case, exiftool answers once per file and keeps the order it was given.
            for ($i = 0; $i -lt $Files.Count; $i++)
            {
                $map[$i] = $entries[$i]
            }

            return $map
        }

        # Something was skipped, so fall back to matching on the file name only.
        $byName = @{}

        foreach ($entry in $entries)
        {
            $property = $entry.PSObject.Properties["SourceFile"]

            if ($null -ne $property -and $property.Value)
            {
                $name = [System.IO.Path]::GetFileName((([string] $property.Value) -replace '/', '\')).ToLower()
                $byName[$name] = $entry
            }
        }

        $matched = 0

        for ($i = 0; $i -lt $Files.Count; $i++)
        {
            $name = $Files[$i].Name.ToLower()

            if ($byName.ContainsKey($name))
            {
                $map[$i] = $byName[$name]
                $matched++
            }
        }

        Write-ExifTrouble -Answered $entries.Count -Wanted $Files.Count -Missing ($Files.Count - $matched) -ErrPath $errFile -Files $Files -Entries $entries -Tool $Tool -UseFast $UseFast
    }
    finally
    {
        if ($null -ne $argFile)
        {
            Remove-Item -LiteralPath $argFile -Force -ErrorAction SilentlyContinue
        }

        Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    }

    return $map
}

function Get-ExifDateInfo
{
    param($Info, [bool] $IsVideo)

    # Date is the first date tag that holds a real date. Bad is set when a date tag is there but
    # holds something that is not a date, 0000:00:00 00:00:00 for example, which is a value that
    # has to be corrected, and BadTag and BadValue then say which tag it was and what stood in it.
    # A file with no date tag at all leaves all of these alone: there is nothing wrong in it, there
    # is simply nothing in it, and that is not the same thing.
    $result = [pscustomobject] @{ Date = $null; Bad = $false; BadTag = ""; BadValue = "" }

    if ($null -eq $Info)
    {
        return $result
    }

    if ($IsVideo)
    {
        $order = @("CreateDate", "MediaCreateDate", "TrackCreateDate", "DateTimeOriginal", "ModifyDate")
    }
    else
    {
        $order = @("DateTimeOriginal", "CreateDate", "ModifyDate")
    }

    foreach ($tag in $order)
    {
        $raw = $Info.PSObject.Properties[$tag]
        if ($null -eq $raw -or -not $raw.Value)
        {
            continue
        }

        $text = ($raw.Value -replace '(\+|\-)\d{2}:\d{2}$', '').Trim()
        $text = ($text -replace 'Z$', '').Trim()

        if ($text -eq "")
        {
            continue
        }

        $parsed = [datetime]::MinValue

        if ([datetime]::TryParseExact($text, "yyyy:MM:dd HH:mm:ss", $null, [System.Globalization.DateTimeStyles]::None, [ref] $parsed))
        {
            $result.Date = $parsed
            return $result
        }

        # The tag is written and what stands in it is not a date. 0000:00:00 00:00:00 is the usual
        # one, phones write it into videos. A tag holding nothing at all was passed over above.
        # The first such value is kept so the list and the screen can name it.
        if (-not $result.Bad)
        {
            $result.BadTag   = $tag
            $result.BadValue = $text
        }

        $result.Bad = $true
    }

    return $result
}

function Show-ExifInfo
{
    param($Info, [string] $Label)

    Write-Host "  $Label" -ForegroundColor DarkGray

    if ($null -eq $Info)
    {
        Write-Host "    (no metadata could be read)" -ForegroundColor DarkGray
        return
    }

    $tags  = @("DateTimeOriginal", "CreateDate", "ModifyDate", "MediaCreateDate", "TrackCreateDate", "FileModifyDate")
    $shown = $false

    foreach ($tag in $tags)
    {
        $property = $Info.PSObject.Properties[$tag]
        if ($null -ne $property -and $property.Value)
        {
            Write-Host ("    {0,-18} {1}" -f $tag, $property.Value)
            $shown = $true
        }
    }

    if (-not $shown)
    {
        Write-Host "    (no date tags present)" -ForegroundColor DarkGray
    }
}

function Get-CsvLine
{
    param([object[]] $Values)

    # Every field is quoted and a quote inside a field is doubled, which is what Excel expects.
    # That way a path holding the separator, a quote or a comma cannot break the columns.
    $fields = New-Object System.Collections.Generic.List[string]

    foreach ($value in $Values)
    {
        $text = ""

        if ($null -ne $value)
        {
            $text = [string] $value
        }

        $fields.Add('"' + ($text -replace '"', '""') + '"')
    }

    return ($fields -join $CsvDelimiter)
}

function Get-CandidateHeader
{
    # DateFrom says where the new date was taken from, the name of the file or the name of a folder,
    # and DateInName is filled in only for the first kind.
    $names = @("File", "Folder", "Kind", "MetadataDate", "FileSystemDate", "DateFrom", "DateInName",
               "TimeInName", "Writes", "Reason")

    return (Get-CsvLine -Values $names)
}

function Get-FixDescription
{
    param($Candidate)

    # What is written and what is written into it, in one column: "metadata : 2024-04-28 15:08:42".
    # The two can hold the same day at a different time of day, each keeping the one it already had,
    # and then both are named.
    $stamp = $Candidate.Target.ToString("yyyy-MM-dd HH:mm:ss")
    $meta  = $Candidate.MetaTarget.ToString("yyyy-MM-dd HH:mm:ss")

    if ($Candidate.ReportOnly)
    {
        return "nothing"
    }

    if ($Candidate.MetaWrong -and $Candidate.FileWrong)
    {
        if ($meta -eq $stamp)
        {
            return "metadata and file date : $stamp"
        }

        return "metadata : $meta, file date : $stamp"
    }

    if ($Candidate.MetaWrong)
    {
        return "metadata : $meta"
    }

    return "file date : $stamp"
}

function Get-CandidateLine
{
    param($Candidate)

    if ($null -eq $Candidate.ExifDate)
    {
        $current = ""
    }
    else
    {
        $current = $Candidate.ExifDate.ToString("yyyy-MM-dd HH:mm:ss")
    }

    if ($Candidate.IsVideo)
    {
        $kind = "video"
    }
    else
    {
        $kind = "photo"
    }

    # A file whose date comes from the name of its folder has no date in its own name, so the two
    # columns about that name stay empty for it rather than repeating the folder date.
    if ($null -eq $Candidate.NameDate)
    {
        $inName     = ""
        $timeInName = ""
    }
    else
    {
        $inName = $Candidate.NameDate.ToString("yyyy-MM-dd HH:mm:ss")

        if ($Candidate.NameHasTime)
        {
            $timeInName = "yes"
        }
        else
        {
            $timeInName = "day only"
        }
    }

    $values = @($Candidate.Name,
                [System.IO.Path]::GetDirectoryName($Candidate.File),
                $kind,
                $current,
                $Candidate.FileDate.ToString("yyyy-MM-dd HH:mm:ss"),
                $Candidate.DateFrom,
                $inName,
                $timeInName,
                (Get-FixDescription -Candidate $Candidate),
                $Candidate.Reason)

    return (Get-CsvLine -Values $values)
}

function Set-ExifDate
{
    param([string] $Tool, [string] $File, [datetime] $MetaStamp, [datetime] $FileStamp, [bool] $IsVideo,
          [bool] $WriteMetadata, [bool] $WriteFileTimes)

    # Two values, because a date taken from the name of a folder moves the day and leaves the time of
    # day where it was, so the metadata and the file system date can end up on the same day at the hour
    # each of them already held. For every other file the two are the same date.
    $metaValue = $MetaStamp.ToString("yyyy:MM:dd HH:mm:ss")
    $fileValue = $FileStamp.ToString("yyyy:MM:dd HH:mm:ss")

    # -m makes exiftool write the file anyway when it finds something it calls a minor error, a video
    # with a few bytes of garbage after its last atom for instance. Without it such a file is refused
    # and can never be corrected. Take -m out to have those files reported as failed instead.
    # -q drops the "1 image files updated" line that would otherwise be reported for every file.
    # Warnings and errors are not affected by it, they are what the caller wants to hear about.
    $arguments = @("-overwrite_original", "-m", "-q")

    if (-not $WriteFileTimes)
    {
        # Writing metadata changes the file, which would move its file system date. -P holds it back.
        $arguments += "-P"
    }

    if ($WriteMetadata)
    {
        $arguments += "-AllDates=$metaValue"

        if ($IsVideo)
        {
            if ($QuickTimeUtc)
            {
                $arguments += @("-api", "QuickTimeUTC=1")
            }

            $arguments += @("-QuickTime:CreateDate=$metaValue", "-QuickTime:ModifyDate=$metaValue",
                            "-TrackCreateDate=$metaValue", "-TrackModifyDate=$metaValue",
                            "-MediaCreateDate=$metaValue", "-MediaModifyDate=$metaValue")
        }
    }

    if ($WriteFileTimes)
    {
        # These two touch nothing inside the file, so a file that only needs its dates straightened
        # out is never rewritten.
        $arguments += @("-FileModifyDate=$fileValue", "-FileCreateDate=$fileValue")
    }

    $arguments += @("--", $File)

    # What exiftool says about the file goes to a temporary file, it must not be merged into the
    # output with 2>&1. Merged, every line it writes there arrives as an error record, and with
    # $ErrorActionPreference = "Stop" the run dies on the first warning about a file, which is
    # exactly what "Possible garbage at end of file" did.
    $errFile = [System.IO.Path]::GetTempFileName()
    $argFile = $null

    # Only inside this function, so that nothing exiftool prints can end the run. The setting at the
    # top of the script still holds everywhere else.
    $ErrorActionPreference = "Continue"

    try
    {
        if ($script:UseArgFile)
        {
            # The whole command goes to exiftool in a UTF-8 argument file, one argument per line, the
            # same way the reading side hands over its names. On the command line a name is passed in
            # the code page of the console, and a letter that has no place there, the t with a comma
            # of "Proprietati" for example, arrives as a question mark. exiftool then takes that for
            # a wildcard and answers "Wildcards don't work in the directory specification" instead of
            # writing the file. -charset stays on the command line so it is read before the names.
            $argFile = [System.IO.Path]::GetTempFileName()
            [System.IO.File]::WriteAllLines($argFile, $arguments, (New-Object System.Text.UTF8Encoding($false)))

            $output = & $Tool "-charset" "filename=utf8" "-@" $argFile 2>$errFile
        }
        else
        {
            $output = & $Tool @arguments 2>$errFile
        }

        $code    = $LASTEXITCODE
        $problem = (Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue)
        $said    = Get-ExifMessage -Text $problem

        if ($said -eq "")
        {
            $said = Get-ExifMessage -Text ($output -join [Environment]::NewLine)
        }

        # A warning comes with a zero exit code, the file is written all the same. Only the exit
        # code says whether it worked, the text is passed on either way.
        return @{ Success = ($code -eq 0); Output = $said }
    }
    finally
    {
        if ($null -ne $argFile)
        {
            Remove-Item -LiteralPath $argFile -Force -ErrorAction SilentlyContinue
        }

        Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Phase 0: ask what to do
# ---------------------------------------------------------------------------

$toolVersion = Test-ExifTool -Tool $ExifTool

Write-Host ""
Write-Host "Fix-PhotoDates" -ForegroundColor Cyan
Write-Host "  exiftool : $toolVersion"
Write-Host ""

$scanPath = $DefaultFolder

while ($true)
{
    $scanPath = Read-Text -Prompt "Folder to scan" -Default $scanPath

    if (Test-Path -LiteralPath $scanPath -PathType Container)
    {
        break
    }

    Write-Host "  Not a folder: $scanPath" -ForegroundColor Red
}

Write-Host ""
Write-Host "What do you want to do?" -ForegroundColor Cyan
Write-Host "  1  Show the files that need a fix and write the list to a file, change nothing"
Write-Host "  2  Show them, then ask me file by file"
Write-Host "  3  Show them, then fix all of them without asking"
Write-Host "  4  Just write the list to the file, do not show it here"
Write-Host "  q  Quit"
Write-Host ""

$mode = Read-Choice -Prompt "Choice" -Valid @("1", "2", "3", "4", "q") -Default "2"

if ($mode -eq "q")
{
    Write-Host "Nothing done." -ForegroundColor Yellow
    return
}

$listOnly = ($mode -eq "1" -or $mode -eq "4")
$fixAll   = ($mode -eq "3")
$showList = ($mode -ne "4")

Write-Host ""
$reportPath = Read-Text -Prompt "Write the list to (a csv for Excel)" -Default $DefaultOutput

Write-Host ""
Write-Host "Running with:" -ForegroundColor Cyan
Write-Host "  folder      : $scanPath (including subfolders)"

if ($listOnly)
{
    Write-Host "  action      : list only, no file is changed"
}
elseif ($fixAll)
{
    Write-Host "  action      : fix every listed file without asking" -ForegroundColor Yellow
}
else
{
    Write-Host "  action      : ask before each file"
}

Write-Host "  list file   : $reportPath"
Write-Host "  order       : newest first"

if ($FixFileSystemDate)
{
    Write-Host "  file dates  : brought to the date from the name when they differ"
}
else
{
    Write-Host "  file dates  : left alone, only metadata is written"
}

if ($CheckFilesWithoutDateInName)
{
    Write-Host "  no date name: read as well, judged by the date in the name of their folder"
}
else
{
    Write-Host "  no date name: files with no date in the name are not read at all"
}

# ---------------------------------------------------------------------------
# Phase 1: scan
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Looking for media files..." -ForegroundColor Cyan

# Name descending puts date named files roughly newest first, which is the order they are read in.
$files = @(Get-ChildItem -LiteralPath $scanPath -File -Recurse |
           Where-Object { $Extensions -contains $_.Extension.ToLower() } |
           Sort-Object -Property Name -Descending)

$total = $files.Count
Write-Host "$total media file(s) found." -ForegroundColor Cyan

# The walk up the folders must not go above the folder being scanned, so its full path is needed.
$scanFull = (Resolve-Path -LiteralPath $scanPath).ProviderPath.TrimEnd('\')

# Files with a date in the name are measured against it. The others are read as well, for the warning
# about impossible dates and for the date in the name of their folder, unless the setting says not to.
# What the folder says is not worked out here but during the scan, because reading a folder name can
# raise a warning about the naming convention and the log file is only opened further down.
$work     = New-Object System.Collections.Generic.List[object]
$withName = 0

foreach ($file in $files)
{
    $nameDate = Get-DateFromName -Name $file.Name

    if ($null -ne $nameDate)
    {
        $withName++
        $work.Add([pscustomobject] @{ File = $file; NameDate = $nameDate.Date; NameHasTime = $nameDate.HasTime })
        continue
    }

    if (-not $CheckFilesWithoutDateInName)
    {
        continue
    }

    $work.Add([pscustomobject] @{ File = $file; NameDate = $null; NameHasTime = $false })
}

# How many of the files with no date in the name sit in a folder whose name gives no date either, and
# so can only be reported on, never fixed. Counted while scanning, since that is where it is found out.
$noFolderDate = 0

$toRead      = $work.Count
$withoutName = $total - $withName
$skipped     = $total - $toRead

Write-Host "$withName have a date in the name, $withoutName do not." -ForegroundColor Cyan

if ($CheckFilesWithoutDateInName)
{
    Write-Host "The ones without are read too, for the date in the name of their folder." -ForegroundColor Cyan
}
else
{
    Write-Host "The ones without are ignored, the check for them is switched off." -ForegroundColor Cyan
}

if ($script:ImplausibleNames -gt 0)
{
    # Their digits read as a date no photo can carry, a serial number or a resolution, so those names
    # count as holding no date at all and the files go with the ones judged by their folder.
    $limit = $EarliestDate.ToString("yyyy-MM-dd")
    Write-Host "$script:ImplausibleNames of those held a date outside $limit to today, a serial number read as one." -ForegroundColor Yellow
}

Write-Host "A file is listed when its file system date is more than $ConflictSeconds second(s) off the date from the name," -ForegroundColor Cyan
Write-Host "or when its metadata date disagrees with both the file system date and the date from the name," -ForegroundColor Cyan
Write-Host "or when a date tag holds a value that is not a date, 0000:00:00 00:00:00 for example." -ForegroundColor Cyan

if ($TimeZoneHoursMax -gt 0)
{
    Write-Host "A difference of a whole number of hours, up to $TimeZoneHoursMax, counts as a time zone and is left alone," -ForegroundColor Cyan
    Write-Host "both between the file system date and the date the file should have, and between the name and the metadata." -ForegroundColor Cyan
}

if ($CheckFilesWithoutDateInName)
{
    Write-Host "A file with no date in its name is listed when its dates are more than $FolderDayTolerance day(s)" -ForegroundColor Cyan
    Write-Host "outside the date its folder is named after, and a date that cannot be right is reported as it is met." -ForegroundColor Cyan
    Write-Host "A folder name spells a date with dots and an interval with a dash: 2016.11.05, 2024.02, 2016.10-12," -ForegroundColor Cyan
    Write-Host "2016.11.05-06, 2016.11.28-12.03, 2015-2016. A name that breaks that is reported and still read." -ForegroundColor Cyan
    Write-Host "When both dates of a file are outside a folder wider than $FolderMaxIntervalDays day(s), the day could only" -ForegroundColor Cyan
    Write-Host "come from an end of it, which is a guess, so the file is only listed and nothing is written for it." -ForegroundColor Cyan
    Write-Host "A metadata date in a January of another year, more than $LostClockHours hour(s) from the time of day of" -ForegroundColor Cyan
    Write-Host "the file, is a clock counting from its own reset, so the file system date is taken whole." -ForegroundColor Cyan
}

if (-not [System.IO.Path]::IsPathRooted($reportPath))
{
    $reportPath = Join-Path (Get-Location).Path $reportPath
}

Write-Host "Writing the list to $reportPath as files are found." -ForegroundColor Cyan

# The log holds what the console shows plus the details of anything that goes wrong with exiftool.
$logPath          = [System.IO.Path]::ChangeExtension($reportPath, ".log")
$script:LogWriter = New-Object System.IO.StreamWriter($logPath, $false, (New-Object System.Text.UTF8Encoding($false)))
$script:LogWriter.AutoFlush = $true

Write-Log "Fix-PhotoDates started"
Write-Log "exiftool  : $toolVersion at $ExifTool"
Write-Log "host      : PowerShell $($PSVersionTable.PSVersion)"
Write-Log "folder    : $scanPath (including subfolders)"
Write-Log "list file : $reportPath"
Write-Log "mode      : $mode"
Write-Log "group size: $ChunkSize"
Write-Log "rule      : file system date more than $ConflictSeconds second(s) off the date from the name, or metadata date off both"
Write-Log "rule      : a date tag holding no real date, 0000:00:00 00:00:00 for example, is always written"
Write-Log "rule      : a difference of a whole number of hours, up to $TimeZoneHoursMax, is a time zone and is left alone"
Write-Log "rule      : that holds for the file system date and for the time in the name against the metadata"
Write-Log "file dates: $(if ($FixFileSystemDate) { 'brought to the date from the name' } else { 'left alone' })"
Write-Log "no tag    : metadata of files with no date tag at all is $(if ($IncludeFilesWithoutExifDate) { 'written' } else { 'left alone' })"
Write-Log "range     : a date from a name is believed between $($EarliestDate.ToString('yyyy-MM-dd')) and $($LatestDate.ToString('yyyy-MM-dd'))"
Write-Log "no name   : files with no date in the name are $(if ($CheckFilesWithoutDateInName) { 'read as well' } else { 'ignored' })"
Write-Log "rule      : for those, a date more than $FolderDayTolerance day(s) outside the date their folder is named after"
Write-Log "rule      : for those, empty metadata is never written, only the file system date is corrected"
Write-Log "rule      : a folder name spells a date with dots and an interval with a dash, 2016.10-12 for instance"
Write-Log "rule      : a folder name that breaks that gets a warning, the date in it is read all the same"
Write-Log "rule      : when both dates of a file are outside a folder wider than $FolderMaxIntervalDays day(s) it is listed, not written"
Write-Log "rule      : a metadata date in the first $ResetClockDays day(s) of a January of another year, more than $LostClockHours hour(s)"
Write-Log "rule      : from the time of day of the file, is a clock counting from its own reset, so the file date is taken whole"
Write-Log "found     : $total media file(s), $withName with a date in the name, $withoutName without"
Write-Log "to read   : $toRead file(s)"
Write-Log "dropped   : $script:ImplausibleNames name(s) held a date outside that range"

Write-Host "Writing a log to $logPath" -ForegroundColor Cyan
Write-Host "Press Esc or Q to stop the scan and work with what was found so far." -ForegroundColor DarkGray
Write-Host ""

# The list is written while scanning, so it can be watched with Get-Content -Wait.
# When the scan ends it is rewritten in date order, with a separate information section after the
# rows to fix. The counts and everything about the run go to the log file next to it.
# The byte order mark is what makes Excel read the names with diacritics as UTF-8.
$writer = New-Object System.IO.StreamWriter($reportPath, $false, (New-Object System.Text.UTF8Encoding($true)))
$writer.AutoFlush = $true
$writer.WriteLine((Get-CandidateHeader))

$candidates = New-Object System.Collections.Generic.List[object]
$largeFileDateDifferences = New-Object System.Collections.Generic.List[object]
$index      = 0
$nextReport = 5
$stopped    = $false
$nothingToDo = 0
$noExifDate  = 0
$badExifDate = 0
$needMeta    = 0
$needFile    = 0
$watch      = [System.Diagnostics.Stopwatch]::StartNew()

while ($index -lt $toRead)
{
    $take  = [Math]::Min($ChunkSize, $toRead - $index)
    $chunk = @($work.GetRange($index, $take))

    # -fast2 is a big win but it skips metadata that video containers can keep out of the header,
    # so a chunk holding any video is read the normal way.
    $chunkHasVideo = $false

    foreach ($item in $chunk)
    {
        if ($VideoExtensions -contains $item.File.Extension.ToLower())
        {
            $chunkHasVideo = $true
            break
        }
    }

    $map = Get-ExifInfoBatch -Tool $ExifTool -Files @($chunk | ForEach-Object { $_.File }) -UseFast (-not $chunkHasVideo)

    for ($slot = 0; $slot -lt $chunk.Count; $slot++)
    {
        $item    = $chunk[$slot]
        $file    = $item.File
        $isVideo = $VideoExtensions -contains $file.Extension.ToLower()
        $info    = $null

        if ($map.ContainsKey($slot))
        {
            $info = $map[$slot]
        }

        $read     = Get-ExifDateInfo -Info $info -IsVideo $isVideo
        $exifDate = $read.Date
        $exifBad  = $read.Bad
        $fileDate = $file.LastWriteTime

        if ($null -eq $item.NameDate)
        {
            # No date in the name, so this file is judged by the name of its folder instead, and by
            # nothing at all when that holds no date either. It is a different question from the one
            # below and it is answered on its own, warnings included. The folder is read here rather
            # than before the scan so that a warning about its name reaches the log as well.
            $interval = Get-FolderDateInterval -Folder $file.DirectoryName -Root $scanFull

            if ($null -eq $interval)
            {
                $noFolderDate++
            }
            elseif (($differenceDays = Get-DaysOutside -Stamp $fileDate -Interval $interval) -gt $FolderDayTolerance)
            {
                $largeFileDateDifferences.Add([pscustomobject] @{
                    File          = $file.FullName
                    Name          = $file.Name
                    Folder        = $file.DirectoryName
                    FileDate      = $fileDate
                    FolderText    = $interval.Text
                    DaysOutside   = [int] [Math]::Ceiling($differenceDays)
                })
            }

            $candidate = New-FolderCandidate -File $file -Interval $interval -Info $info -IsVideo $isVideo -Read $read

            if ($null -eq $candidate)
            {
                if ($null -ne $interval)
                {
                    # It sits where its folder says it should, so as far as can be told it is right.
                    # A file with no folder date is not counted here, nothing was measured about it.
                    $nothingToDo++
                }

                continue
            }

            if ($candidate.ReportOnly)
            {
                # In the list for the record only, so it is not counted among the files to be written.
                $script:FolderReported++
            }
            else
            {
                if ($candidate.MetaWrong)
                {
                    $needMeta++
                }

                if ($candidate.FileWrong)
                {
                    $needFile++
                }
            }

            $script:FolderFixes++
            $candidates.Add($candidate)
            $writer.WriteLine((Get-CandidateLine -Candidate $candidate))
            continue
        }

        # The date the file should end up with. The name decides the day, and the time of day as well
        # when it carries one. A name that holds only a day has to take the time from somewhere else:
        # from the metadata when that already sits on the same day, and from the file system date
        # otherwise, since that is the only clock time left. Midnight is never forced on a file.
        if ($item.NameHasTime)
        {
            $nameGap = 0

            if ($null -ne $exifDate)
            {
                $nameGap = ($exifDate - $item.NameDate).TotalSeconds
            }

            if ($null -ne $exifDate -and [Math]::Abs($nameGap) -le $NameTimeSlackSeconds)
            {
                # The two tell the same time bar a rounding, so the metadata stays as it is and only
                # the file system date is brought to it.
                $target = $exifDate
            }
            elseif ($null -ne $exifDate -and (Test-TimeZoneShift -Seconds $nameGap))
            {
                # A whole number of hours between the name and the metadata is the same moment written
                # in two time zones, the phone naming the file in one and storing the other, and there
                # is no telling which of the two is the right one. The metadata keeps its own reading
                # rather than being moved to a zone just as arbitrary.
                $script:TimeZoneKept++
                $target = $exifDate
            }
            else
            {
                $target = $item.NameDate
            }
        }
        elseif ($null -ne $exifDate -and $exifDate.Date -eq $item.NameDate.Date)
        {
            $target = $exifDate
        }
        else
        {
            $target = $item.NameDate.Add($fileDate.TimeOfDay)
        }

        # The metadata is only called wrong when it disagrees with the file system date as well as
        # with the target. Metadata and file system date that say the same thing are two witnesses
        # for that date, and writing the target over a date that already says it changes nothing.
        # A date tag holding a value that is not a date is wrong whatever the rest says, that is a
        # field to be repaired. A file with no date tag at all is not wrong, it is only bare, and
        # nothing is written into it unless the setting at the top says so.
        if ($exifBad)
        {
            $badExifDate++
            $metaWrong = $true
        }
        elseif ($null -eq $exifDate)
        {
            $noExifDate++
            $metaWrong = $IncludeFilesWithoutExifDate
        }
        elseif ([Math]::Abs(($exifDate - $fileDate).TotalSeconds) -gt $ConflictSeconds -and
                [Math]::Abs(($exifDate - $target).TotalSeconds) -gt $ConflictSeconds)
        {
            $metaWrong = $true
        }
        else
        {
            $metaWrong = $false
        }

        # The file system date has no such doubt about it, the name decides and it follows. The one
        # exception is a gap of a whole number of hours, which is a time zone and not a wrong date, and
        # rewriting it would only move the file to another zone just as arbitrary as the one it is in.
        $fileGap   = ($fileDate - $target).TotalSeconds
        $fileWrong = $FixFileSystemDate -and [Math]::Abs($fileGap) -gt $ConflictSeconds

        if ($fileWrong -and (Test-TimeZoneShift -Seconds $fileGap))
        {
            $script:TimeZoneLeft++
            $fileWrong = $false
        }

        if (-not $metaWrong -and -not $fileWrong)
        {
            $nothingToDo++
            continue
        }

        $reasons = New-Object System.Collections.Generic.List[string]

        if ($metaWrong)
        {
            $needMeta++

            # The reason names the value that is being complained about, so a row in the list can be
            # judged without opening the file again.
            if ($exifBad)
            {
                $reasons.Add("$($read.BadTag) holds $($read.BadValue), which is not a date")
            }
            elseif ($null -eq $exifDate)
            {
                $reasons.Add("no date in the metadata")
            }
            else
            {
                $reasons.Add("the metadata date $($exifDate.ToString('yyyy-MM-dd HH:mm:ss')) is wrong")
            }
        }

        if ($fileWrong)
        {
            $needFile++
            # Not "the date from the name": when the name and the metadata agree bar a rounding it is
            # the metadata second that is written, so the reason has to say the date, not where from.
            $reasons.Add("the file system date is not the date the file was taken")
        }

        $candidate = [pscustomobject] @{
            File        = $file.FullName
            Name        = $file.Name
            IsVideo     = $isVideo
            Source      = "name"
            DateFrom    = "name"
            FolderText  = ""
            NameDate    = $item.NameDate
            NameHasTime = $item.NameHasTime
            ExifDate    = $exifDate
            FileDate    = $fileDate
            Target      = $target
            MetaTarget  = $target
            LostClock   = $false
            MetaWrong   = $metaWrong
            FileWrong   = $fileWrong
            ReportOnly  = $false
            Info        = $info
            Reason      = ($reasons -join ", and ")
        }

        $candidates.Add($candidate)
        $writer.WriteLine((Get-CandidateLine -Candidate $candidate))
    }

    $index  += $take
    $percent = [int](($index / [Math]::Max($toRead, 1)) * 100)

    # Time per file so far, used for the estimate of what is left.
    $remaining = [int](($watch.Elapsed.TotalSeconds / $index) * ($toRead - $index))
    $forFix    = $candidates.Count - $script:FolderReported
    $status    = "{0,3}%   {1} of {2}   {3} to fix" -f $percent, $index, $toRead, $forFix

    Write-Progress -Activity "Reading metadata" -Status $status -PercentComplete $percent -SecondsRemaining $remaining

    # A line in the console every 5 percent, so the progress is visible in a log as well.
    if ($percent -ge $nextReport -or $index -eq $toRead)
    {
        $left = ""

        if ($index -lt $toRead)
        {
            $left = ", about {0} left" -f (Format-Duration -Seconds $remaining)
        }

        Write-Host ("  {0,3}%   {1} of {2} read, {3} to fix so far{4}" -f $percent, $index, $toRead, $forFix, $left) -ForegroundColor DarkGray

        while ($nextReport -le $percent)
        {
            $nextReport += 5
        }
    }

    if (Test-StopKey)
    {
        $stopped = $true
        break
    }
}

$watch.Stop()
Write-Progress -Activity "Reading metadata" -Completed

if ($stopped)
{
    # Only the log says so, a line like this in the csv would land in the middle of the columns.
    Write-Host ""
    Write-Both -Text "Scan stopped after $index of $toRead file(s)." -Color Yellow
}

$writer.Close()
$writer.Dispose()

# Report and process by the date the file is going to end up with, newest first. That is the date from
# the name for most of them and the date from the folder for the rest, so the two kinds sort together.
$candidates = @($candidates | Sort-Object -Property @{ Expression = "Target"; Descending = $true }, @{ Expression = "Name"; Descending = $true })

# The rows that are only in the list for the record are not offered and not counted among the fixes.
$toFix = $candidates.Count - $script:FolderReported

Write-Host ""
Write-Both -Text "Read $index of $toRead file(s) in $(Format-Duration -Seconds $watch.Elapsed.TotalSeconds), $skipped were not read." -Color Cyan
Write-Both -Text "$nothingToDo file(s) are already right." -Color Cyan
Write-Both -Text "Of those with a date in the name, $noExifDate had no date tag and $badExifDate had a tag with no real date." -Color Cyan
Write-Both -Text "$needFile file(s) need their file system date corrected, $needMeta need their metadata written." -Color Cyan

if ($script:TimeZoneLeft -gt 0 -or $script:TimeZoneKept -gt 0)
{
    $shown = "$script:TimeZoneLeft file system date(s) and $script:TimeZoneKept metadata date(s) were a whole number"
    Write-Both -Text $shown -Color Cyan
    Write-Both -Text "of hours off, taken for a time zone and left as they are." -Color Cyan
}

if ($CheckFilesWithoutDateInName)
{
    $shown = "$script:FolderFixes listed file(s) have no date in their name and are judged by their folder."
    Write-Both -Text $shown -Color Cyan

    if ($noFolderDate -gt 0)
    {
        $shown = "$noFolderDate file(s) had no date in the name and none in the folder either, only read for the warnings."
        Write-Both -Text $shown -Color Cyan
    }

    if ($script:FolderReported -gt 0)
    {
        $shown = "$script:FolderReported of them sit in a folder too wide to say which day it was, listed but not written."
        Write-Both -Text $shown -Color Cyan
    }

    if ($script:LostClocks -gt 0)
    {
        $shown = "$script:LostClocks had a metadata date reading as a camera clock that lost its date, so the"
        Write-Both -Text $shown -Color Cyan
        Write-Both -Text "file system date is taken whole for them, its time of day included." -Color Cyan
    }

    if ($script:FolderWarned.Count -gt 0)
    {
        $shown = "$($script:FolderWarned.Count) folder name(s) do not keep to the naming convention, see $logPath"
        Write-Both -Text $shown -Color Yellow
    }

    if ($script:DateWarnings -gt 0)
    {
        $shown = "$script:DateWarnings warning(s) in all about dates and folder names, all of them in $logPath"
        Write-Both -Text $shown -Color Yellow
    }
}

if ($script:ExifTroubleGroups -gt 0)
{
    Write-Both -Text "$script:ExifTroubleGroups group(s) were not answered in full by exiftool, the files in them may be listed wrongly. See $logPath" -Color Yellow
}

if ($script:FolderReported -gt 0)
{
    Write-Both -Text "$toFix file(s) to fix, and $script:FolderReported more in the list that are only reported." -Color Cyan
}
else
{
    Write-Both -Text "$toFix file(s) to fix." -Color Cyan
}

Write-Host ""

if ($candidates.Count -gt 0 -and $showList)
{
    $candidates |
        Select-Object -Property @{ Name = "Name";       Expression = { $_.Name } },
                                @{ Name = "Metadata";   Expression = { if ($null -eq $_.ExifDate) { "-" } else { $_.ExifDate.ToString("yyyy-MM-dd HH:mm:ss") } } },
                                @{ Name = "FileSystem"; Expression = { $_.FileDate.ToString("yyyy-MM-dd HH:mm:ss") } },
                                @{ Name = "From";       Expression = { $_.DateFrom } },
                                @{ Name = "Writes";     Expression = { Get-FixDescription -Candidate $_ } },
                                @{ Name = "Reason";     Expression = { $_.Reason } } |
        Format-Table -AutoSize |
        Out-String -Width 200 |
        Write-Host
}

# ---------------------------------------------------------------------------
# Phase 2: rewrite the list, now in date order
# ---------------------------------------------------------------------------

# The counts of the run go to the log, while the csv keeps the fix list and a separate information section.
Write-Log "counts    : found $total, no date in the name $withoutName, not read $skipped, read $index of $toRead"
Write-Log "counts    : already right $nothingToDo, no date tag $noExifDate, tag without a real date $badExifDate"
Write-Log "counts    : need a file date $needFile, need metadata $needMeta"
Write-Log "counts    : listed because of their folder $script:FolderFixes, no folder date either $noFolderDate"
Write-Log "counts    : reported only, their folder too wide to write from $script:FolderReported"
Write-Log "counts    : metadata dates reading as a clock that lost its date $script:LostClocks"
Write-Log "counts    : warnings $script:DateWarnings, folder names against the convention $($script:FolderWarned.Count)"
Write-Log "counts    : left alone as a time zone, file dates $script:TimeZoneLeft, metadata $script:TimeZoneKept"
Write-Log "counts    : listed $($candidates.Count), of them to fix $toFix"
Write-Log "counts    : folder-judged file dates more than $FolderDayTolerance days outside the folder date $($largeFileDateDifferences.Count)"

if ($stopped)
{
    Write-Log "counts    : the scan was stopped by the user, the list is only what was read"
}

$report = New-Object System.Collections.Generic.List[string]
$report.Add((Get-CandidateHeader))

foreach ($candidate in $candidates)
{
    $report.Add((Get-CandidateLine -Candidate $candidate))
}

$report.Add("")
$report.Add((Get-CsvLine -Values @("Information: folder-judged files with file dates more than $FolderDayTolerance days outside the folder date")))
$report.Add((Get-CsvLine -Values @("File", "Folder", "FileSystemDate", "FolderDate", "DaysOutside")))

if ($largeFileDateDifferences.Count -gt 0)
{
    foreach ($difference in ($largeFileDateDifferences | Sort-Object -Property @{ Expression = "DaysOutside"; Descending = $true }, @{ Expression = "Name"; Descending = $true }))
    {
        $report.Add((Get-CsvLine -Values @($difference.Name,
                                          $difference.Folder,
                                          $difference.FileDate.ToString("yyyy-MM-dd HH:mm:ss"),
                                          $difference.FolderText,
                                          $difference.DaysOutside)))
    }
}
else
{
    $report.Add((Get-CsvLine -Values @("No matching files", "", "", "", "")))
}

[System.IO.File]::WriteAllLines($reportPath, $report, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "List complete in $reportPath, $($candidates.Count) row(s) plus the header and $($largeFileDateDifferences.Count) large file-date difference(s)." -ForegroundColor Green
Write-Host ""

if ($candidates.Count -eq 0)
{
    Write-Both -Text "Nothing to do." -Color Green
    $script:LogWriter.Close()
    $script:LogWriter.Dispose()
    return
}

if ($toFix -eq 0)
{
    # Every row in the list is one of the reported ones, so the list is the whole of the answer.
    Write-Both -Text "Nothing to write, the listed file(s) all sit in folders too wide to take a day from." -Color Green
    $script:LogWriter.Close()
    $script:LogWriter.Dispose()
    return
}

if ($listOnly)
{
    Write-Both -Text "List only was requested, no files were changed." -Color Yellow
    $script:LogWriter.Close()
    $script:LogWriter.Dispose()
    return
}

# ---------------------------------------------------------------------------
# Phase 3: confirm and apply, one file at a time
# ---------------------------------------------------------------------------

if ($fixAll)
{
    Write-Host "Fixing all listed files without asking. Press Esc or Q to stop." -ForegroundColor Yellow
}
else
{
    Write-Host "Review each file. y = fix, n = skip, a = fix all remaining, q = quit" -ForegroundColor Cyan
}

# A date that comes from the name of a folder is a guess about the whole folder, so it is asked about
# even in the mode that fixes everything, unless the setting at the top says otherwise. Answering a
# switches that off as well, it means fix the rest and stop asking.
$askFolder  = $AlwaysAskForFolderDates
$folderAsks = $script:FolderFixes - $script:FolderReported

if ($fixAll -and $askFolder -and $folderAsks -gt 0)
{
    Write-Host "The $folderAsks file(s) whose date comes from a folder name are still asked about, one at a time." -ForegroundColor Yellow
}

Write-Host ""

$applyAll = $fixAll
$fixed    = 0
$failed   = 0
$ignored  = 0
$position = 0

# What was done to each file goes in the log, one line per file as it happens, so the log is the
# whole story of the run: the counts at the top, then every file that was written, skipped or failed.
Write-Log ""
Write-Log "--- what was done, one line per file ---"

Write-Host "Writing what is done to $logPath" -ForegroundColor Cyan
Write-Host ""

foreach ($candidate in $candidates)
{
    if ($candidate.ReportOnly)
    {
        # In the list for the record and nothing more: the folder around it is too wide for its edge to
        # be a day worth writing. Not shown and not asked about, the row in the csv says it all.
        continue
    }

    $position++
    $percent = [int](($position / $toFix) * 100)

    Write-Progress -Activity "Fixing" -Status ("{0,3}%   {1} of {2}   {3}" -f $percent, $position, $toFix, $candidate.Name) -PercentComplete $percent

    Write-Host ("[{0} of {1}, {2}%] {3}" -f $position, $toFix, $percent, $candidate.File) -ForegroundColor White

    if ($candidate.Source -eq "folder")
    {
        # Loud on purpose: the name of this file says nothing about when it was taken, the date below
        # is only what the folder around it is called.
        Write-Host ("  the file name holds no date, the folder it sits in says {0}" -f $candidate.FolderText) -ForegroundColor Yellow
    }

    Show-ExifInfo -Info $candidate.Info -Label "old metadata:"

    Write-Host "  file system:" -ForegroundColor DarkGray
    Write-Host ("    {0,-18} {1}" -f "LastWriteTime", $candidate.FileDate.ToString("yyyy:MM:dd HH:mm:ss"))

    # The day is the same in both, the time of day is each one's own, so both values are shown.
    $value     = $candidate.Target.ToString("yyyy:MM:dd HH:mm:ss")
    $metaValue = $candidate.MetaTarget.ToString("yyyy:MM:dd HH:mm:ss")

    if ($candidate.MetaWrong)
    {
        Write-Host "  new metadata:" -ForegroundColor DarkGray
        Write-Host ("    {0,-18} {1}" -f "DateTimeOriginal", $metaValue)
        Write-Host ("    {0,-18} {1}" -f "CreateDate", $metaValue)
        Write-Host ("    {0,-18} {1}" -f "ModifyDate", $metaValue)

        if ($candidate.IsVideo)
        {
            Write-Host ("    {0,-18} {1}" -f "MediaCreateDate", $metaValue)
            Write-Host ("    {0,-18} {1}" -f "TrackCreateDate", $metaValue)
        }

        if ($metaValue -ne $value)
        {
            Write-Host "    the day comes from the folder, the time of day is the one it already held" -ForegroundColor DarkGray

            if ($metaValue.Substring(0, 10) -ne $value.Substring(0, 10))
            {
                Write-Host "    that time of day falls on the other side of midnight, so the day goes with it" -ForegroundColor DarkGray
            }
        }
        elseif ($candidate.LostClock)
        {
            Write-Host "    the camera clock had lost its date, so the time of day comes from the file as well" -ForegroundColor DarkGray
        }
    }

    if ($candidate.FileWrong)
    {
        Write-Host "  new file system dates:" -ForegroundColor DarkGray
        Write-Host ("    {0,-18} {1}" -f "FileModifyDate", $value)
        Write-Host ("    {0,-18} {1}" -f "FileCreateDate", $value)
    }

    Write-Host ("  reason: {0}" -f $candidate.Reason) -ForegroundColor DarkGray

    $answer = "y"

    # A file whose date comes from a folder name is asked about even when everything else is being
    # written without asking, that is the whole point of the setting.
    $mustAsk = (-not $applyAll) -or ($askFolder -and $candidate.Source -eq "folder")

    if (-not $mustAsk)
    {
        # No prompt in this mode, so the key is the only way out.
        if (Test-StopKey)
        {
            $answer = "q"
        }
    }
    else
    {
        do
        {
            $answer = (Read-Host "  Apply? [y/n/a/q]").Trim().ToLower()
            if ($answer -eq "")
            {
                $answer = "n"
            }
        }
        while ($answer -notin @("y", "n", "a", "q"))
    }

    if ($answer -eq "q")
    {
        Write-Host "  Stopped by user." -ForegroundColor Yellow
        Write-Log "STOPPED $($candidate.File) : stopped here, nothing was written"
        break
    }

    if ($answer -eq "a")
    {
        # Fix the rest and stop asking, folder dates included, that is what was just answered about.
        $applyAll  = $true
        $askFolder = $false
        $answer    = "y"
    }

    if ($answer -eq "n")
    {
        Write-Host "  Skipped." -ForegroundColor DarkGray
        Write-Host ""
        $ignored++
        Write-Log "SKIPPED $($candidate.File)"
        continue
    }

    $outcome = Set-ExifDate -Tool $ExifTool -File $candidate.File -MetaStamp $candidate.MetaTarget `
                            -FileStamp $candidate.Target -IsVideo $candidate.IsVideo `
                            -WriteMetadata $candidate.MetaWrong -WriteFileTimes $candidate.FileWrong

    # The answer of exiftool can hold line breaks, they would spread one file over several log lines.
    $said = ($outcome.Output -replace '\r?\n', ' ').Trim()

    # What was written, which is two dates when the metadata kept a time of day of its own.
    if (-not $candidate.MetaWrong)
    {
        $written = $value
    }
    elseif (-not $candidate.FileWrong -or $metaValue -eq $value)
    {
        $written = $metaValue
    }
    else
    {
        $written = "$metaValue in the metadata, $value in the file date"
    }

    # Where the date written came from, so a line in the log can be judged on its own later.
    $from = ""

    if ($candidate.Source -eq "folder")
    {
        $from = " (from the folder $($candidate.FolderText))"
    }

    if ($outcome.Success)
    {
        # exiftool can write the file and still have something to say about it, a warning about a
        # container that is not quite clean for example. The file is fixed, the remark is kept.
        if ($said -eq "")
        {
            Write-Host "  Updated." -ForegroundColor Green
            Write-Log "FIXED   $($candidate.File) -> $written$from"
        }
        else
        {
            Write-Host "  Updated, exiftool said: $said" -ForegroundColor Green
            Write-Log "FIXED   $($candidate.File) -> $written$from : $said"
        }

        $fixed++
    }
    else
    {
        Write-Host "  Failed: $said" -ForegroundColor Red
        $failed++
        Write-Log "FAILED  $($candidate.File) -> $written$from : $said"
    }

    Write-Host ""
}

Write-Progress -Activity "Fixing" -Completed

if ($script:FolderReported -gt 0)
{
    $shown = "Done. fixed: $fixed, skipped: $ignored, failed: $failed, reported only: $script:FolderReported"
    Write-Both -Text $shown -Color Cyan
}
else
{
    Write-Both -Text "Done. fixed: $fixed, skipped: $ignored, failed: $failed" -Color Cyan
}

Write-Host "The list is in $reportPath" -ForegroundColor Green
Write-Host "What was done is in $logPath" -ForegroundColor Green

$script:LogWriter.Close()
$script:LogWriter.Dispose()
