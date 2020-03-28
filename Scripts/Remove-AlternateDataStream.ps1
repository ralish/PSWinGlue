<#
    .SYNOPSIS
    Remove common unwanted alternate data streams from files

    .DESCRIPTION
    Alternate data streams can contain important (meta)data which shouldn't be removed, however, they can also be a nuisance.

    This is particularly true of certain common alternate data streams used to perform additional prompting of the user for "untrusted" files.

    This cmdlet provides options to remove common alternate data streams which may be unwanted while preserving any other alternate data streams.

    .PARAMETER Dropbox
    Remove alternate data streams added by Dropbox.

    These data streams are not publicly documented but appear to at least contain a unique machine identifier for tracking purposes.

    .PARAMETER Path
    Directory from which to remove specified alternate data streams from files.

    .PARAMETER Recurse
    Recurse into subdirectories.

    .PARAMETER ZoneIdentifier
    Remove the Zone Identifier alternate data stream.

    This stream indicates the origin of a downloaded file and is typically used to trigger additional prompts or protections on opening "untrusted" files.

    .EXAMPLE
    Remove-AlternateDataStreams -Path D:\Library -ZoneIdentifier -Recurse

    Removes all Zone.Identifier alternate data streams recursively from files in D:\Library.

    .NOTES
    For bulk removal of all alternate data streams consider using Sysinternals Streams:
    https://docs.microsoft.com/en-us/sysinternals/downloads/streams

    .LINK
    https://github.com/ralish/PSWinGlue
#>

#Requires -Version 3.0

[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory)]
    [String]$Path,

    [Switch]$ZoneIdentifier,
    [Switch]$Dropbox,

    [Switch]$Recurse
)

try {
    $AdsPath = Get-Item -Path $Path
} catch {
    throw ('Provided path is invalid: {0}' -f $Path)
}

if ($AdsPath -isnot [IO.DirectoryInfo]) {
    throw ('Expected directory but received: {0}' -f $AdsPath.GetType().Name)
}

$Streams = @()

if ($ZoneIdentifier) {
    $Streams += 'Zone.Identifier'
}

if ($Dropbox) {
    $Streams += 'com.dropbox.attrs'
    $Streams += 'com.dropbox.attributes'
}

if (!$Streams) {
    throw 'No alternate data streams to remove were specified.'
}

Get-ChildItem -Path $AdsPath -Recurse:$Recurse -File | Remove-Item -Stream $Streams
