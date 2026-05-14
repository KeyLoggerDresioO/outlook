Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$start = Get-Date
$ts    = $start.ToString('yyyy-MM-dd_HH-mm')
$log   = "$env:ProgramData\ESET\RemoteAdministrator\Agent\EraAgentApplicationData\Logs\deploy_outlook_pwa_$ts.txt"

function wl {
    param([string]$m, [string]$l = 'INFO')
    $t = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $e = "[$t] [$l] $m"
    $e | Out-File $log -Append -Encoding utf8
}

$hdr = @(
    '================================================================================'
    'Script Execution Log'
    '================================================================================'
    "Script Name    : deploy_outlook_pwa"
    "Log File       : $log"
    "Start Time     : $($start.ToString('yyyy-MM-dd HH:mm:ss'))"
    "Computer Name  : $env:COMPUTERNAME"
    "User Context   : $env:USERNAME"
    "OS Version     : $([System.Environment]::OSVersion.VersionString)"
    "PowerShell Ver : $($PSVersionTable.PSVersion.ToString())"
    '================================================================================'
    ''
) -join [Environment]::NewLine

$hdr | Out-File $log -Encoding utf8 -Force

function New-Shortcut {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetPath,

        [Parameter()]
        [string]$ShortcutPath = (Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath 'New Shortcut.lnk'),

        [Parameter()]
        [string[]]$Arguments,

        [Parameter()]
        [string[]]$HotKey,

        [Parameter()]
        [string]$WorkingDirectory,

        [Parameter()]
        [string]$Description,

        [Parameter(ParameterSetName = 'IconDownload', Mandatory)]
        [string]$IconName,

        [Parameter(ParameterSetName = 'IconDownload', Mandatory)]
        [string]$IconURL,

        [Parameter(ParameterSetName = 'IconDownload')]
        [string]$IconPath = 'C:\Temp',

        [Parameter()]
        [ValidateSet('Default', 'Maximized', 'Minimized')]
        [string]$WindowStyle = 'Default',

        [Parameter()]
        [switch]$RunAsAdmin
    )
    begin {
        if ($IconURL) {
            Write-Verbose "Downloading icon from $IconURL"
            if (-Not (Test-Path -Path $IconPath)) {
                Write-Verbose "Creating directory $IconPath"
                New-Item -ItemType Directory -Path $IconPath -Force
            }
            $IconPath = $IconPath.TrimEnd('\')
            Write-Verbose "Downloading icon to $IconPath\$IconName"
            Invoke-WebRequest -Uri $IconURL -OutFile "$IconPath\$IconName"
            $IconLocation = "$IconPath\$IconName"
        }
    }
    Process {
        switch ($WindowStyle) {
            'Default'   { $style = 1; break }
            'Maximized' { $style = 3; break }
            'Minimized' { $style = 7 }
        }
        $WshShell = New-Object -ComObject WScript.Shell
        $shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $shortcut.TargetPath = $TargetPath
        $shortcut.WindowStyle = $style
        if ($Arguments)        { $shortcut.Arguments = $Arguments -join ' ' }
        if ($HotKey)           { $shortcut.Hotkey = ($HotKey -join '+').ToUpperInvariant() }
        if ($IconLocation)     { $shortcut.IconLocation = $IconLocation }
        if ($Description)      { $shortcut.Description = $Description }
        if ($WorkingDirectory) { $shortcut.WorkingDirectory = $WorkingDirectory }

        $shortcut.Save()

        if ($RunAsAdmin) {
            [byte[]]$bytes = [System.IO.File]::ReadAllBytes($ShortcutPath)
            $bytes[21] = $bytes[21] -bor 32
            [System.IO.File]::WriteAllBytes($ShortcutPath, $bytes)
        }
    }
    End {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

try {
    # Gdy skrypt dziala w kontekscie SYSTEM (np. ESET Agent), HKCU nie istnieje.
    # Odczytaj domyslna przegladarke z profilu pierwszego zalogowanego uzytkownika;
    # jesli nie uda sie — uzyj Microsoft Edge jako fallback.
    # Edge is always preferred — check both Program Files locations first.
    $edgePath    = Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"
    $edgePathX86 = Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"

    wl "Checking for Microsoft Edge installation..." 'DEBUG'

    if (Test-Path $edgePath) {
        $DefaultBrowserPath = $edgePath
        wl "Edge found (x64): $edgePath"
    } elseif (Test-Path $edgePathX86) {
        $DefaultBrowserPath = $edgePathX86
        wl "Edge found (x86): $edgePathX86"
    } else {
        # Edge not found — fall back to the default browser of the first logged-on user.
        wl "Edge not found. Falling back to default browser from user registry hive." 'WARNING'
        $DefaultBrowserPath = $null

        $userProfiles = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" |
            Where-Object { $_.GetValue("ProfileImagePath") -match "C:\\Users\\" -and $_.GetValue("ProfileImagePath") -notmatch "systemprofile|LocalService|NetworkService" } |
            Select-Object -ExpandProperty PSPath

        foreach ($profile in $userProfiles) {
            $sid = Split-Path -Leaf $profile
            try {
                $hive = "HKU:\$sid\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice"
                if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
                    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
                }
                $progId = (Get-ItemProperty $hive -ErrorAction Stop).ProgId
                $command = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("$progId\shell\open\command").GetValue("")
                if ($command -match '"([^"]*)"') {
                    $DefaultBrowserPath = $matches[1]
                    wl "Default browser resolved from SID $sid : $DefaultBrowserPath"
                    break
                }
            } catch {
                wl "Could not read browser registry for SID $sid : $_" 'DEBUG'
                continue
            }
        }

        if (-not $DefaultBrowserPath) {
            wl "Cannot determine browser path and Edge not found." 'ERROR'
            throw "[!] Cannot determine browser path and Edge not found."
        }
    }

    $shortcutPath = Join-Path -Path ([Environment]::GetFolderPath("CommonDesktopDirectory")) -ChildPath 'Outlook (PWA).lnk'
    wl "Creating shortcut: $shortcutPath"
    wl "Target browser  : $DefaultBrowserPath" 'DEBUG'

    $props = @{
        'ShortcutPath' = $shortcutPath
        'TargetPath'   = $DefaultBrowserPath
        'Arguments'    = '-profile-directory=Default -app=https://outlook.cloud.microsoft/mail/'
        'IconName'     = "outlook_pwa.ico"
        'IconURL'      = "https://raw.githubusercontent.com/KeyLoggerDresioO/outlook/refs/heads/main/outlook_pwa.ico"
        'IconPath'     = "C:\Temp"
    }

    New-Shortcut @props
    wl "Shortcut created successfully." 'SUCCESS'

    $end = Get-Date
    $dur = $end - $start
    $ftr = @(
        ''
        '================================================================================'
        'Script Execution Complete'
        '================================================================================'
        "End Time       : $($end.ToString('yyyy-MM-dd HH:mm:ss'))"
        "Duration       : $($dur.ToString())"
        "Exit Status    : SUCCESS"
        '================================================================================'
    ) -join [Environment]::NewLine
    $ftr | Out-File $log -Append -Encoding utf8
    Write-Output "[OK] Udalo sie stworzyc skrot do aplikacji Outlook (PWA) na pulpicie."
}
catch {
    wl "Unhandled error: $_" 'ERROR'
    $end = Get-Date
    $dur = $end - $start
    $ftr = @(
        ''
        '================================================================================'
        'Script Execution Complete'
        '================================================================================'
        "End Time       : $($end.ToString('yyyy-MM-dd HH:mm:ss'))"
        "Duration       : $($dur.ToString())"
        "Exit Status    : FAILED"
        '================================================================================'
    ) -join [Environment]::NewLine
    $ftr | Out-File $log -Append -Encoding utf8
    Write-Output "[!] Nie udalo sie stworzyc skrotu do aplikacji Outlook (PWA) na pulpicie. Powod: $_"
    exit 1
}
