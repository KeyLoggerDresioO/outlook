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

# Gdy skrypt dziala w kontekscie SYSTEM (np. ESET Agent), HKCU nie istnieje.
# Odczytaj domyslna przegladarke z profilu pierwszego zalogowanego uzytkownika;
# jesli nie uda sie — uzyj Microsoft Edge jako fallback.
# Edge is always preferred — check both Program Files locations first.
$edgePath    = Join-Path $env:ProgramFiles "Microsoft\Edge\Application\msedge.exe"
$edgePathX86 = Join-Path ${env:ProgramFiles(x86)} "Microsoft\Edge\Application\msedge.exe"

if (Test-Path $edgePath) {
    $DefaultBrowserPath = $edgePath
} elseif (Test-Path $edgePathX86) {
    $DefaultBrowserPath = $edgePathX86
} else {
    # Edge not found — fall back to the default browser of the first logged-on user.
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
                break
            }
        } catch {
            continue
        }
    }

    if (-not $DefaultBrowserPath) {
        throw "[!] Cannot determine browser path and Edge not found."
    }
}

$props = @{
    'ShortcutPath' = Join-Path -Path ([Environment]::GetFolderPath("CommonDesktopDirectory")) -ChildPath 'Outlook (PWA).lnk'
    'TargetPath'   = $DefaultBrowserPath
    'Arguments'    = '-profile-directory=Default -app=https://outlook.cloud.microsoft/mail/'
    'IconName'     = "outlook_pwa.ico"
    'IconURL'      = "https://raw.githubusercontent.com/KeyLoggerDresioO/outlook/refs/heads/main/outlook_pwa.ico"
    'IconPath'     = "C:\Temp"
}

New-Shortcut @props
