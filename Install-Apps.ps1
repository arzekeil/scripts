param (
    [switch]$Help
)

# Display Help if -Help flag is used
if ($Help) {
    Write-Host "Usage:"
    Write-Host "  Install-Apps.ps1 [-Help]"
    Write-Host ""
    Write-Host "Options:"
    Write-Host "  -Help            Show this help message."
    exit
}

$appIDs = @(
    @{name = "Microsoft.VisualStudioCode" }
    @{name = "dbeaver.dbeaver" }
    @{name = "WinSCP.WinSCP" }
    @{name = "Python.Launcher" }
    @{name = "Git.Git" }
    @{name = "StrawberryPerl.StrawberryPerl" }
    @{name = "Microsoft.PowerShell" }
    @{name = "Neovim.Neovim"}
    @{name = "MiKTeX.MiKTeX" }
    @{name = "Python.Python.3.13" }
    @{name = "Brave.Brave" }
    @{name = "Discord.Discord" }
    @{name = "Spotify.Spotify" }
    @{name = "TheDocumentFoundation.LibreOffice" }
    @{name = "Bitwarden.Bitwarden" }
    @{name = "JGraph.Draw" }
);

Foreach ($appID in $appIDs) {
    #check if the app is already installed
    $listApp = winget list --exact -q $appID.name
    if (![String]::Join("", $listApp).Contains($appID.name)) {
        Write-host "Installing:" $appID.name
        if ($null -ne $appID.source) {
            winget install --exact $appID.name --source $appID.source --accept-package-agreements
        }
        else {
            winget install --exact $appID.name --accept-package-agreements
        }
    }
    else {
        Write-host "Skipping Install of" $appID.name
    }
}

Write-Host "Installions Completed."
Write-Host "Exiting Now. Goodbye."
