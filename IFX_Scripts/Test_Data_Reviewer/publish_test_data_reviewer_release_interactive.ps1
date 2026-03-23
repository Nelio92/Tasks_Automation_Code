param(
    [string]$Version,
    [switch]$SkipBuild
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Read-RequiredValue {
    param(
        [string]$Prompt,
        [string]$DefaultValue = ""
    )

    if ($DefaultValue) {
        $value = Read-Host "$Prompt [$DefaultValue]"
        if ([string]::IsNullOrWhiteSpace($value)) {
            return $DefaultValue
        }
        return $value.Trim()
    }

    while ($true) {
        $value = Read-Host $Prompt
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value.Trim()
        }
    }
}

function Read-SecretValue {
    param([string]$Prompt)

    while ($true) {
        $secureValue = Read-Host $Prompt -AsSecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureValue)
        try {
            $plainValue = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }

        if (-not [string]::IsNullOrWhiteSpace($plainValue)) {
            return $plainValue
        }
    }
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$resolvedVersion = $Version
if (-not $resolvedVersion) {
    $resolvedVersion = Read-RequiredValue -Prompt "Release version" -DefaultValue "v1.0.0"
}

$env:TDR_TEAM_REPO_URL = Read-RequiredValue -Prompt "GitLab team repo URL" -DefaultValue "https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git"
$env:TDR_TEAM_REPO_USERNAME = Read-RequiredValue -Prompt "GitLab username for token auth" -DefaultValue "oauth2"
$env:TDR_RELEASE_GIT_USER_NAME = Read-RequiredValue -Prompt "Git commit author name" -DefaultValue "Wandji Lionel Wilfried (PSS RF D RAD PTE TE4)"
$env:TDR_RELEASE_GIT_USER_EMAIL = Read-RequiredValue -Prompt "Git commit author email" -DefaultValue "LionelWilfried.Wandji@infineon.com"
$env:TDR_TEAM_REPO_TOKEN = Read-SecretValue -Prompt "GitLab personal access token"

$publishArgs = @('-Version', $resolvedVersion)
if ($SkipBuild) {
    $publishArgs += '-SkipBuild'
}

& (Join-Path $scriptRoot 'publish_test_data_reviewer_release.ps1') @publishArgs