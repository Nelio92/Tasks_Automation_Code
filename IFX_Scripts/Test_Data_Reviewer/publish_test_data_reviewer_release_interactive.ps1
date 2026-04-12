param(
    [string]$Version,
    [string]$ReleaseDir = "release_pyinstaller",
    [string]$PackageDir = "release_packages",
    [string]$TeamRepoUrl = $(if ($env:TDR_TEAM_REPO_URL) { $env:TDR_TEAM_REPO_URL } else { $env:TDR_GITLAB_PROJECT_URL }),
    [string]$TeamRepoBranch = $(if ($env:TDR_TEAM_REPO_BRANCH) { $env:TDR_TEAM_REPO_BRANCH } else { "main" }),
    [string]$TeamRepoSubdir = $(if ($env:TDR_TEAM_REPO_SUBDIR) { $env:TDR_TEAM_REPO_SUBDIR } else { "TestDataReviewer" }),
    [string]$TeamRepoToken = $(if ($env:TDR_TEAM_REPO_TOKEN) { $env:TDR_TEAM_REPO_TOKEN } else { $env:TDR_GITLAB_TOKEN }),
    [string]$TeamRepoUsername = $(if ($env:TDR_TEAM_REPO_USERNAME) { $env:TDR_TEAM_REPO_USERNAME } else { "oauth2" }),
    [string]$GitUserName = $(if ($env:TDR_RELEASE_GIT_USER_NAME) { $env:TDR_RELEASE_GIT_USER_NAME } else { "Wandji Lionel Wilfried (PSS RF D RAD PTE TE4)" }),
    [string]$GitUserEmail = $(if ($env:TDR_RELEASE_GIT_USER_EMAIL) { $env:TDR_RELEASE_GIT_USER_EMAIL } else { "LionelWilfried.Wandji@infineon.com" }),
    [switch]$SkipBuild,
    [switch]$NoPush
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

function Get-RepoRoot {
    param([string]$StartPath)

    $current = Resolve-Path $StartPath
    while ($true) {
        if (Test-Path (Join-Path $current ".git")) {
            return $current
        }
        $parent = Split-Path -Parent $current
        if (-not $parent -or $parent -eq $current) {
            throw "Could not locate git repository root from $StartPath"
        }
        $current = $parent
    }
}

function Get-ReleaseVersion {
    param([string]$ExplicitVersion, [string]$RepoRoot)

    if ($ExplicitVersion) {
        return $ExplicitVersion.Trim()
    }

    if ($env:GITHUB_REF_TYPE -eq "tag" -and $env:GITHUB_REF_NAME) {
        return $env:GITHUB_REF_NAME.Trim()
    }

    $shortSha = (git -C $RepoRoot rev-parse --short HEAD).Trim()
    if (-not $shortSha) {
        $shortSha = "unknown"
    }
    return "draft-$(Get-Date -Format 'yyyyMMdd-HHmmss')-$shortSha"
}

function Assert-ValidReleaseVersion {
    param(
        [string]$OriginalVersion,
        [string]$SanitizedVersion
    )

    if ([string]::IsNullOrWhiteSpace($SanitizedVersion)) {
        throw "Resolved release version is empty after sanitization. Provide a version like v1.0.1"
    }

    if ($SanitizedVersion.StartsWith("-")) {
        throw "Invalid release version '$OriginalVersion'. Versions must start with a letter or digit, for example v1.0.1"
    }

    if ($SanitizedVersion -notmatch '^[0-9A-Za-z]') {
        throw "Invalid release version '$OriginalVersion'. Versions must start with a letter or digit, for example v1.0.1"
    }
}

function Assert-OfficialReleaseVersion {
    param([string]$ReleaseVersion)

    if ($ReleaseVersion -notmatch '^v\d+\.\d+\.\d+$') {
        throw "Invalid official release version '$ReleaseVersion'. Use semantic version format v<major>.<minor>.<patch>, for example v1.0.1"
    }
}

function Get-AuthenticatedRepoUrl {
    param(
        [string]$RepoUrl,
        [string]$Token,
        [string]$Username
    )

    if (-not $Token) {
        return $RepoUrl
    }
    if (-not $RepoUrl.StartsWith("https://", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Authenticated publishing currently supports only https repository URLs"
    }

    $uri = [System.Uri]$RepoUrl
    $builder = [System.UriBuilder]::new($uri)
    $builder.UserName = [System.Uri]::EscapeDataString($Username)
    $builder.Password = [System.Uri]::EscapeDataString($Token)
    return $builder.Uri.AbsoluteUri
}

function Write-ReleaseMetadata {
    param(
        [string]$MetadataPath,
        [hashtable]$Metadata
    )

    $Metadata | ConvertTo-Json -Depth 6 | Set-Content -Path $MetadataPath -Encoding UTF8
}

function Copy-ReleaseArtifacts {
    param(
        [string]$SourceZipPath,
        [string]$ReleaseVersion,
        [string]$RepoPublishRoot
    )

    $zipFileName = Split-Path -Leaf $SourceZipPath
    $versionDir = Join-Path $RepoPublishRoot (Join-Path "releases" $ReleaseVersion)
    $latestDir = Join-Path $RepoPublishRoot "latest"
    New-Item -ItemType Directory -Force -Path $versionDir | Out-Null
    New-Item -ItemType Directory -Force -Path $latestDir | Out-Null

    Copy-Item $SourceZipPath (Join-Path $versionDir $zipFileName) -Force

    $latestZipName = "TestDataReviewer-latest.zip"
    Copy-Item $SourceZipPath (Join-Path $latestDir $latestZipName) -Force
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Get-RepoRoot -StartPath $scriptRoot

if (-not $Version) {
    $Version = Read-RequiredValue -Prompt "Release version" -DefaultValue "v1.0.3"
}
if (-not $TeamRepoUrl -and -not $NoPush) {
    $TeamRepoUrl = Read-RequiredValue -Prompt "GitLab team repo URL" -DefaultValue "https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git"
}
if (-not $TeamRepoUsername -and -not $NoPush) {
    $TeamRepoUsername = Read-RequiredValue -Prompt "GitLab username for token auth" -DefaultValue "oauth2"
}
if (-not $GitUserName -and -not $NoPush) {
    $GitUserName = Read-RequiredValue -Prompt "Git commit author name" -DefaultValue "Wandji Lionel Wilfried (PSS RF D RAD PTE TE4)"
}
if (-not $GitUserEmail -and -not $NoPush) {
    $GitUserEmail = Read-RequiredValue -Prompt "Git commit author email" -DefaultValue "LionelWilfried.Wandji@infineon.com"
}
if (-not $TeamRepoToken -and -not $NoPush) {
    $TeamRepoToken = Read-SecretValue -Prompt "GitLab personal access token"
}

$resolvedReleaseVersion = Get-ReleaseVersion -ExplicitVersion $Version -RepoRoot $repoRoot
$releaseVersion = ($resolvedReleaseVersion -replace '[^0-9A-Za-z._-]', '-')
Assert-ValidReleaseVersion -OriginalVersion $resolvedReleaseVersion -SanitizedVersion $releaseVersion
if (-not $NoPush) {
    Assert-OfficialReleaseVersion -ReleaseVersion $releaseVersion
}

$releaseRoot = Join-Path $scriptRoot $ReleaseDir
$packageRoot = Join-Path $scriptRoot $PackageDir
$versionPackageDir = Join-Path $packageRoot $releaseVersion
$zipFileName = "TestDataReviewer-$releaseVersion.zip"
$zipPath = Join-Path $versionPackageDir $zipFileName
$hashPath = Join-Path $versionPackageDir "$zipFileName.sha256.txt"
$metadataPath = Join-Path $versionPackageDir "release-metadata.json"

if (-not $SkipBuild) {
    & (Join-Path $scriptRoot "build_test_data_reviewer_exe.ps1") -ReleaseDir $ReleaseDir
}

if (-not (Test-Path $releaseRoot)) {
    throw "Release directory not found: $releaseRoot"
}

New-Item -ItemType Directory -Force -Path $versionPackageDir | Out-Null
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

Compress-Archive -Path (Join-Path $releaseRoot '*') -DestinationPath $zipPath -Force
$hash = (Get-FileHash -Path $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
"SHA256  $zipFileName  $hash" | Set-Content -Path $hashPath -Encoding ASCII

$originUrl = (git -C $repoRoot remote get-url origin).Trim()
$commitSha = (git -C $repoRoot rev-parse HEAD).Trim()
$metadata = @{
    tool = "TestDataReviewer"
    version = $releaseVersion
    builtAtUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    sourceRepository = $originUrl
    sourceCommit = $commitSha
    zipFile = $zipFileName
    sha256 = $hash
}
Write-ReleaseMetadata -MetadataPath $metadataPath -Metadata $metadata

Write-Host "Release package created: $zipPath"
Write-Host "SHA256 file created:  $hashPath"

if (-not $TeamRepoUrl) {
    Write-Warning "No team repository URL provided. Skipping team-repo publish step."
    exit 0
}

if ($NoPush) {
    Write-Host "Skipping team-repo publish because -NoPush was set."
    exit 0
}

$authenticatedRepoUrl = Get-AuthenticatedRepoUrl -RepoUrl $TeamRepoUrl -Token $TeamRepoToken -Username $TeamRepoUsername
$tempCloneRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("tdr-team-release-" + [System.Guid]::NewGuid().ToString("N"))
$cloneDir = Join-Path $tempCloneRoot "team-repo"
New-Item -ItemType Directory -Force -Path $tempCloneRoot | Out-Null

try {
    git clone --depth 1 --branch $TeamRepoBranch $authenticatedRepoUrl $cloneDir | Out-Null
    $repoPublishRoot = Join-Path $cloneDir $TeamRepoSubdir
    Copy-ReleaseArtifacts -SourceZipPath $zipPath -ReleaseVersion $releaseVersion -RepoPublishRoot $repoPublishRoot

    Write-Host "Team repo clone:   $cloneDir"
    Write-Host "Publish root:      $repoPublishRoot"
    Write-Host "Versioned zip path: $(Join-Path $repoPublishRoot (Join-Path ('releases/' + $releaseVersion) $zipFileName))"
    Write-Host "Latest zip path:    $(Join-Path $repoPublishRoot 'latest/TestDataReviewer-latest.zip')"

    git -C $cloneDir config user.name $GitUserName
    git -C $cloneDir config user.email $GitUserEmail
    git -C $cloneDir add --all --force
    Write-Host "Git status after staging:"
    git -C $cloneDir status --short --ignored

    $stagedChanges = git -C $cloneDir diff --cached --name-only
    $hasChanges = -not [string]::IsNullOrWhiteSpace(($stagedChanges | Out-String))

    if (-not $hasChanges) {
        Write-Host "Team repository already up to date for $releaseVersion"
        exit 0
    }

    git -C $cloneDir commit -m "Publish TestDataReviewer $releaseVersion" | Out-Null
    git -C $cloneDir push origin $TeamRepoBranch | Out-Null
    Write-Host "Published TestDataReviewer $releaseVersion to team repository branch $TeamRepoBranch"
}
finally {
    if (Test-Path $tempCloneRoot) {
        Remove-Item -Path $tempCloneRoot -Recurse -Force
    }
}