param(
    [string]$Version,
    [string]$ReleaseDir = "release_pyinstaller",
    [string]$PackageDir = "release_packages",
    [string]$GitLabProjectUrl = $(if ($env:TDR_GITLAB_PROJECT_URL) { $env:TDR_GITLAB_PROJECT_URL } else { $env:TDR_TEAM_REPO_URL }),
    [string]$GitLabToken = $(if ($env:TDR_GITLAB_TOKEN) { $env:TDR_GITLAB_TOKEN } else { $env:TDR_TEAM_REPO_TOKEN }),
    [string]$GenericPackageName = $(if ($env:TDR_GITLAB_PACKAGE_NAME) { $env:TDR_GITLAB_PACKAGE_NAME } else { "test-data-reviewer" }),
    [string]$LatestVersionLabel = $(if ($env:TDR_GITLAB_LATEST_LABEL) { $env:TDR_GITLAB_LATEST_LABEL } else { "latest" }),
    [switch]$SkipBuild,
    [switch]$NoUpload
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

function Get-GitLabProjectPath {
    param(
        [string]$ProjectUrl
    )

    $uri = [System.Uri]$ProjectUrl
    $projectPath = $uri.AbsolutePath.Trim('/')
    if ($projectPath.EndsWith('.git', [System.StringComparison]::OrdinalIgnoreCase)) {
        $projectPath = $projectPath.Substring(0, $projectPath.Length - 4)
    }
    if (-not $projectPath) {
        throw "Could not derive GitLab project path from $ProjectUrl"
    }
    return $projectPath
}

function Get-GitLabApiBaseUrl {
    param(
        [string]$ProjectUrl
    )

    $uri = [System.Uri]$ProjectUrl
    return ($uri.GetLeftPart([System.UriPartial]::Authority).TrimEnd('/') + "/api/v4")
}

function Assert-GitLabHostResolvable {
    param(
        [string]$ProjectUrl
    )

    $uri = [System.Uri]$ProjectUrl
    try {
        [System.Net.Dns]::GetHostEntry($uri.Host) | Out-Null
    }
    catch {
        throw "GitLab host '$($uri.Host)' is not resolvable from this runner. If this is an internal-only GitLab instance, use a self-hosted GitHub Actions runner on the corporate network or run publish_test_data_reviewer_release.ps1 locally from a machine that can reach it."
    }
}

function Write-ReleaseMetadata {
    param(
        [string]$MetadataPath,
        [hashtable]$Metadata
    )

    $Metadata | ConvertTo-Json -Depth 6 | Set-Content -Path $MetadataPath -Encoding UTF8
}

function Upload-GitLabGenericPackageFile {
    param(
        [string]$ProjectUrl,
        [string]$Token,
        [string]$PackageName,
        [string]$PackageVersion,
        [string]$FilePath,
        [string]$FileName
    )

    if (-not $ProjectUrl) {
        throw "GitLab project URL is required for package-registry upload"
    }
    if (-not $Token) {
        throw "GitLab token is required for package-registry upload"
    }

    $projectPath = Get-GitLabProjectPath -ProjectUrl $ProjectUrl
    $apiBaseUrl = Get-GitLabApiBaseUrl -ProjectUrl $ProjectUrl
    $encodedProject = [System.Uri]::EscapeDataString($projectPath)
    $encodedPackage = [System.Uri]::EscapeDataString($PackageName)
    $encodedVersion = [System.Uri]::EscapeDataString($PackageVersion)
    $encodedFileName = [System.Uri]::EscapeDataString($FileName)
    $uploadUrl = "$apiBaseUrl/projects/$encodedProject/packages/generic/$encodedPackage/$encodedVersion/$encodedFileName"

    Invoke-WebRequest `
        -Uri $uploadUrl `
        -Method Put `
        -Headers @{ "PRIVATE-TOKEN" = $Token } `
        -InFile $FilePath `
        -ContentType "application/octet-stream" | Out-Null

    return $uploadUrl
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Get-RepoRoot -StartPath $scriptRoot
$releaseVersion = Get-ReleaseVersion -ExplicitVersion $Version -RepoRoot $repoRoot
$releaseVersion = ($releaseVersion -replace '[^0-9A-Za-z._-]', '-')

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

if (-not $GitLabProjectUrl) {
    Write-Warning "No GitLab project URL provided. Skipping package-registry upload step."
    exit 0
}

if ($NoUpload) {
    Write-Host "Skipping package-registry upload because -NoUpload was set."
    exit 0
}

Assert-GitLabHostResolvable -ProjectUrl $GitLabProjectUrl

$versionUploadUrls = @{}
$versionUploadUrls[$zipFileName] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $releaseVersion -FilePath $zipPath -FileName $zipFileName
$versionUploadUrls["$zipFileName.sha256.txt"] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $releaseVersion -FilePath $hashPath -FileName "$zipFileName.sha256.txt"
$versionUploadUrls["release-metadata.json"] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $releaseVersion -FilePath $metadataPath -FileName "release-metadata.json"

$latestZipName = "TestDataReviewer-latest.zip"
$latestHashPath = Join-Path $versionPackageDir "TestDataReviewer-latest.sha256.txt"
$latestMetadataPath = Join-Path $versionPackageDir "release-metadata-latest.json"
$latestVersionPath = Join-Path $versionPackageDir "LATEST_VERSION.txt"
"SHA256  $latestZipName  $hash" | Set-Content -Path $latestHashPath -Encoding ASCII
$metadataWithLatest = @{}
foreach ($key in $metadata.Keys) {
    $metadataWithLatest[$key] = $metadata[$key]
}
$metadataWithLatest["latestLabel"] = $LatestVersionLabel
$metadataWithLatest["publishedVersion"] = $releaseVersion
Write-ReleaseMetadata -MetadataPath $latestMetadataPath -Metadata $metadataWithLatest
$releaseVersion | Set-Content -Path $latestVersionPath -Encoding ASCII

$latestUploadUrls = @{}
$latestUploadUrls[$latestZipName] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $LatestVersionLabel -FilePath $zipPath -FileName $latestZipName
$latestUploadUrls["TestDataReviewer-latest.sha256.txt"] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $LatestVersionLabel -FilePath $latestHashPath -FileName "TestDataReviewer-latest.sha256.txt"
$latestUploadUrls["release-metadata.json"] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $LatestVersionLabel -FilePath $latestMetadataPath -FileName "release-metadata.json"
$latestUploadUrls["LATEST_VERSION.txt"] = Upload-GitLabGenericPackageFile -ProjectUrl $GitLabProjectUrl -Token $GitLabToken -PackageName $GenericPackageName -PackageVersion $LatestVersionLabel -FilePath $latestVersionPath -FileName "LATEST_VERSION.txt"

Write-Host "Uploaded release package to GitLab Generic Package Registry"
Write-Host "Package name: $GenericPackageName"
Write-Host "Versioned package: $releaseVersion"
Write-Host "Latest label: $LatestVersionLabel"