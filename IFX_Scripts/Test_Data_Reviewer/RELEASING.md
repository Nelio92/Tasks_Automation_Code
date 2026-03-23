# Test Data Reviewer - Release Publishing

This project now supports a simple release flow that packages `release_pyinstaller` into a zip file and publishes only zip files into a separate GitLab repository that the team can browse directly.

## Release model

- Source code stays in this private GitHub repository.
- Team delivery happens through versioned zip artifacts committed into the team-facing GitLab repository.
- No Python source files, build scripts, requirements files, or other source assets are published to the GitLab repo.
- Each published version contains:
  - `TestDataReviewer-<version>.zip`
- The team repository is updated in two places:
  - `TestDataReviewer/releases/<version>/TestDataReviewer-<version>.zip`
  - `TestDataReviewer/latest/TestDataReviewer-latest.zip`

## Local publishing

Run from `IFX_Scripts/Test_Data_Reviewer`:

```powershell
./publish_test_data_reviewer_release.ps1 -Version v1.0.0 -TeamRepoUrl https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git
```

This is the recommended path for the official team release because your local machine can already reach the internal GitLab host.

The GitLab repo receives only the zipped release artifacts. The source code remains only in the private GitHub repository.

Optional parameters:

- `-SkipBuild`
  - reuse the existing `release_pyinstaller` folder
- `-NoPush`
  - build the zip locally without pushing to the team GitLab repo
- `-TeamRepoBranch main`
- `-TeamRepoSubdir TestDataReviewer`

Environment variables supported:

- `TDR_TEAM_REPO_URL`
- default target for this project: `https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git`
- `TDR_TEAM_REPO_BRANCH`
- `TDR_TEAM_REPO_SUBDIR`
- `TDR_TEAM_REPO_TOKEN`
- `TDR_TEAM_REPO_USERNAME`
- `TDR_RELEASE_GIT_USER_NAME`
- `TDR_RELEASE_GIT_USER_EMAIL`

If `TDR_TEAM_REPO_TOKEN` is set, the script injects credentials into the HTTPS clone URL for push access.

Recommended values:

- `TDR_TEAM_REPO_URL = https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git`
- `TDR_TEAM_REPO_USERNAME = oauth2`
- `TDR_RELEASE_GIT_USER_NAME = Wandji Lionel Wilfried (PSS RF D RAD PTE TE4)`
- `TDR_RELEASE_GIT_USER_EMAIL = LionelWilfried.Wandji@infineon.com`

## First official release command

On your internal Windows machine, open PowerShell in `IFX_Scripts/Test_Data_Reviewer` and run:

```powershell
$env:TDR_TEAM_REPO_URL = "https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git"
$env:TDR_TEAM_REPO_TOKEN = "<your GitLab personal access token>"
$env:TDR_TEAM_REPO_USERNAME = "oauth2"
$env:TDR_RELEASE_GIT_USER_NAME = "Wandji Lionel Wilfried (PSS RF D RAD PTE TE4)"
$env:TDR_RELEASE_GIT_USER_EMAIL = "LionelWilfried.Wandji@infineon.com"

./publish_test_data_reviewer_release.ps1 -Version v1.0.0
```

## GitHub Actions automation

The workflow file is:

- `.github/workflows/publish-test-data-reviewer-release.yml`

It supports:

- manual package build through `workflow_dispatch`
- automatic package build when a tag matching `tdr-v*` is pushed

The workflow now uses `windows-latest` again and is intended only as a hosted build check. It packages the release with `-NoPush` so it does not try to reach the internal GitLab server.

## Recommended versioning

Use semantic versions for official releases:

- `v1.0.0`
- `v1.0.1`
- `v1.1.0`

Recommended tagging command:

```powershell
git tag tdr-v1.0.0
git push origin tdr-v1.0.0
```

The workflow strips the `tdr-` prefix and packages the release as `v1.0.0`.