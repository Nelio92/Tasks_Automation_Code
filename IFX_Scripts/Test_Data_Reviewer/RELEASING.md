# Test Data Reviewer - Release Publishing

This project now supports an automated release flow for distributing the packaged `release_pyinstaller` folder as artifacts in the GitLab Generic Package Registry.

## Release model

- Source code stays in this private GitHub repository.
- Team delivery happens through versioned artifacts published to the GitLab Generic Package Registry of the team-facing GitLab project.
- Each published version contains:
  - `TestDataReviewer-<version>.zip`
  - SHA256 checksum file
  - release metadata JSON
- Each release is uploaded twice:
  - once under the real version such as `v1.0.0`
  - once under a rolling `latest` label for easy access

## Local publishing

Run from `IFX_Scripts/Test_Data_Reviewer`:

```powershell
./publish_test_data_reviewer_release.ps1 -Version v1.0.0 -GitLabProjectUrl https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git
```

Optional parameters:

- `-SkipBuild`
  - reuse the existing `release_pyinstaller` folder
- `-NoUpload`
  - build the zip and metadata locally without uploading to GitLab
- `-GenericPackageName test-data-reviewer`
- `-LatestVersionLabel latest`

Environment variables supported:

- `TDR_GITLAB_PROJECT_URL`
- default target for this project: `https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git`
- `TDR_GITLAB_TOKEN`
- `TDR_GITLAB_PACKAGE_NAME`
- `TDR_GITLAB_LATEST_LABEL`

The script uploads through the GitLab Packages API using your personal access token in the `PRIVATE-TOKEN` request header.

## GitHub Actions automation

The workflow file is:

- `.github/workflows/publish-test-data-reviewer-release.yml`

The workflow is configured to run on a self-hosted Windows runner with these labels:

- `self-hosted`
- `Windows`
- `X64`
- `infineon-intra`

It supports:

- manual publish through `workflow_dispatch`
- automatic publish when a tag matching `tdr-v*` is pushed

Recommended repository secrets:

- `TDR_GITLAB_PROJECT_URL`
  - GitLab project HTTPS URL without credentials
  - for this project: `https://gitlab.intra.infineon.com/wandji/test-data-reviewer.git`
- `TDR_GITLAB_TOKEN`
  - GitLab personal access token with `read_api` and `write_package_registry`

Optional repository variables:

- `TDR_GITLAB_PACKAGE_NAME`
  - recommended value: `test-data-reviewer`
- `TDR_GITLAB_LATEST_LABEL`
  - recommended value: `latest`

## Self-hosted runner setup

Because `gitlab.intra.infineon.com` is only reachable from inside the corporate network, this workflow must run on a self-hosted Windows runner located on an internal machine.

Recommended machine requirements:

- Windows 10 or Windows Server with stable network access to GitHub and GitLab
- local administrator rights for the initial setup
- enough free disk space for Python environments, PyInstaller builds, and temporary artifacts
- outbound access to:
  - `github.com`
  - `api.github.com`
  - `objects.githubusercontent.com`
  - `gitlab.intra.infineon.com`

Recommended runner labels:

- `self-hosted`
- `Windows`
- `X64`
- `infineon-intra`

High-level setup steps:

1. In the GitHub repository, open `Settings > Actions > Runners`
2. Click `New self-hosted runner`
3. Choose `Windows` and `x64`
4. On the internal Windows machine, create a dedicated folder such as `C:\actions-runner\test-data-reviewer`
5. Download the runner package from GitHub onto that machine
6. Extract it into the runner folder
7. Run the GitHub-provided configuration command and add the custom label `infineon-intra`
8. Install the runner as a Windows service
9. Start the service and confirm the runner shows as `Idle` in GitHub

Typical commands on the runner machine look like this:

```powershell
mkdir C:\actions-runner\test-data-reviewer
cd C:\actions-runner\test-data-reviewer

# Download the runner zip from the URL shown in GitHub Settings > Actions > Runners
# Expand-Archive .\actions-runner-win-x64-<version>.zip -DestinationPath .

.\config.cmd --url https://github.com/Nelio92/Tasks_Automation_Code --token <registration-token> --labels infineon-intra
.\svc install
.\svc start
```

Recommended operational practices:

- run the service under a dedicated technical account if your IT rules require it
- keep the runner machine inside the corporate network
- keep Python and PowerShell available on the machine
- periodically update the runner when GitHub announces runner updates
- avoid using the same machine for unrelated interactive work

After the runner is online, start a fresh workflow run from `main`. The job should then be picked up by the internal runner instead of `windows-latest`.

## Package URLs

The uploaded files follow this GitLab API pattern:

```text
https://gitlab.intra.infineon.com/api/v4/projects/wandji%2Ftest-data-reviewer/packages/generic/test-data-reviewer/v1.0.0/TestDataReviewer-v1.0.0.zip
```

And the rolling latest zip is published at:

```text
https://gitlab.intra.infineon.com/api/v4/projects/wandji%2Ftest-data-reviewer/packages/generic/test-data-reviewer/latest/TestDataReviewer-latest.zip
```

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

The workflow strips the `tdr-` prefix and publishes the release as `v1.0.0`.