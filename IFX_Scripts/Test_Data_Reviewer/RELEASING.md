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