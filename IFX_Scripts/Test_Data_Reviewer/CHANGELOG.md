# Changelog

All notable changes to Test Data Reviewer should be documented in this file.

## v1.0.2 - 2026-03-26

### Changed
- Removed `output_folder` from the user YAML configuration.
- The tool now creates and uses `Outputs` automatically under the configured `input_folder`.
- Updated the default packaged configuration to use the placeholder path `your_input_folder`.

### Packaging
- Rebuilt the executable and refreshed the packaged release artifacts for the updated user-facing configuration.

## v1.0.1 - 2026-03-26

### Changed
- Updated CDF plot y-axis scaling to a probability-style percentage view.
- Added CDF y-axis levels at 0.01, 0.1, 1, 10, 50, 90, 99, 99.9, and 99.99.
- Improved consistency of recorded CDF plots with the existing data visualization tool.

### Packaging
- Rebuilt the executable and refreshed the packaged release artifacts.

## v1.0.0 - 2026-03-25

### Added
- First official team release of Test Data Reviewer.
- Packaged executable delivery for team usage through the team-facing GitLab repository.
- Automated review workbook generation with embedded plot-sheet navigation and CDF plots.
- Support for configurable release publishing from the private source repository to the team distribution repository.

### Included
- Review metrics for fails, Cpk thresholds, site-to-site delta, unique values, skewness, and multimodality.
- Release packaging via PyInstaller with a ready-to-use executable and configuration files.