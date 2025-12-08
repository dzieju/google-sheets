# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- Quadra tab now applies column name mapping when rendering table headers
  - Mapping can be configured in `~/.google_sheets_settings.json` using the `quadra_column_names` key
  - Supports both dictionary mapping (selective columns) and list mapping (all columns in order)
  - Case-insensitive matching with whitespace normalization
  - Unmapped columns fall back to default Polish names
  - See README.md section "Użycie w interfejsie Quadra (GUI)" for configuration examples

### Added
- Comprehensive unit tests for Quadra column name mapping (`test_quadra_column_mapping.py`)
- Integration tests verifying end-to-end mapping flow
- Documentation in README.md for Quadra column mapping configuration
