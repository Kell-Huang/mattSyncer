# dataSyncer

A desktop application with graphical interface for updating target files
with data from source files, supporting intelligent column matching and
new column addition.

## Features

- **Intelligent Column Matching**: Automatically matches columns between
  source and target files using prefix stripping and hybrid similarity scoring
- **New Column Addition**: Add source columns that don't exist in the target
  file as new columns with automatic prefix detection
- **SKU-Based Alignment**: Updates data row by row based on SKU matching
- **Formula Preservation**: Detects and preserves Excel formulas during update
- **Robust File Reading**:
  - Handles UTF-8 BOM and various CSV encodings via automatic detection and normalization
  - Prevents type inference errors by treating mapped columns as strings
  - Cleans column names to avoid hidden characters causing matching failures
- **Duplicate Detection**:
  - Detects and warns about duplicate SKUs in source and target files
  - Source duplicates are deduplicated automatically (first occurrence kept)
- **Compact UI Layout**:
  - Top info bar for source/target summary and status
  - Left configuration panel (file, SKU, formula) with compact vertical layout
  - Mapping panel occupies the main workspace
  - Action panel fixed at bottom with right-aligned controls
- **Console Logging**: All log messages are printed to the console for easy monitoring      during execution

## Requirements

- PySide6>=6.5.0
- polars>=0.19.0
- fastexcel>=0.10.0
- openpyxl>=3.1.0
- xlsxwriter>=3.0.0
- chardet>=5.0.0

## Installation

```bash
pip install -r requirements.txt
