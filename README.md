# R2 Sequence Extractor

A Python application that parses Rockwell Automation L5X (RSLogix 5000) XML files to extract industrial automation sequences, actions, actuators, transitions, and digital inputs. It generates comprehensive Excel reports for data analysis.

## Features

- **Interactive File Selection**: Automatically discovers L5X files and provides interactive selection
- **Sequence Extraction**: Parses automation sequences with actions and actuators
- **Transition Analysis**: Extracts transition permissions and states
- **Digital Input Mapping**: Identifies all UDT_DigitalInputHal tags across the project
- **Excel Reports**: Generates multi-sheet Excel workbooks with:
  - Complete Flow (integrated sequences and transitions)
  - Sequences & Actuators (detailed breakdown)
  - Transitions (permission tables)
  - Digital Inputs (tag mappings)

## Requirements

- Python 3.7+
- Dependencies:
  - openpyxl (Excel generation)
  - Standard library: xml.etree.ElementTree, pathlib, re

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd R2_Sequence_Extractor
```

2. Install dependencies:
```bash
pip install openpyxl
```

## Usage

Run the application:
```bash
python main.py
```

The interactive CLI will:
1. Automatically discover all L5X files in the current directory
2. Prompt you to select files (all, single, or multiple)
3. Process selected files
4. Generate Excel reports in the `output/` directory

### Output Location

Excel files are saved to:
```
output/{L5X_filename}/{fixture_name}_{routine_name}.xlsx
```

**Fixture Name Extraction:**
- The application automatically extracts the fixture name from the L5X filename
- Example: `_010UA1_Fixture_Em0105_Program.L5X` → fixture name is `010UA1`
- Output file: `010UA1_EmStatesAndSequences_R2.xlsx`
- If the pattern is not found, it defaults to `complete_` prefix

## Architecture

### Project Structure

```
R2_Sequence_Extractor/
├── main.py                          # Entry point with interactive CLI
├── src/
│   ├── core/
│   │   ├── xml_navigator.py         # XML parsing utilities
│   │   └── base_extractor.py        # Abstract base for extractors
│   ├── pipeline/
│   │   └── extraction_pipeline.py   # Main orchestrator
│   ├── extractors/
│   │   ├── actuator_extractor.py    # Actuator extraction
│   │   ├── transition_extractor.py  # Transition extraction
│   │   └── digital_input_extractor.py # Digital input extraction
│   ├── validators/
│   │   └── array_validator.py       # Data validation
│   └── exporters/
│       └── excel_exporter.py        # Excel report generation
├── output/                          # Generated reports
└── CLAUDE.md                        # Developer guidance
```

### How It Works

1. **Discovery**: XMLNavigator loads L5X files and finds routines starting with `EmStatesAndSequences`
2. **Extraction**:
   - Parses action patterns: `EmSeqList[seq][step][action] := ActionMM{X}{State}.outActionNum`
   - Extracts sequence names from `#region Sequence {N} - {Name}` comments
   - Finds actuator descriptions in corresponding `Cm{digits}_MM{N}` routines
   - Processes transition permissions from `EmTransitionStates[X].AutoStartPerms.Y`
   - Scans entire file for UDT_DigitalInputHal tags
3. **Validation**: Ensures actuator array continuity
4. **Export**: Generates styled Excel workbook with multiple sheets

### Key Patterns

- **Actions**: Follow pattern `ActionMM{N}{State}` (e.g., `ActionMM4Work`, `ActionMM12Home`)
- **MM Routines**: Actuator data in routines named `Cm{digits}_MM{N}` containing MOVE statements
- **Sequences**: Identified by `#region Sequence {index} - {descriptive_name}` comments
- **Transitions**: Use `#region Transition State {index} - {descriptive_name}` with AutoStartPerms

## Output Format

Each Excel workbook contains multiple sheets:

### Complete_Flow
Integrated view combining sequences and transitions

### Sequences_Actuators
Detailed sequences with:
- Sequence Name
- Step Number
- Action Type
- Actuator Index
- Actuator Description

### Transitions
Transition permissions table with:
- Transition Index
- Transition Name
- Permission Details

### Digital_Inputs
All UDT_DigitalInputHal tags with:
- Program Name
- Tag Name
- Full Path

## Building Standalone Executable

You can create a standalone Windows executable (.exe) that doesn't require Python installation.

### Prerequisites

- Python 3.7+
- PyInstaller (will be installed automatically by the build script)

### Build Steps

1. Run the build script:
```bash
build_exe.bat
```

The script will:
- Check Python installation
- Install/verify PyInstaller
- Install all required dependencies
- Build the executable using PyInstaller
- Place the result in `dist/R2_Sequence_Extractor.exe`

### Using the Executable

Once built, the `.exe` file is completely standalone:
1. Copy `R2_Sequence_Extractor.exe` to any Windows PC
2. Place your L5X files in the same folder
3. Double-click the `.exe` to run
4. No Python installation required on the target machine

### Manual Build

If you prefer to build manually:
```bash
pip install pyinstaller
pyinstaller main.spec
```

The executable will be in `dist/R2_Sequence_Extractor.exe`.

## Development

### Adding New Extractors

1. Inherit from `BaseExtractor` in [src/core/base_extractor.py](src/core/base_extractor.py)
2. Implement required methods:
   - `find_items()`: Locate data in XML
   - `validate_items()`: Validate extracted data
   - `format_output()`: Format for export
3. Register in [src/pipeline/extraction_pipeline.py](src/pipeline/extraction_pipeline.py)

### Logging

Centralized logging is configured in [src/core/logger.py](src/core/logger.py). Logs are written to `app.log`.

## Recent Changes

- **Fixture Name in Output Files**: Excel filenames now include the fixture name extracted from L5X filename (e.g., `010UA1_EmStatesAndSequences_R2.xlsx`)
- Added centralized logging and constants
- Implemented Digital Inputs extractor for UDT_DigitalInputHal tags
- Added interactive file selection with Typer and Rich
- Improved Excel readability with styling
- Removed CSV/JSON exporters (Excel-only output)

## License

[Add your license here]

## Contributing

[Add contributing guidelines here]
