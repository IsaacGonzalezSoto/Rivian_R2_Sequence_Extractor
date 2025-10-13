# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

R2_Sequence_Extractor is a Python application that parses Rockwell Automation L5X (RSLogix 5000) XML files to extract industrial automation sequences, actions, actuators, transitions, and digital inputs. It generates Excel reports with multiple sheets for comprehensive data analysis.

## Running the Application

```bash
python main.py
```

The application runs interactively:
- Automatically discovers L5X files in the current directory
- Prompts user to select: all files, one file, or multiple files
- Processes selected files and generates Excel output in `output/` directory

## Architecture

### Entry Point
- **main.py**: Interactive CLI that discovers L5X files, handles user file selection, and orchestrates batch processing

### Core Components

**src/core/**
- **xml_navigator.py**: Utility for navigating L5X XML structure using common XPath patterns. Provides methods to find routines, tags, programs, and extract routine content. Includes fixture program identification logic using pattern matching.
- **base_extractor.py**: Abstract base class implementing Template Method pattern for all extractors. Defines extraction flow: find_items() → validate_items() → format_output(). Supports program-scoped extraction for multi-fixture files.

**src/pipeline/**
- **extraction_pipeline.py**: Main orchestrator that coordinates extractors, validators, and exporters. Automatically detects fixture programs (single or multiple) in L5X files using pattern `_\d{3}UA\d_` or "Fixture" keyword. Processes routines starting with 'EmStatesAndSequences', extracts sequences/actions/actuators, transitions, and digital inputs. For multi-fixture files, creates subfolder per fixture; for single-fixture files, maintains backward compatibility with flat structure.

**src/extractors/**
- **actuator_extractor.py**: Finds actuator descriptions from MM routines by parsing MOVE statements: `MOVE('DESCRIPTION', MM{X}Cyls[INDEX].Stg.Name)`
- **actuator_group_extractor.py**: Extracts actuator group tags (AOI_Actuator) with tag names (MM1, MM2, etc.) and descriptions ("Group1 Clamps", "Group 4 Pins")
- **transition_extractor.py**: Extracts transition permissions from `EmTransitionStates[X].AutoStartPerms.Y` assignments
- **digital_input_extractor.py**: Extracts all UDT_DigitalInputHal tags across the entire L5X file
- **part_sensor_extractor.py**: Maps digital input sensors to part present detection routines (`Cm{digits}_Part{X}`)
- **valve_mapping_extractor.py**: Maps MM groups to physical valve manifolds and positions by extracting MM commands from fixture programs and parsing AOI_ValveManifold_V* calls (V4, V8, V12, V16) in MapIo program (multi-fixture files only)

**src/validators/**
- **array_validator.py**: Validates actuator arrays to ensure no missing indices

**src/exporters/**
- **excel_exporter.py**: Exports data to Excel (.xlsx) with multiple sheets using openpyxl. Creates styled sheets with headers and color-coded data.

### Data Flow

1. ExtractionPipeline initializes and loads L5X file
2. XMLNavigator identifies fixture programs using pattern `_\d{3}UA\d_` or "Fixture" keyword, validated by presence of EmStatesAndSequences routines
3. For each fixture program:
   - Determines output folder: subfolder for multi-fixture files, base folder for single-fixture (backward compatible)
   - For each EmStatesAndSequences routine in the program:
     - Parses regex pattern `EmSeqList[seq][step][action] := ActionMM{X}{State}.outActionNum`
     - Extracts sequence names from `#region Sequence {N} - {Name}` comments
     - For each ActionMM{X} found, ActuatorExtractor finds corresponding MM routine and extracts actuators
     - TransitionExtractor processes transition permissions from the same routine
     - DigitalInputExtractor scans program-scoped for UDT_DigitalInputHal tags
     - PartSensorExtractor identifies Part routines and maps sensors to parts using pattern: `XIC(SENSOR.Out.Value) OTE(Part{X}.inpSensors.Y)`
     - ActuatorGroupExtractor scans program-scoped for AOI_Actuator tags (MM groups)
     - ValveMappingExtractor extracts valve mappings from MapIo (multi-fixture only): extracts MM command names from fixture MM routines, finds AOI_ValveManifold_V* calls (V4, V8, V12, V16) in MapIo, matches commands to valve positions
4. ArrayValidator checks actuator index continuity
5. ExcelExporter creates multi-sheet workbook with filename: `{fixture_name}_{routine_name}.xlsx`

### Key Patterns

**Action Naming**: Actions follow pattern `ActionMM{N}{State}` (e.g., ActionMM4Work, ActionMM12Home)

**MM Routines**: Actuator data is in routines named `Cm{digits}_MM{N}` containing MOVE statements

**Sequence Detection**: Sequences identified by `#region Sequence {index} - {descriptive_name}` or hardcoded `EmSeqList[N].Name := 'Name'`

**Transition Detection**: Transitions use `#region Transition State {index} - {descriptive_name}` with AutoStartPerms assignments

**Fixture Identification**: Fixtures identified by:
1. Primary pattern: `_\d{3}UA\d_` in program name (e.g., `_010UA1_Fixture_Em0105`)
2. Secondary pattern: "Fixture" keyword in program name
3. Validation: Must contain at least one `EmStatesAndSequences` routine
Supports both single-fixture L5X files (e.g., `_010UA1_Fixture_Em0105_Program.L5X`) and multi-fixture L5X files (e.g., `BL03FFLR_PLC01.L5X` containing multiple fixture programs).

**Part Present Detection**: Part routines follow pattern `Cm{digits}_Part{X}` where X is the part number. Sensors are mapped to parts using ladder logic: `XIC(SENSOR_NAME.Out.Value) OTE(Part{X}.inpSensors.Y)`. The system validates that the number of Part routines matches the number of `AOI_Part` tags.

**Valve Mapping (Multi-Fixture Only)**: For multi-fixture files with MapIo program:
1. Supports AOI_ValveManifold_V4, V8, V12, and V16 (generic version detection using pattern `AOI_ValveManifold_V\d+`)
2. Extracts MM command names from fixture program's MM routines (pattern: `,XIC(MM{N}.outWork) OTE(MM{N}_ToWork.Inp.Value`)
3. Finds MapIo program and searches for AOI_ValveManifold_V* calls that reference the fixture (handles backslash escapes)
4. Parses AOI parameters:
   - Parameter 3: Manifold name (e.g., `_010UA1KJ1_KEB1_Hw`)
   - Parameters 6-7: MM1 Work/Home commands
   - Parameters 8-9: MM2 Work/Home commands (continues in pairs)
5. Calculates valve positions: valve_index = ((param_position - 6) // 2) + 1, valve position format: "{valve_index}A" (Work) or "{valve_index}B" (Home)
6. Handles monoestable valves (Spare.DO entries show as "N/A")
7. Single-fixture files without MapIo: valve mapping columns exist but remain empty (backward compatible)

**Known Limitations**: The Part-Sensor extractor expects the simple pattern `XIC(SENSOR.Out.Value) OTE(PartX.inpSensors.Y)` with minimal spacing. If there is complex conditional logic between the XIC and OTE (e.g., `XIC(SENSOR.Out.Value) ,XIC(MM5.stsAtWork) XIO(Spare.Off) ] OTE(PartX.inpSensors.Y)`), the sensor may not be detected. This affects approximately 11% of test cases (1 out of 9 tested files). Example: `_040_UA1_Em0202_Program.L5X` Part3 and Part4.

## Output Structure

### Single-Fixture Files
Generated Excel files in `output/{L5X_filename}/`:
```
output/_010UA1_Fixture_Em0105_Program/
├── 010UA1_Fixture_Em0105_EmStatesAndSequences_Common.xlsx
└── 010UA1_Fixture_Em0105_EmStatesAndSequences_R2S.xlsx
```

### Multi-Fixture Files
Generated Excel files in `output/{L5X_filename}/{program_name}/`:
```
output/BL03FFLR_PLC01/
├── _010UA1_Fixture_Em0105/
│   ├── 010UA1_Fixture_Em0105_EmStatesAndSequences_Common.xlsx
│   └── 010UA1_Fixture_Em0105_EmStatesAndSequences_R2S.xlsx
├── _020UA1_Fixture_Em0207/
│   ├── 020UA1_Fixture_Em0207_EmStatesAndSequences_Common.xlsx
│   └── 020UA1_Fixture_Em0207_EmStatesAndSequences_R2S.xlsx
├── _030UA1_Fixture_Em0301/
│   └── ...
└── _040UA1_Fixture_Em0302/
    └── ...
```

### Excel File Structure
Each Excel file contains:
- **Complete_Flow**: Integrated view of sequences and transitions
- **Sequences_Actuators**: Detailed sequences with actuators, MM group descriptions, and valve mappings
  - Column G: MM_Group_Description (e.g., 'Group1 Clamps', 'Group 4 Pins')
  - Column H: Manifold (e.g., '_010UA1KJ1_KEB1_Hw') - Multi-fixture only
  - Column I: Valve_Work (e.g., '1A', '2A') - Multi-fixture only
  - Column J: Valve_Home (e.g., '1B', '2B') - Multi-fixture only
- **Transitions**: Transition permissions table
- **Digital_Inputs**: All UDT_DigitalInputHal tags with Program/Tag names/Parent names/Part Assignment (e.g., 'Part1', 'Part2', or 'N/A')

## Dependencies

Core libraries:
- xml.etree.ElementTree: XML parsing
- openpyxl: Excel generation
- typer: CLI framework
- rich: Terminal UI formatting
- pathlib: File operations
- re: Regex pattern matching
