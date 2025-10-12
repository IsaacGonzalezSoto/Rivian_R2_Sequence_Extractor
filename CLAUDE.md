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
- **xml_navigator.py**: Utility for navigating L5X XML structure using common XPath patterns. Provides methods to find routines, tags, and extract routine content.
- **base_extractor.py**: Abstract base class implementing Template Method pattern for all extractors. Defines extraction flow: find_items() → validate_items() → format_output()

**src/pipeline/**
- **extraction_pipeline.py**: Main orchestrator that coordinates extractors, validators, and exporters. Processes routines starting with 'EmStatesAndSequences', extracts sequences/actions/actuators, transitions, and digital inputs. Automatically extracts fixture name from L5X filename (e.g., `_010UA1_Fixture_...` → `010UA1`) for output file naming.

**src/extractors/**
- **actuator_extractor.py**: Finds actuator descriptions from MM routines by parsing MOVE statements: `MOVE('DESCRIPTION', MM{X}Cyls[INDEX].Stg.Name)`
- **actuator_group_extractor.py**: Extracts actuator group tags (AOI_Actuator) with tag names (MM1, MM2, etc.) and descriptions ("Group1 Clamps", "Group 4 Pins")
- **transition_extractor.py**: Extracts transition permissions from `EmTransitionStates[X].AutoStartPerms.Y` assignments
- **digital_input_extractor.py**: Extracts all UDT_DigitalInputHal tags across the entire L5X file
- **part_sensor_extractor.py**: Maps digital input sensors to part present detection routines (`Cm{digits}_Part{X}`)

**src/validators/**
- **array_validator.py**: Validates actuator arrays to ensure no missing indices

**src/exporters/**
- **excel_exporter.py**: Exports data to Excel (.xlsx) with multiple sheets using openpyxl. Creates styled sheets with headers and color-coded data.

### Data Flow

1. ExtractionPipeline initializes and extracts fixture name from L5X filename using pattern `_([A-Z0-9]+)_Fixture`
2. XMLNavigator loads L5X file and finds routines starting with 'EmStatesAndSequences'
3. For each routine:
   - Parses regex pattern `EmSeqList[seq][step][action] := ActionMM{X}{State}.outActionNum`
   - Extracts sequence names from `#region Sequence {N} - {Name}` comments
   - For each ActionMM{X} found, ActuatorExtractor finds corresponding MM routine and extracts actuators
   - TransitionExtractor processes transition permissions from the same routine
   - DigitalInputExtractor scans entire L5X for UDT_DigitalInputHal tags
   - PartSensorExtractor identifies Part routines and maps sensors to parts using pattern: `XIC(SENSOR.Out.Value) OTE(Part{X}.inpSensors.Y)`
   - ActuatorGroupExtractor scans entire L5X for AOI_Actuator tags (MM groups)
4. ArrayValidator checks actuator index continuity
5. ExcelExporter creates multi-sheet workbook with filename: `{fixture_name}_{routine_name}.xlsx`

### Key Patterns

**Action Naming**: Actions follow pattern `ActionMM{N}{State}` (e.g., ActionMM4Work, ActionMM12Home)

**MM Routines**: Actuator data is in routines named `Cm{digits}_MM{N}` containing MOVE statements

**Sequence Detection**: Sequences identified by `#region Sequence {index} - {descriptive_name}` or hardcoded `EmSeqList[N].Name := 'Name'`

**Transition Detection**: Transitions use `#region Transition State {index} - {descriptive_name}` with AutoStartPerms assignments

**Fixture Name Extraction**: Fixture name extracted from L5X filename using pattern `_([A-Z0-9]+)_Fixture` (e.g., `_010UA1_Fixture_Em0105_Program.L5X` → `010UA1`). Falls back to `complete` if pattern not found.

**Part Present Detection**: Part routines follow pattern `Cm{digits}_Part{X}` where X is the part number. Sensors are mapped to parts using ladder logic: `XIC(SENSOR_NAME.Out.Value) OTE(Part{X}.inpSensors.Y)`. The system validates that the number of Part routines matches the number of `AOI_Part` tags.

**Known Limitations**: The Part-Sensor extractor expects the simple pattern `XIC(SENSOR.Out.Value) OTE(PartX.inpSensors.Y)` with minimal spacing. If there is complex conditional logic between the XIC and OTE (e.g., `XIC(SENSOR.Out.Value) ,XIC(MM5.stsAtWork) XIO(Spare.Off) ] OTE(PartX.inpSensors.Y)`), the sensor may not be detected. This affects approximately 11% of test cases (1 out of 9 tested files). Example: `_040_UA1_Em0202_Program.L5X` Part3 and Part4.

## Output Structure

Generated Excel files in `output/{L5X_filename}/`:
- **{fixture_name}_{routine_name}.xlsx** (e.g., `010UA1_EmStatesAndSequences_R2.xlsx`) containing:
  - Complete_Flow: Integrated view of sequences and transitions
  - Sequences_Actuators: Detailed sequences with actuators
  - Transitions: Transition permissions table
  - Digital_Inputs: All UDT_DigitalInputHal tags with Program/Tag names/Parent names/Part Assignment (e.g., 'Part1', 'Part2', or 'N/A')
  - Actuator_Groups: All AOI_Actuator tags (MM groups) with Program/Tag Name/Description (e.g., 'MM1' - 'Group1 Clamps', 'MM4' - 'Group 4 Pins')

**Note**: The fixture name is automatically extracted from the L5X filename. For files like `_010UA1_Fixture_Em0105_Program.L5X`, the output will be `010UA1_EmStatesAndSequences_R2.xlsx` instead of the old format `complete_EmStatesAndSequences_R2.xlsx`.

## Dependencies

Core libraries:
- xml.etree.ElementTree: XML parsing
- openpyxl: Excel generation
- typer: CLI framework
- rich: Terminal UI formatting
- pathlib: File operations
- re: Regex pattern matching
