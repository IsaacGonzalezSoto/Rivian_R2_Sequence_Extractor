"""
Extraction pipeline that orchestrates the entire process.
Coordinates extractors, validators, and Excel exporter.
"""
import os
import re
from typing import List, Dict, Any
from collections import defaultdict
from ..core.xml_navigator import XMLNavigator
from ..core.logger import get_logger
from ..core.constants import (
    PATTERN_ACTION_ASSIGNMENT,
    PATTERN_SEQUENCE_REGION,
    PATTERN_SEQUENCE_NAME,
    PATTERN_ACTION_NAME,
    ROUTINE_PREFIX_SEQUENCES,
    DEFAULT_OUTPUT_FOLDER,
    EXCEL_FILE_PREFIX,
    EXCEL_FILE_EXTENSION
)
from ..extractors.actuator_extractor import ActuatorExtractor
from ..extractors.transition_extractor import TransitionExtractor
from ..extractors.digital_input_extractor import DigitalInputExtractor
from ..extractors.part_sensor_extractor import PartSensorExtractor
from ..extractors.actuator_group_extractor import ActuatorGroupExtractor
from ..validators.array_validator import ArrayValidator
from ..exporters.excel_exporter import ExcelExporter

logger = get_logger(__name__)


class ExtractionPipeline:
    """
    Main extraction process orchestrator.
    Coordinates data extraction, validation, and Excel export.
    """
    
    def __init__(self, l5x_file_path: str, output_folder: str = 'output', debug: bool = True):
        """
        Initialize the extraction pipeline.

        Args:
            l5x_file_path: Path to the L5X file
            output_folder: Folder to save Excel results
            debug: Debug mode for detailed logging
        """
        self.l5x_file_path = l5x_file_path
        self.output_folder = output_folder
        self.debug = debug

        # Extract fixture name from L5X filename (used as fallback)
        self.fixture_name = self._extract_fixture_name(l5x_file_path)

        # Create output folder if it doesn't exist
        try:
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
                logger.info(f"Created output folder: {output_folder}")
        except OSError as e:
            logger.error(f"Failed to create output folder: {output_folder} - {str(e)}")
            raise RuntimeError(f"Cannot create output folder: {str(e)}") from e

        # Initialize XML navigator
        self.navigator = XMLNavigator(l5x_file_path)

        # Initialize components
        self.actuator_extractor = ActuatorExtractor(debug=debug)
        self.transition_extractor = TransitionExtractor(debug=debug)
        self.digital_input_extractor = DigitalInputExtractor(debug=debug)
        self.part_sensor_extractor = PartSensorExtractor(debug=debug)
        self.actuator_group_extractor = ActuatorGroupExtractor(debug=debug)
        self.validator = ArrayValidator(debug=debug)
        self.excel_exporter = ExcelExporter()

    def _extract_fixture_name(self, l5x_file_path: str) -> str:
        """
        Extract fixture name from L5X filename.

        Examples:
            _010_UA1_Em0106_Program.L5X -> 010_UA1_Em0106
            _010UA1_Fixture_Em0105.L5X -> 010UA1_Fixture_Em0105
            MyFixture_Program.L5X -> MyFixture

        Args:
            l5x_file_path: Path to the L5X file

        Returns:
            Fixture name extracted from filename
        """
        import os
        filename = os.path.basename(l5x_file_path)

        # Remove .L5X extension (case insensitive)
        base_name = re.sub(r'\.L5X$', '', filename, flags=re.IGNORECASE)

        # Remove common suffixes like _Program, _Fixture, etc.
        base_name = re.sub(r'_Program$', '', base_name, flags=re.IGNORECASE)

        # Remove leading underscore if present
        if base_name.startswith('_'):
            base_name = base_name[1:]

        # If we got an empty string somehow, use fallback
        if not base_name:
            base_name = EXCEL_FILE_PREFIX.rstrip('_')

        if self.debug:
            logger.debug(f"Extracted fixture name: {base_name} from {filename}")

        return base_name

    def run(self) -> List[Dict[str, Any]]:
        """
        Execute the complete extraction pipeline.

        Detects fixture programs in the L5X file:
        - Single-fixture files: Process as before (backward compatible)
        - Multi-fixture files: Create subfolder per fixture

        Returns:
            List of dictionaries with processed routine information
        """
        if self.debug:
            logger.info("="*60)
            logger.info("COMPLETE EXTRACTOR: SEQUENCES → ACTIONS → ACTUATORS")
            logger.info("="*60)

        # Identify fixture programs
        fixture_programs = self.navigator.find_fixture_programs()

        if not fixture_programs:
            logger.warning("No fixture programs found in L5X file")
            return []

        logger.info(f"Identified {len(fixture_programs)} fixture program(s)")

        all_routines_data = []

        # Determine if we need subfolders (multi-fixture scenario)
        multi_fixture_mode = len(fixture_programs) > 1

        for fixture_info in fixture_programs:
            program_name = fixture_info['program_name']
            em_routines = fixture_info['em_routines']

            logger.info("="*60)
            logger.info(f"PROCESSING FIXTURE: {program_name}")
            logger.info(f"EmStatesAndSequences routines: {len(em_routines)}")
            logger.info("="*60)

            # Determine output folder for this fixture
            if multi_fixture_mode:
                # Create subfolder for this fixture
                fixture_output_folder = os.path.join(self.output_folder, program_name)
                try:
                    if not os.path.exists(fixture_output_folder):
                        os.makedirs(fixture_output_folder)
                        logger.info(f"Created fixture subfolder: {fixture_output_folder}")
                except OSError as e:
                    logger.error(f"Failed to create fixture subfolder: {fixture_output_folder} - {str(e)}")
                    continue
            else:
                # Single fixture: use base output folder (backward compatible)
                fixture_output_folder = self.output_folder

            # Process each EmStatesAndSequences routine in this fixture
            for routine_name in em_routines:
                logger.info(f"  → Processing routine: {routine_name}")
                routine_data = self.process_sequence_routine(
                    routine_name=routine_name,
                    program_name=program_name,
                    fixture_output_folder=fixture_output_folder
                )
                all_routines_data.append(routine_data)

        # Final summary
        logger.info("="*60)
        logger.info("✓ PROCESS COMPLETED")
        logger.info("="*60)
        logger.info(f"Total fixtures processed: {len(fixture_programs)}")
        logger.info(f"Total routines processed: {len(all_routines_data)}")
        logger.info(f"Files generated in: {self.output_folder}/")

        return all_routines_data
    
    def process_sequence_routine(self, routine_name: str, program_name: str = None, fixture_output_folder: str = None) -> Dict[str, Any]:
        """
        Process a complete sequence routine.

        Args:
            routine_name: Name of the routine to process
            program_name: Name of the program containing the routine (optional, for multi-fixture support)
            fixture_output_folder: Output folder for this specific fixture (optional)

        Returns:
            Dictionary with processing information
        """
        logger.info("="*60)
        logger.info(f"✓ PROCESSING: {routine_name}")
        if program_name:
            logger.info(f"  Program: {program_name}")
        logger.info("="*60)

        # Use provided output folder or fall back to default
        output_folder = fixture_output_folder if fixture_output_folder else self.output_folder

        # Extract actions and sequences with actuators
        sequences_data = self.extract_sequences_with_actuators(routine_name, program_name)

        # Format sequences for output
        sequences_output = {
            'routine_name': routine_name,
            'sequences': sequences_data
        }

        # Extract transitions
        transitions_output = self.extract_transitions(routine_name, program_name)

        # Extract digital inputs (scoped to program if multi-fixture)
        digital_inputs_output = self.extract_digital_inputs(program_name)

        # Extract actuator groups (scoped to program if multi-fixture)
        actuator_groups_output = self.extract_actuator_groups(program_name)

        # Export to Excel (one file with multiple sheets)
        # Extract fixture name from program_name if provided
        if program_name:
            # Extract fixture identifier from program name
            # Example: _010UA1_Fixture_Em0105 -> 010UA1_Fixture_Em0105
            fixture_name = program_name.lstrip('_')
        else:
            fixture_name = self.fixture_name

        excel_filename = f'{fixture_name}_{routine_name}{EXCEL_FILE_EXTENSION}'
        excel_path = os.path.join(output_folder, excel_filename)

        try:
            self.excel_exporter.export(sequences_output, transitions_output, digital_inputs_output, actuator_groups_output, excel_path)
            logger.info(f"✓ Excel file saved to: {excel_path}")
            logger.info(f"  - Sheet 1: Sequences_Actuators ({len(sequences_data)} sequences, {actuator_groups_output.get('group_count', 0)} MM groups)")
            logger.info(f"  - Sheet 2: Transitions ({transitions_output.get('transition_count', 0)} transitions)")
            logger.info(f"  - Sheet 3: Digital Inputs ({digital_inputs_output.get('input_count', 0)} tags)")
        except Exception as e:
            logger.error(f"Failed to export Excel file: {excel_path} - {str(e)}")
            raise RuntimeError(f"Excel export failed: {str(e)}") from e

        return {
            'routine_name': routine_name,
            'program_name': program_name,
            'excel_file': excel_path,
            'sequences_count': len(sequences_data),
            'transitions_count': transitions_output.get('transition_count', 0),
            'digital_inputs_count': digital_inputs_output.get('input_count', 0),
            'actuator_groups_count': actuator_groups_output.get('group_count', 0)
        }
    
    def extract_transitions(self, routine_name: str, program_name: str = None) -> Dict[str, Any]:
        """
        Extract transitions for a routine.

        Args:
            routine_name: Name of the routine to process
            program_name: Optional program name for scoping (multi-fixture support)

        Returns:
            Dictionary with transition data
        """
        if self.debug:
            logger.debug(f"Extracting Transitions for {routine_name}")
            if program_name:
                logger.debug(f"  Scoped to program: {program_name}")

        # Extract transitions using the extractor
        transitions_data = self.transition_extractor.extract(
            self.navigator.get_root(),
            routine_name,
            program_name=program_name
        )

        return transitions_data
    
    def extract_digital_inputs(self, program_name: str = None) -> Dict[str, Any]:
        """
        Extract digital input tags from all programs and map them to parts.
        This is independent from sequences and transitions.

        Args:
            program_name: Optional program name to scope extraction (multi-fixture support)

        Returns:
            Dictionary with digital input data
        """
        if self.debug:
            logger.debug("Extracting Digital Inputs (UDT_DigitalInputHal)")
            if program_name:
                logger.debug(f"  Scoped to program: {program_name}")

        # Get the appropriate root element (scoped to program if provided)
        if program_name:
            program_element = self.navigator.find_program_by_name(program_name)
            if not program_element:
                logger.warning(f"Program not found: {program_name}, using full L5X")
                search_root = self.navigator.get_root()
            else:
                search_root = program_element
        else:
            search_root = self.navigator.get_root()

        # Extract all digital inputs from the search scope
        digital_inputs = self.digital_input_extractor.extract_all_digital_inputs(
            search_root,
            program_name=program_name
        )

        # Extract sensor-to-part mappings
        if self.debug:
            logger.debug("Extracting Part-Sensor relationships")

        sensor_to_parts = self.part_sensor_extractor.extract_all_part_sensors(
            search_root,
            program_name=program_name
        )

        # Update digital inputs with part assignments
        digital_inputs = self.digital_input_extractor.update_part_assignments(
            digital_inputs,
            sensor_to_parts
        )

        # Format output
        return self.digital_input_extractor.format_output(digital_inputs)

    def extract_actuator_groups(self, program_name: str = None) -> Dict[str, Any]:
        """
        Extract actuator group tags from all programs.
        This is independent from sequences and transitions.

        Args:
            program_name: Optional program name to scope extraction (multi-fixture support)

        Returns:
            Dictionary with actuator group data
        """
        if self.debug:
            logger.debug("Extracting Actuator Groups (AOI_Actuator)")
            if program_name:
                logger.debug(f"  Scoped to program: {program_name}")

        # Get the appropriate root element (scoped to program if provided)
        if program_name:
            program_element = self.navigator.find_program_by_name(program_name)
            if not program_element:
                logger.warning(f"Program not found: {program_name}, using full L5X")
                search_root = self.navigator.get_root()
            else:
                search_root = program_element
        else:
            search_root = self.navigator.get_root()

        # Extract all actuator groups from the search scope
        actuator_groups = self.actuator_group_extractor.extract_all_actuator_groups(
            search_root,
            program_name=program_name
        )

        # Format output
        return self.actuator_group_extractor.format_output(actuator_groups)
    
    def extract_sequences_with_actuators(self, routine_name: str, program_name: str = None) -> List[Dict[str, Any]]:
        """
        Extract sequences with their associated actions and actuators.

        Args:
            routine_name: Name of the sequence routine
            program_name: Optional program name for scoping (multi-fixture support)

        Returns:
            List of sequences with all information
        """
        # Pattern to extract action assignments
        pattern = PATTERN_ACTION_ASSIGNMENT

        # Patterns to extract sequence names
        region_pattern = PATTERN_SEQUENCE_REGION
        name_pattern = PATTERN_SEQUENCE_NAME

        # Structure to store results
        sequences = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
        sequence_names = {}  # Map sequence_index -> descriptive_name

        # Get routine (scoped to program if provided)
        if program_name:
            # Find routine within the specific program
            program_element = self.navigator.find_program_by_name(program_name)
            if program_element:
                routines = program_element.findall(f'.//Routines/Routine[@Name="{routine_name}"]')
                routine = routines[0] if routines else None
            else:
                logger.warning(f"Program not found: {program_name}, searching globally")
                routine = self.navigator.find_routine_by_name(routine_name)
        else:
            routine = self.navigator.find_routine_by_name(routine_name)

        if not routine:
            logger.warning(f"Routine not found: {routine_name}")
            return []
        
        # Process routine lines
        lines = self.navigator.get_routine_lines(routine)
        
        for line in lines:
            line_text = line.text if line.text else ''
            
            # Check for #region comments to extract sequence names
            region_match = re.search(region_pattern, line_text)
            if region_match:
                seq_idx = int(region_match.group(1))
                seq_name = region_match.group(2).strip()
                sequence_names[seq_idx] = seq_name
                if self.debug:
                    logger.debug(f"Found sequence name from #region: Sequence {seq_idx} - {seq_name}")
            
            # If not found in #region, check hardcoded Name
            if not region_match:
                name_match = re.search(name_pattern, line_text)
                if name_match:
                    seq_idx = int(name_match.group(1))
                    seq_name = name_match.group(2).strip()
                    # Only store if we haven't found it in #region
                    if seq_idx not in sequence_names:
                        sequence_names[seq_idx] = seq_name
                        if self.debug:
                            logger.debug(f"Found sequence name from hardcoded: Sequence {seq_idx} - {seq_name}")
            
            # Search for action assignments
            matches = re.finditer(pattern, line_text)
            
            for match in matches:
                seq_idx = int(match.group(1))
                step_idx = int(match.group(2))
                action_idx = int(match.group(3))
                action_name = match.group(4)
                
                # Parse action name
                action_info = self.parse_action_name(action_name)
                
                if action_info:
                    mm_number = action_info['mm_number']
                    state = action_info['state']
                    
                    if self.debug:
                        logger.debug(f"Action found: {action_name}")
                        logger.debug(f"  MM: {mm_number}, State: {state}")
                        logger.debug(f"  Sequence[{seq_idx}].Step[{step_idx}].ActionNumber[{action_idx}]")
                        logger.debug(f"  Searching for actuators...")
                    
                    # Extract actuators
                    actuators = self.actuator_extractor.find_actuators_for_mm(
                        self.navigator.get_root(), 
                        mm_number
                    )
                    
                    # Validate actuators
                    validation = self.validator.validate_actuators(
                        self.navigator.get_root(),
                        mm_number,
                        actuators
                    )
                    
                    sequences[seq_idx][step_idx][action_idx] = {
                        'action_name': action_name,
                        'mm_number': mm_number,
                        'state': state,
                        'actuators': actuators,
                        'validation': validation,
                        'full_assignment': match.group(0)
                    }
                else:
                    # Action that doesn't follow the pattern
                    if self.debug:
                        logger.debug(f"Action without MM pattern: {action_name}")
                    
                    sequences[seq_idx][step_idx][action_idx] = {
                        'action_name': action_name,
                        'mm_number': None,
                        'state': None,
                        'actuators': [],
                        'validation': None,
                        'full_assignment': match.group(0)
                    }
        
        # Convert to list structured format and add sequence names
        return self.format_sequences(sequences, sequence_names)

    def parse_action_name(self, action_name: str) -> Dict[str, str]:
        """
        Parse an action name to extract MM number and state.

        Args:
            action_name: Action name (e.g., 'ActionMM4Work')

        Returns:
            Dictionary with mm_number and state, or None if pattern doesn't match
        """
        match = re.match(PATTERN_ACTION_NAME, action_name)

        if match:
            return {
                'mm_number': match.group(1),
                'state': match.group(2)
            }
        return None

    def format_sequences(self, sequences: Dict, sequence_names: Dict[int, str] = None) -> List[Dict[str, Any]]:
        """
        Format sequences into structured output.

        Args:
            sequences: Nested dictionary with sequence data
            sequence_names: Optional dictionary mapping sequence_index to descriptive_name

        Returns:
            List of formatted sequences
        """
        if sequence_names is None:
            sequence_names = {}

        formatted = []

        for seq_idx in sorted(sequences.keys()):
            sequence = {
                'sequence_index': seq_idx,
                'sequence_name': sequence_names.get(seq_idx, None),
                'steps': []
            }

            for step_idx in sorted(sequences[seq_idx].keys()):
                step = {
                    'step_index': step_idx,
                    'actions': []
                }

                for action_idx in sorted(sequences[seq_idx][step_idx].keys()):
                    action_data = sequences[seq_idx][step_idx][action_idx]

                    action = {
                        'action_index': action_idx,
                        'action_name': action_data['action_name'],
                        'mm_number': action_data['mm_number'],
                        'state': action_data['state'],
                        'actuator_count': len(action_data['actuators']),
                        'actuators': action_data['actuators']
                    }

                    if action_data['validation']:
                        action['validation'] = action_data['validation']

                    step['actions'].append(action)

                sequence['steps'].append(step)

            formatted.append(sequence)

        return formatted