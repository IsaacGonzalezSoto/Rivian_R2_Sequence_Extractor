"""
Extraction pipeline that orchestrates the entire process.
Coordinates extractors, validators, and Excel exporter.
"""
import os
import re
from typing import List, Dict, Any
from collections import defaultdict
from ..core.xml_navigator import XMLNavigator
from ..extractors.actuator_extractor import ActuatorExtractor
from ..extractors.transition_extractor import TransitionExtractor
from ..validators.array_validator import ArrayValidator
from ..exporters.excel_exporter import ExcelExporter


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
        
        # Create output folder if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        
        # Initialize XML navigator
        self.navigator = XMLNavigator(l5x_file_path)
        
        # Initialize components
        self.actuator_extractor = ActuatorExtractor(debug=debug)
        self.transition_extractor = TransitionExtractor(debug=debug)
        self.validator = ArrayValidator(debug=debug)
        self.excel_exporter = ExcelExporter()
    
    def run(self) -> List[Dict[str, Any]]:
        """
        Execute the complete extraction pipeline.
        
        Returns:
            List of dictionaries with processed routine information
        """
        if self.debug:
            print("="*60)
            print("COMPLETE EXTRACTOR: SEQUENCES → ACTIONS → ACTUATORS")
            print("="*60)
        
        # Search for sequence routines
        sequence_routines = self.navigator.find_routines_starting_with('EmStatesAndSequences')
        
        if self.debug:
            print(f"\nSequence routines found: {len(sequence_routines)}")
            for routine in sequence_routines:
                routine_name = routine.get('Name', 'Unknown')
                print(f"  → Sequence routine: {routine_name}")
        
        # Process each routine
        all_routines_data = []
        
        for routine in sequence_routines:
            routine_name = routine.get('Name', 'Unknown')
            routine_data = self.process_sequence_routine(routine_name)
            all_routines_data.append(routine_data)
        
        # Final summary
        print(f"\n{'='*60}")
        print(f"✓ PROCESS COMPLETED")
        print(f"{'='*60}")
        print(f"Total routines processed: {len(all_routines_data)}")
        print(f"Files generated in: {self.output_folder}/")
        
        return all_routines_data
    
    def process_sequence_routine(self, routine_name: str) -> Dict[str, Any]:
        """
        Process a complete sequence routine.
        
        Args:
            routine_name: Name of the routine to process
            
        Returns:
            Dictionary with processing information
        """
        print(f"\n{'='*60}")
        print(f"✓ PROCESSING: {routine_name}")
        print(f"{'='*60}")
        
        # Extract actions and sequences with actuators
        sequences_data = self.extract_sequences_with_actuators(routine_name)
        
        # Format sequences for output
        sequences_output = {
            'routine_name': routine_name,
            'sequences': sequences_data
        }
        
        # Extract transitions
        transitions_output = self.extract_transitions(routine_name)
        
        # Export to Excel (one file with multiple sheets)
        excel_filename = f'complete_{routine_name}.xlsx'
        excel_path = os.path.join(self.output_folder, excel_filename)
        
        self.excel_exporter.export(sequences_output, transitions_output, excel_path)
        
        print(f"\n✓ Excel file saved to: {excel_path}")
        print(f"  - Sheet 1: Sequences_Actuators ({len(sequences_data)} sequences)")
        print(f"  - Sheet 2: Transitions ({transitions_output.get('transition_count', 0)} transitions)")
        
        return {
            'routine_name': routine_name,
            'excel_file': excel_path,
            'sequences_count': len(sequences_data),
            'transitions_count': transitions_output.get('transition_count', 0)
        }
    
    def extract_transitions(self, routine_name: str) -> Dict[str, Any]:
        """
        Extract transitions for a routine.
        
        Args:
            routine_name: Name of the routine to process
            
        Returns:
            Dictionary with transition data
        """
        if self.debug:
            print(f"\n--- Extracting Transitions for {routine_name} ---")
        
        # Extract transitions using the extractor
        transitions_data = self.transition_extractor.extract(
            self.navigator.get_root(),
            routine_name
        )
        
        return transitions_data
    
    def extract_sequences_with_actuators(self, routine_name: str) -> List[Dict[str, Any]]:
        """
        Extract sequences with their associated actions and actuators.
        
        Args:
            routine_name: Name of the sequence routine
            
        Returns:
            List of sequences with all information
        """
        # Pattern to extract action assignments
        pattern = r'EmSeqList\[(\d+)\]\.Step\[(\d+)\]\.ActionNumber\[(\d+)\]\s*:=\s*(\w+)\.outActionNum'
        
        # Patterns to extract sequence names
        region_pattern = r'#region\s+Sequence\s+(\d+)\s+-\s+(.+)'
        name_pattern = r"EmSeqList\[(\d+)\]\.Name\s*:=\s*'([^']+)'"
        
        # Structure to store results
        sequences = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
        sequence_names = {}  # Map sequence_index -> descriptive_name
        
        # Get routine
        routine = self.navigator.find_routine_by_name(routine_name)
        if not routine:
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
                    print(f"  Found sequence name from #region: Sequence {seq_idx} - {seq_name}")
            
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
                            print(f"  Found sequence name from hardcoded: Sequence {seq_idx} - {seq_name}")
            
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
                        print(f"\n  Action found: {action_name}")
                        print(f"    MM: {mm_number}, State: {state}")
                        print(f"    Sequence[{seq_idx}].Step[{step_idx}].ActionNumber[{action_idx}]")
                        print(f"    Searching for actuators...")
                    
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
                        print(f"\n  Action without MM pattern: {action_name}")
                    
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
        Parse an action name to extract MM and state.
        
        Args:
            action_name: Action name (e.g., 'ActionMM4Work')
            
        Returns:
            Dictionary with mm_number and state
        """
        pattern = r'Action(MM\d+)(\w+)'
        match = re.match(pattern, action_name)
        
        if match:
            return {
                'mm_number': match.group(1),
                'state': match.group(2)
            }
        return None
    
    def format_sequences(self, sequences: Dict, sequence_names: Dict[int, str] = None) -> List[Dict[str, Any]]:
        """
        Format sequences for structured output.
        
        Args:
            sequences: Nested dictionary with sequences
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
                'sequence_name': sequence_names.get(seq_idx, None),  # New field
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
    
    def extract_sequences_with_actuators(self, routine_name: str) -> List[Dict[str, Any]]:
        """
        Extract sequences with their associated actions and actuators.
        
        Args:
            routine_name: Name of the sequence routine
            
        Returns:
            List of sequences with complete information
        """
        # Pattern to extract action assignments
        pattern = r'EmSeqList\[(\d+)\]\.Step\[(\d+)\]\.ActionNumber\[(\d+)\]\s*:=\s*(\w+)\.outActionNum'
        
        # Patterns to extract sequence names
        region_pattern = r'#region\s+Sequence\s+(\d+)\s+-\s+(.+)'
        name_pattern = r"EmSeqList\[(\d+)\]\.Name\s*:=\s*'([^']+)'"
        
        # Structure to store results
        sequences = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
        sequence_names = {}  # Map sequence_index -> descriptive_name
        
        # Get routine
        routine = self.navigator.find_routine_by_name(routine_name)
        if not routine:
            return []
        
        # Process routine lines
        lines = self.navigator.get_routine_lines(routine)
        
        for line in lines:
            line_text = line.text if line.text else ''
            
            # Extract sequence names from #region comments
            region_match = re.search(region_pattern, line_text)
            if region_match:
                seq_idx = int(region_match.group(1))
                seq_name = region_match.group(2).strip()
                sequence_names[seq_idx] = seq_name
                if self.debug:
                    print(f"  Found sequence: [{seq_idx}] {seq_name}")
            
            # Extract sequence names from hardcoded assignments
            if not region_match:
                name_match = re.search(name_pattern, line_text)
                if name_match:
                    seq_idx = int(name_match.group(1))
                    seq_name = name_match.group(2).strip()
                    if seq_idx not in sequence_names:
                        sequence_names[seq_idx] = seq_name
                        if self.debug:
                            print(f"  Found sequence: [{seq_idx}] {seq_name}")
            
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
                        print(f"\n  Action: {action_name}")
                        print(f"    MM: {mm_number}, State: {state}")
                        print(f"    Seq[{seq_idx}].Step[{step_idx}].Action[{action_idx}]")
                        print(f"    Extracting actuators...")
                    
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
                    # Action without MM pattern
                    if self.debug:
                        print(f"\n  Action (no MM pattern): {action_name}")
                    
                    sequences[seq_idx][step_idx][action_idx] = {
                        'action_name': action_name,
                        'mm_number': None,
                        'state': None,
                        'actuators': [],
                        'validation': None,
                        'full_assignment': match.group(0)
                    }
        
        # Convert to structured list format
        return self.format_sequences(sequences, sequence_names)
    
    def parse_action_name(self, action_name: str) -> Dict[str, str]:
        """
        Parse an action name to extract MM number and state.
        
        Args:
            action_name: Action name (e.g., 'ActionMM4Work')
            
        Returns:
            Dictionary with mm_number and state, or None if pattern doesn't match
        """
        pattern = r'Action(MM\d+)(\w+)'
        match = re.match(pattern, action_name)
        
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