"""
Extraction pipeline that orchestrates the entire process.
Coordinates extractors, validators, and exporters.
"""
import os
import re
from typing import List, Dict, Any
from collections import defaultdict
from ..core.xml_navigator import XMLNavigator
from ..extractors.actuator_extractor import ActuatorExtractor
from ..extractors.transition_extractor import TransitionExtractor
from ..validators.array_validator import ArrayValidator
from ..exporters.json_exporter import JSONExporter
from ..exporters.csv_exporter import CSVExporter
from ..exporters.transition_csv_exporter import TransitionCSVExporter


class ExtractionPipeline:
    """
    Main extraction process orchestrator.
    Coordinates data extraction, validation, and export.
    """
    
    def __init__(self, l5x_file_path: str, output_folder: str = 'output', debug: bool = True):
        """
        Initialize the extraction pipeline.
        
        Args:
            l5x_file_path: Path to the L5X file
            output_folder: Folder to save results
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
        self.json_exporter = JSONExporter()
        self.csv_exporter = CSVExporter()
        self.transition_csv_exporter = TransitionCSVExporter()
    
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
        
        # Format for output
        output_data = {
            'routine_name': routine_name,
            'sequences': sequences_data
        }
        
        # Export sequences to JSON
        json_filename = f'complete_{routine_name}.json'
        json_path = os.path.join(self.output_folder, json_filename)
        self.json_exporter.export(output_data, json_path)
        
        # Export sequences to CSV
        csv_filename = f'complete_{routine_name}.csv'
        csv_path = os.path.join(self.output_folder, csv_filename)
        self.csv_exporter.export(output_data, csv_path)
        
        print(f"\n✓ Sequences JSON saved to: {json_path}")
        print(f"✓ Sequences CSV saved to: {csv_path}")
        print(f"✓ Sequences processed: {len(sequences_data)}")
        
        # Process transitions
        transitions_result = self.process_transitions(routine_name)
        
        return {
            'routine_name': routine_name,
            'sequences_file': json_path,
            'sequences_csv': csv_path,
            'sequences_count': len(sequences_data),
            'transitions_file': transitions_result.get('json_path'),
            'transitions_csv': transitions_result.get('csv_path'),
            'transitions_count': transitions_result.get('count', 0)
        }
    
    def process_transitions(self, routine_name: str) -> Dict[str, Any]:
        """
        Process transitions for a routine.
        
        Args:
            routine_name: Name of the routine to process
            
        Returns:
            Dictionary with transition processing information
        """
        print(f"\n--- Processing Transitions for {routine_name} ---")
        
        # Extract transitions
        transitions_data = self.transition_extractor.extract(
            self.navigator.get_root(),
            routine_name
        )
        
        if not transitions_data['transitions']:
            print("  No transitions found")
            return {'count': 0}
        
        # Export transitions to JSON
        json_filename = f'transitions_{routine_name}.json'
        json_path = os.path.join(self.output_folder, json_filename)
        self.json_exporter.export(transitions_data, json_path)
        
        # Export transitions to CSV
        csv_filename = f'transitions_{routine_name}.csv'
        csv_path = os.path.join(self.output_folder, csv_filename)
        self.transition_csv_exporter.export(transitions_data, csv_path)
        
        print(f"✓ Transitions JSON saved to: {json_path}")
        print(f"✓ Transitions CSV saved to: {csv_path}")
        print(f"✓ Transitions processed: {transitions_data['transition_count']}")
        
        return {
            'json_path': json_path,
            'csv_path': csv_path,
            'count': transitions_data['transition_count']
        }
    
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
        
        # Structure to store results
        sequences = defaultdict(lambda: defaultdict(lambda: defaultdict(dict)))
        
        # Get routine
        routine = self.navigator.find_routine_by_name(routine_name)
        if not routine:
            return []
        
        # Process routine lines
        lines = self.navigator.get_routine_lines(routine)
        
        for line in lines:
            line_text = line.text if line.text else ''
            
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
        
        # Convert to structured list
        return self.format_sequences(sequences)
    
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
    
    def format_sequences(self, sequences: Dict) -> List[Dict[str, Any]]:
        """
        Format sequences for structured output.
        
        Args:
            sequences: Nested dictionary with sequences
            
        Returns:
            List of formatted sequences
        """
        formatted = []
        
        for seq_idx in sorted(sequences.keys()):
            sequence = {
                'sequence_index': seq_idx,
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