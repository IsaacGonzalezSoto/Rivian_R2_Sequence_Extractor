"""
Data exporter to CSV format.
"""
import csv
from typing import Dict, Any, List


class CSVExporter:
    """
    Exports sequence and actuator data to CSV format.
    One row per individual actuator.
    """
    
    def export(self, data: Dict[str, Any], output_path: str):
        """
        Export data to CSV with specific format.
        
        Args:
            data: Dictionary with sequence data
            output_path: Output CSV file path
        """
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'Routine',
                'Sequence',
                'Step',
                'Action_Index',
                'Action_Name',
                'MM_Number',
                'State',
                'Actuator_Count',
                'Actuators',
                'Validation_Status',
                'Missing_Indices'
            ]
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            routine_name = data['routine_name']
            
            for sequence in data['sequences']:
                seq_idx = sequence['sequence_index']
                
                for step in sequence['steps']:
                    step_idx = step['step_index']
                    
                    for action in step['actions']:
                        self._write_action_rows(writer, routine_name, seq_idx, step_idx, action)
    
    def _write_action_rows(self, writer, routine_name: str, seq_idx: int, step_idx: int, action: Dict[str, Any]):
        """
        Write rows corresponding to an action.
        
        Args:
            writer: CSV writer
            routine_name: Routine name
            seq_idx: Sequence index
            step_idx: Step index
            action: Action data
        """
        # Determine validation status
        validation_status = 'N/A'
        missing_indices = ''
        
        if 'validation' in action:
            val = action['validation']
            if val['is_valid'] == True:
                validation_status = 'OK'
            elif val['is_valid'] == False:
                validation_status = 'WARNING'
                missing_indices = ','.join(map(str, val['missing_indices']))
            elif val['is_valid'] is None:
                validation_status = 'UNKNOWN'
        
        # Format state: "Work" → "TO WORK" (uppercase)
        state_formatted = ''
        if action['state']:
            state_formatted = f"TO {action['state'].upper()}"
        
        # If no actuators, create one row without actuator
        if not action['actuators']:
            row = {
                'Routine': routine_name,
                'Sequence': seq_idx,
                'Step': step_idx,
                'Action_Index': action['action_index'],
                'Action_Name': action['action_name'],
                'MM_Number': action['mm_number'] or '',
                'State': state_formatted,
                'Actuator_Count': action['actuator_count'],
                'Actuators': '',
                'Validation_Status': validation_status,
                'Missing_Indices': missing_indices
            }
            writer.writerow(row)
        else:
            # Create one row per actuator
            for actuator in action['actuators']:
                # Show original data without special formatting
                actuator_description = actuator['description']
                
                row = {
                    'Routine': routine_name,
                    'Sequence': seq_idx,
                    'Step': step_idx,
                    'Action_Index': action['action_index'],
                    'Action_Name': action['action_name'],
                    'MM_Number': action['mm_number'] or '',
                    'State': state_formatted,
                    'Actuator_Count': action['actuator_count'],
                    'Actuators': actuator_description,
                    'Validation_Status': validation_status,
                    'Missing_Indices': missing_indices
                }
                
                writer.writerow(row)