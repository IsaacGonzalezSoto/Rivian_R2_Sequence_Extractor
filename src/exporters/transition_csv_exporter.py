"""
CSV exporter specifically for transition permissions.
"""
import csv
from typing import Dict, Any


class TransitionCSVExporter:
    """
    Exports transition permission data to CSV format.
    One row per permission.
    """
    
    def export(self, data: Dict[str, Any], output_path: str):
        """
        Export transition data to CSV.
        
        Args:
            data: Dictionary with transition data
            output_path: Output CSV file path
        """
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'Routine',
                'Transition_Index',
                'Permission_Count',
                'Permission_Index',
                'Permission_Value',
                'Comment'
            ]
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            routine_name = data['routine_name']
            
            for transition in data['transitions']:
                trans_idx = transition['transition_index']
                permission_count = transition['permission_count']
                
                for permission in transition['permissions']:
                    row = {
                        'Routine': routine_name,
                        'Transition_Index': trans_idx,
                        'Permission_Count': permission_count,
                        'Permission_Index': permission['permission_index'],
                        'Permission_Value': permission['permission_value'],
                        'Comment': permission['comment']
                    }
                    
                    writer.writerow(row)