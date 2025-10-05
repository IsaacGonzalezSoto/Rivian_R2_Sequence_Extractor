"""
Data exporter to JSON format.
"""
import json
from typing import Dict, Any


class JSONExporter:
    """
    Exports sequence and actuator data to JSON format.
    """
    
    def export(self, data: Dict[str, Any], output_path: str, indent: int = 2):
        """
        Export data to JSON with readable format.
        
        Args:
            data: Dictionary with data to export
            output_path: Output JSON file path
            indent: Indentation level for readable format
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=indent, ensure_ascii=False)