"""
Extractor for Digital Input tags (UDT_DigitalInputHal).
Extracts tags with DataType='UDT_DigitalInputHal' from all programs.
"""
from typing import List, Dict, Any
from ..core.base_extractor import BaseExtractor


class DigitalInputExtractor(BaseExtractor):
    """
    Extracts Digital Input tags of type UDT_DigitalInputHal.
    Captures tag name, description, and parent name from configuration.
    """
    
    def get_pattern(self) -> str:
        """
        This extractor doesn't use regex patterns.
        It searches by XML structure and attributes.
        
        Returns:
            Empty string (not used)
        """
        return ""
    
    def extract_all_digital_inputs(self, root) -> List[Dict[str, Any]]:
        """
        Extract all digital input tags from all programs in the L5X.
        
        Args:
            root: XML tree root
            
        Returns:
            List of digital input tags with their information
        """
        digital_inputs = []
        
        if self.debug:
            print("\n[DigitalInputExtractor] Searching for UDT_DigitalInputHal tags...")
        
        # Find all programs in the controller
        programs = root.findall('.//Programs/Program')
        
        if self.debug:
            print(f"  Found {len(programs)} program(s) to search")
        
        for program in programs:
            program_name = program.get('Name', 'Unknown')
            
            # Find all tags with DataType="UDT_DigitalInputHal"
            tags = program.findall('.//Tags/Tag[@DataType="UDT_DigitalInputHal"]')
            
            if tags and self.debug:
                print(f"  Program '{program_name}': {len(tags)} digital input tag(s)")
            
            for tag in tags:
                tag_name = tag.get('Name', '')
                
                # Extract description
                description_elem = tag.find('Description')
                description = ''
                if description_elem is not None and description_elem.text:
                    description = description_elem.text.strip()
                
                # Extract parent name from Cfg/ParentName/DATA
                parent_name = self._extract_parent_name(tag)
                
                digital_input = {
                    'program': program_name,
                    'tag_name': tag_name,
                    'description': description,
                    'parent_name': parent_name
                }
                
                digital_inputs.append(digital_input)
                
                if self.debug:
                    print(f"    [{tag_name}] Parent: {parent_name}")
        
        if self.debug:
            print(f"  Total digital inputs found: {len(digital_inputs)}")
        
        return digital_inputs
    
    def _extract_parent_name(self, tag) -> str:
        """
        Extract parent name from tag structure.
        Path: Data/Structure/StructureMember[@Name="Cfg"]/StructureMember[@Name="ParentName"]/DataValueMember[@Name="DATA"]
        
        Args:
            tag: Tag XML element
            
        Returns:
            Parent name string, or empty string if not found
        """
        try:
            # Navigate to Data/Structure
            structure = tag.find('.//Data/Structure')
            if structure is None:
                return ''
            
            # Find StructureMember with Name="Cfg"
            cfg_member = structure.find('.//StructureMember[@Name="Cfg"]')
            if cfg_member is None:
                return ''
            
            # Find StructureMember with Name="ParentName" inside Cfg
            parent_member = cfg_member.find('.//StructureMember[@Name="ParentName"]')
            if parent_member is None:
                return ''
            
            # Find DataValueMember with Name="DATA"
            data_value = parent_member.find('.//DataValueMember[@Name="DATA"]')
            if data_value is None:
                return ''
            
            # Get the text content and clean it
            if data_value.text:
                # Remove quotes and whitespace
                parent_name = data_value.text.strip().strip("'\"")
                return parent_name
            
            return ''
            
        except Exception as e:
            if self.debug:
                print(f"    Warning: Could not extract parent name - {str(e)}")
            return ''
    
    def find_items(self, root, routine_name: str) -> List[Dict[str, Any]]:
        """
        Not used for this extractor - we extract from all programs at once.
        
        Args:
            root: XML tree root
            routine_name: Not used
            
        Returns:
            Empty list
        """
        # This method is required by BaseExtractor but not used
        # Digital inputs are extracted from all programs at once
        return []
    
    def format_output(self, digital_inputs: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Format extracted data for output.
        
        Args:
            digital_inputs: List of digital input tags
            
        Returns:
            Dictionary with formatted output
        """
        return {
            'extractor_type': 'DigitalInputExtractor',
            'input_count': len(digital_inputs),
            'digital_inputs': digital_inputs
        }