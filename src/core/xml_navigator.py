"""
Utilities for navigating the L5X file XML tree.
Provides common search and element access functions.
"""
import xml.etree.ElementTree as ET
from typing import List, Optional


class XMLNavigator:
    """
    Navigator for RSLogix 5000 L5X file XML trees.
    Encapsulates common paths and searches.
    """
    
    # Common XPath routes
    ROUTINES_PATH = './Controller/Programs/Program/Routines/Routine'
    TAGS_PATH = './Controller/Programs/Program/Tags/Tag'
    
    def __init__(self, l5x_file_path: str):
        """
        Initialize the navigator with an L5X file.
        
        Args:
            l5x_file_path: Path to the L5X file
        """
        self.tree = ET.parse(l5x_file_path)
        self.root = self.tree.getroot()
    
    def get_root(self):
        """Returns the XML tree root."""
        return self.root
    
    def find_all_routines(self) -> List[ET.Element]:
        """
        Find all routines in the program.
        
        Returns:
            List of Routine elements
        """
        return self.root.findall(self.ROUTINES_PATH)
    
    def find_routine_by_name(self, routine_name: str) -> Optional[ET.Element]:
        """
        Search for a specific routine by name.
        
        Args:
            routine_name: Routine name
            
        Returns:
            Routine element or None if not found
        """
        routines = self.root.findall(f'{self.ROUTINES_PATH}[@Name="{routine_name}"]')
        return routines[0] if routines else None
    
    def find_routines_by_pattern(self, pattern: str) -> List[ET.Element]:
        """
        Search for routines matching a pattern in the name.
        
        Args:
            pattern: Pattern to search (can be regex)
            
        Returns:
            List of matching routines
        """
        import re
        all_routines = self.find_all_routines()
        matching_routines = []
        
        for routine in all_routines:
            routine_name = routine.get('Name', '')
            if re.search(pattern, routine_name):
                matching_routines.append(routine)
        
        return matching_routines
    
    def find_routines_starting_with(self, prefix: str) -> List[ET.Element]:
        """
        Search for routines starting with a specific prefix.
        
        Args:
            prefix: Prefix to search
            
        Returns:
            List of matching routines
        """
        all_routines = self.find_all_routines()
        return [r for r in all_routines if r.get('Name', '').startswith(prefix)]
    
    def get_routine_lines(self, routine: ET.Element) -> List[ET.Element]:
        """
        Get all code lines from an ST routine.
        
        Args:
            routine: Routine element
            
        Returns:
            List of Line elements
        """
        # First try in STContent
        lines = routine.findall('.//STContent/Line')
        
        # If none found, search directly
        if not lines:
            lines = routine.findall('.//Line')
        
        return lines
    
    def get_routine_rungs(self, routine: ET.Element) -> List[ET.Element]:
        """
        Get all rungs from a Ladder (RLL) routine.
        
        Args:
            routine: Routine element
            
        Returns:
            List of Rung elements
        """
        return routine.findall('.//Rung')
    
    def find_tag_by_name(self, tag_name: str) -> Optional[ET.Element]:
        """
        Search for a specific tag by name.
        
        Args:
            tag_name: Tag name
            
        Returns:
            Tag element or None if not found
        """
        tags = self.root.findall(f'{self.TAGS_PATH}[@Name="{tag_name}"]')
        return tags[0] if tags else None
    
    def get_tag_dimension(self, tag_name: str) -> Optional[int]:
        """
        Get the dimension of an array tag.
        
        Args:
            tag_name: Tag name
            
        Returns:
            Array dimension or None if not found
        """
        import re
        tag = self.find_tag_by_name(tag_name)
        
        if tag:
            dimensions = tag.get('Dimensions', '')
            if dimensions:
                try:
                    return int(dimensions)
                except ValueError:
                    # If complex format, extract first number
                    match = re.search(r'\d+', dimensions)
                    if match:
                        return int(match.group(0))
        
        return None
    
    def get_routine_info(self, routine: ET.Element) -> dict:
        """
        Get basic information from a routine.
        
        Args:
            routine: Routine element
            
        Returns:
            Dictionary with routine information
        """
        return {
            'name': routine.get('Name', 'Unknown'),
            'type': routine.get('Type', 'Unknown'),
            'description': routine.find('.//Description')
        }