"""
Utilities for navigating the L5X file XML tree.
Provides common search and element access functions.
"""
import xml.etree.ElementTree as ET
from typing import List, Optional
from .logger import get_logger

logger = get_logger(__name__)


class XMLNavigator:
    """
    Navigator for RSLogix 5000 L5X file XML trees.
    Encapsulates common paths and searches.
    """
    
    # Common XPath routes
    ROUTINES_PATH = './Controller/Programs/Program/Routines/Routine'
    TAGS_PATH = './Controller/Programs/Program/Tags/Tag'
    
    def __init__(self, l5x_file_path: str = None, root: ET.Element = None):
        """
        Initialize the navigator with an L5X file or existing root element.

        Args:
            l5x_file_path: Path to the L5X file (optional if root provided)
            root: Existing root element (optional if l5x_file_path provided)

        Raises:
            ValueError: If neither l5x_file_path nor root is provided
            FileNotFoundError: If l5x_file_path doesn't exist
            ET.ParseError: If XML parsing fails
        """
        if l5x_file_path:
            try:
                self.tree = ET.parse(l5x_file_path)
                self.root = self.tree.getroot()
                logger.info(f"Successfully loaded L5X file: {l5x_file_path}")
            except FileNotFoundError as e:
                logger.error(f"L5X file not found: {l5x_file_path}")
                raise FileNotFoundError(f"L5X file not found: {l5x_file_path}") from e
            except ET.ParseError as e:
                logger.error(f"Failed to parse XML file: {l5x_file_path} - {str(e)}")
                raise ET.ParseError(f"Invalid XML format in {l5x_file_path}: {str(e)}") from e
            except Exception as e:
                logger.error(f"Unexpected error loading L5X file: {l5x_file_path} - {str(e)}")
                raise RuntimeError(f"Failed to load L5X file: {str(e)}") from e
        elif root is not None:
            self.root = root
            self.tree = None
            logger.debug("Initialized XMLNavigator with existing root element")
        else:
            error_msg = "Either l5x_file_path or root must be provided"
            logger.error(error_msg)
            raise ValueError(error_msg)
    
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

    def find_all_programs(self) -> List[ET.Element]:
        """
        Find all programs in the controller.

        Returns:
            List of Program elements
        """
        return self.root.findall('.//Controller/Programs/Program')

    def find_programs_by_pattern(self, pattern: str) -> List[ET.Element]:
        """
        Search for programs matching a pattern in the name.

        Args:
            pattern: Pattern to search (regex)

        Returns:
            List of matching programs with their info
        """
        import re
        all_programs = self.find_all_programs()
        matching_programs = []

        for program in all_programs:
            program_name = program.get('Name', '')
            if re.search(pattern, program_name):
                matching_programs.append(program)

        return matching_programs

    def find_fixture_programs(self) -> List[dict]:
        """
        Identify fixture programs in the L5X file.

        Fixtures are identified by:
        1. Primary: Name contains pattern _\d{3}UA\d_ (e.g., _010UA1_)
        2. Secondary: Name contains "Fixture" word
        3. Validation: Must have at least one EmStatesAndSequences routine

        Returns:
            List of dictionaries with fixture program information:
            [
                {
                    'program_element': ET.Element,
                    'program_name': str,
                    'em_routines': List[str]  # Names of EmStatesAndSequences routines
                }
            ]
        """
        import re

        all_programs = self.find_all_programs()
        fixture_programs = []

        # Pattern for fixture identification: _\d{3}UA\d_
        fixture_pattern = r'_\d{3}UA\d_'

        for program in all_programs:
            program_name = program.get('Name', '')

            # Check if it matches fixture pattern or contains "Fixture"
            is_fixture_candidate = (
                re.search(fixture_pattern, program_name) or
                'Fixture' in program_name
            )

            if is_fixture_candidate:
                # Validate: must have EmStatesAndSequences routines
                routines = program.findall('.//Routines/Routine')
                em_routines = [
                    r.get('Name') for r in routines
                    if r.get('Name', '').startswith('EmStatesAndSequences')
                ]

                if em_routines:
                    fixture_programs.append({
                        'program_element': program,
                        'program_name': program_name,
                        'em_routines': em_routines
                    })
                    logger.debug(f"Identified fixture program: {program_name} with {len(em_routines)} EmStatesAndSequences routines")

        logger.info(f"Found {len(fixture_programs)} fixture program(s)")
        return fixture_programs

    def find_program_by_name(self, program_name: str) -> Optional[ET.Element]:
        """
        Find a program by its exact name.

        Args:
            program_name: Program name to search

        Returns:
            Program element or None if not found
        """
        programs = self.root.findall(f'.//Controller/Programs/Program[@Name="{program_name}"]')
        return programs[0] if programs else None

    def find_routines_in_program(self, program_name: str, prefix: str = None) -> List[ET.Element]:
        """
        Find routines within a specific program, optionally filtered by prefix.

        Args:
            program_name: Name of the program to search in
            prefix: Optional prefix filter for routine names

        Returns:
            List of Routine elements
        """
        program = self.find_program_by_name(program_name)
        if not program:
            logger.warning(f"Program not found: {program_name}")
            return []

        routines = program.findall('.//Routines/Routine')

        if prefix:
            routines = [r for r in routines if r.get('Name', '').startswith(prefix)]

        return routines