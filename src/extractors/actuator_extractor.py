"""
Specialized extractor for actuators from MM routines.
"""
import re
from typing import Dict, List, Any
from ..core.base_extractor import BaseExtractor
from ..core.logger import get_logger

logger = get_logger(__name__)


class ActuatorExtractor(BaseExtractor):
    """
    Extracts actuator information from MM routines in the L5X.
    Searches for patterns: MOVE('DESCRIPTION', MM{X}Cyls[INDEX].Stg.Name)
    """
    
    def get_pattern(self) -> str:
        """
        Returns the regex pattern to extract actuators.
        
        Returns:
            Regex pattern for MOVE statements
        """
        return r"MOVE\('([^']+)',\s*{mm_number}Cyls\[(\d+)\]\.Stg\.Name\)"
    
    def find_items(self, root, routine_name: str, program_name: str = None) -> List[Dict[str, Any]]:
        """
        Search for actuators in a specific MM routine.

        Args:
            root: XML tree root
            routine_name: Routine name (e.g., 'Cm010507_MM4')
            program_name: Optional program name for scoping (multi-fixture support)

        Returns:
            List of dictionaries with actuator information
        """
        from ..core.xml_navigator import XMLNavigator

        navigator = XMLNavigator(root=root)

        # Search for the specific routine (scoped to program if provided)
        if program_name:
            program_element = navigator.find_program_by_name(program_name)
            if program_element:
                routines = program_element.findall(f'.//Routines/Routine[@Name="{routine_name}"]')
                routine = routines[0] if routines else None
            else:
                if self.debug:
                    logger.warning(f"Program not found: {program_name}, searching globally")
                routine = navigator.find_routine_by_name(routine_name)
        else:
            routine = navigator.find_routine_by_name(routine_name)
        
        if not routine:
            if self.debug:
                logger.warning(f"Routine not found: {routine_name}")
            return []

        # Extract MM number from routine name
        mm_match = re.search(r'(MM\d+)', routine_name)
        if not mm_match:
            if self.debug:
                logger.warning(f"Could not extract MM number from: {routine_name}")
            return []
        
        mm_number = mm_match.group(1)
        
        # Get pattern with specific MM number
        pattern = self.get_pattern().replace('{mm_number}', mm_number)
        
        actuators = []
        
        # Search in all rungs of the routine
        rungs = navigator.get_routine_rungs(routine)
        
        for rung in rungs:
            text_element = rung.find('.//Text')
            if text_element is not None and text_element.text:
                text = text_element.text
                
                # Find all pattern matches
                matches = re.finditer(pattern, text)
                for match in matches:
                    description = match.group(1)
                    index = int(match.group(2))
                    
                    actuators.append({
                        'index': index,
                        'description': description,
                        'mm_number': mm_number
                    })

                    if self.debug:
                        logger.debug(f"    [{index}] {description}")
        
        # Sort by index
        actuators.sort(key=lambda x: x['index'])
        
        return actuators
    
    def find_actuators_for_mm(self, root, mm_number: str, program_name: str = None) -> List[Dict[str, Any]]:
        """
        Search for all actuators for a specific MM.
        Automatically finds the corresponding routine.

        Args:
            root: XML tree root
            mm_number: MM number (e.g., 'MM4')
            program_name: Optional program name for scoping (multi-fixture support)

        Returns:
            List of found actuators
        """
        from ..core.xml_navigator import XMLNavigator

        navigator = XMLNavigator(root=root)

        # Pattern to find the correct routine (e.g., Cm010507_MM4)
        routine_pattern = f'_{mm_number}$|_{mm_number}_'

        # Search for routines containing the pattern (scoped to program if provided)
        if program_name:
            program_element = navigator.find_program_by_name(program_name)
            if program_element:
                all_routines = program_element.findall('.//Routines/Routine[@Name]')
                matching_routines = [r for r in all_routines if r.get('Name') and
                                   (r.get('Name').endswith(f'_{mm_number}') or f'_{mm_number}_' in r.get('Name'))]
            else:
                if self.debug:
                    logger.warning(f"Program not found: {program_name}, searching globally")
                matching_routines = navigator.find_routines_by_pattern(routine_pattern)
        else:
            matching_routines = navigator.find_routines_by_pattern(routine_pattern)

        if not matching_routines:
            if self.debug:
                logger.warning(f"No routine found for {mm_number}")
            return []

        # Use the first matching routine
        routine = matching_routines[0]
        routine_name = routine.get('Name', '')

        if self.debug:
            logger.debug(f"  Found actuator routine: {routine_name}")

        # Extract actuators
        return self.find_items(root, routine_name, program_name=program_name)