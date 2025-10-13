"""
Specialized extractor for transition permissions from EmStatesAndSequences routines.
"""
import re
from typing import Dict, List, Any
from collections import defaultdict
from ..core.base_extractor import BaseExtractor
from ..core.logger import get_logger

logger = get_logger(__name__)


class TransitionExtractor(BaseExtractor):
    """
    Extracts transition permission information from EmStatesAndSequences routines.
    Searches for patterns: EmTransitionStates[X].AutoStartPerms.Y := Value; //Comment
    """
    
    def get_pattern(self) -> str:
        """
        Returns the regex pattern to extract transition permissions.
        
        Returns:
            Regex pattern for EmTransitionStates assignments
        """
        return r'EmTransitionStates\[(\d+)\]\.AutoStartPerms\.(\d+)\s*:=\s*([^;]+);\s*(?://(.*))?'
    
    def find_items(self, root, routine_name: str, program_name: str = None) -> List[Dict[str, Any]]:
        """
        Search for transition permissions in an EmStatesAndSequences routine.

        Args:
            root: XML tree root or Program element
            routine_name: Routine name (e.g., 'EmStatesAndSequences_R2S')
            program_name: Optional program name for scoping (multi-fixture support)

        Returns:
            List of dictionaries with transition permission information
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
                logger.warning(f"Program not found: {program_name}, searching globally")
                routine = navigator.find_routine_by_name(routine_name)
        else:
            routine = navigator.find_routine_by_name(routine_name)

        if not routine:
            if self.debug:
                logger.warning(f"Routine not found: {routine_name}")
            return []
        
        # Get the regex pattern
        pattern = self.get_pattern()
        
        # Pattern to extract transition names from #region comments
        region_pattern = r'#region\s+Transition\s+State\s+(\d+)\s+-\s+(.+)'
        
        # Structure to store transitions
        transitions = defaultdict(list)
        transition_names = {}  # Map transition_index -> descriptive_name
        
        # Get all lines from the routine
        lines = navigator.get_routine_lines(routine)
        
        for line in lines:
            line_text = line.text if line.text else ''
            
            # Check for #region comments to extract transition names
            region_match = re.search(region_pattern, line_text)
            if region_match:
                trans_idx = int(region_match.group(1))
                trans_name = region_match.group(2).strip()
                transition_names[trans_idx] = trans_name
                if self.debug:
                    logger.debug(f"Found transition name: State {trans_idx} - {trans_name}")
            
            # Search for all pattern matches (permissions)
            matches = re.finditer(pattern, line_text)
            for match in matches:
                transition_idx = int(match.group(1))
                permission_idx = int(match.group(2))
                permission_value = match.group(3).strip()
                comment = match.group(4).strip() if match.group(4) else ""
                
                permission_data = {
                    'permission_index': permission_idx,
                    'permission_value': permission_value,
                    'comment': comment,
                    'full_assignment': match.group(0)
                }
                
                transitions[transition_idx].append(permission_data)

                if self.debug:
                    logger.debug(f"Transition[{transition_idx}].Permission[{permission_idx}] = {permission_value}")
                    if comment:
                        logger.debug(f"  Comment: {comment}")
        
        # Convert to list format
        result = []
        for trans_idx in sorted(transitions.keys()):
            # Sort permissions by index
            permissions = sorted(transitions[trans_idx], key=lambda x: x['permission_index'])
            
            # Get descriptive name if available
            descriptive_name = transition_names.get(trans_idx, None)
            
            result.append({
                'transition_index': trans_idx,
                'transition_name': descriptive_name,  # New field
                'permission_count': len(permissions),
                'permissions': permissions
            })
        
        return result
    
    def format_output(self, routine_name: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Format extracted data for output.
        
        Args:
            routine_name: Routine name
            items: List of transitions with permissions
            
        Returns:
            Dictionary with formatted output
        """
        return {
            'routine_name': routine_name,
            'extractor_type': 'TransitionExtractor',
            'transition_count': len(items),
            'transitions': items
        }