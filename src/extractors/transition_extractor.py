"""
Specialized extractor for transition permissions from EmStatesAndSequences routines.
"""
import re
from typing import Dict, List, Any
from collections import defaultdict
from ..core.base_extractor import BaseExtractor


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
    
    def find_items(self, root, routine_name: str) -> List[Dict[str, Any]]:
        """
        Search for transition permissions in an EmStatesAndSequences routine.
        
        Args:
            root: XML tree root
            routine_name: Routine name (e.g., 'EmStatesAndSequences_R2S')
            
        Returns:
            List of dictionaries with transition permission information
        """
        from ..core.xml_navigator import XMLNavigator
        
        navigator = XMLNavigator.__new__(XMLNavigator)
        navigator.root = root
        
        # Search for the specific routine
        routine = navigator.find_routine_by_name(routine_name)
        
        if not routine:
            if self.debug:
                print(f"  Warning: Routine not found: {routine_name}")
            return []
        
        # Get the regex pattern
        pattern = self.get_pattern()
        
        # Structure to store transitions
        transitions = defaultdict(list)
        
        # Get all lines from the routine
        lines = navigator.get_routine_lines(routine)
        
        for line in lines:
            line_text = line.text if line.text else ''
            
            # Search for all pattern matches
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
                    print(f"  Transition[{transition_idx}].Permission[{permission_idx}] = {permission_value}")
                    if comment:
                        print(f"    Comment: {comment}")
        
        # Convert to list format
        result = []
        for trans_idx in sorted(transitions.keys()):
            # Sort permissions by index
            permissions = sorted(transitions[trans_idx], key=lambda x: x['permission_index'])
            
            result.append({
                'transition_index': trans_idx,
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