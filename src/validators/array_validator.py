"""
Validator to verify the completeness of actuator arrays.
"""
from typing import List, Dict, Any
from ..core.xml_navigator import XMLNavigator
from ..core.logger import get_logger

logger = get_logger(__name__)


class ArrayValidator:
    """
    Validates that all indices of an array have an assigned description.
    """
    
    def __init__(self, debug: bool = False):
        """
        Initialize the validator.
        
        Args:
            debug: Debug mode for logging
        """
        self.debug = debug
    
    def validate_actuators(self, root, mm_number: str, actuators: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Validates that all indices of the actuator array have a description.

        Args:
            root: XML tree root
            mm_number: MM number (e.g., 'MM4')
            actuators: List of found actuators

        Returns:
            Dictionary with validation information
        """
        navigator = XMLNavigator(root=root)
        
        array_name = f'{mm_number}Cyls'
        dimension = navigator.get_tag_dimension(array_name)
        
        validation = {
            'is_valid': True,
            'array_name': array_name,
            'array_dimension': dimension,
            'descriptions_found': len(actuators),
            'missing_indices': []
        }
        
        if dimension is None:
            validation['is_valid'] = None
            validation['warning'] = f'Tag {array_name} not found'
            if self.debug:
                logger.warning(f"No dimension found for {array_name}")
            return validation
        
        # Get indices that have a description
        found_indices = set(act['index'] for act in actuators)
        
        # Check missing indices (0 to dimension-1)
        expected_indices = set(range(dimension))
        missing_indices = sorted(expected_indices - found_indices)
        
        if missing_indices:
            validation['is_valid'] = False
            validation['missing_indices'] = missing_indices
            if self.debug:
                logger.warning(f"Missing descriptions for indices: {missing_indices}")
                logger.warning(f"Array dimension: {dimension}, Descriptions found: {len(actuators)}")
        else:
            if self.debug:
                logger.info(f"✓ Validation OK: All {dimension} indices have a description")
        
        return validation
    
    def validate_all(self, root, actuators_by_mm: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Dict[str, Any]]:
        """
        Validates multiple actuator arrays.
        
        Args:
            root: XML tree root
            actuators_by_mm: Dictionary with MM as key and actuator list as value
            
        Returns:
            Dictionary with validations for each MM
        """
        validations = {}
        
        for mm_number, actuators in actuators_by_mm.items():
            validations[mm_number] = self.validate_actuators(root, mm_number, actuators)
        
        return validations