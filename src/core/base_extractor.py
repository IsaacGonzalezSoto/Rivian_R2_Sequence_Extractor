"""
Abstract base class for all extractors.
Defines the template method pattern for L5X data extraction.
"""
from abc import ABC, abstractmethod
from typing import Dict, List, Any
from .logger import get_logger

logger = get_logger(__name__)


class BaseExtractor(ABC):
    """
    Abstract base class for L5X data extractors.
    
    Template Method Pattern: Defines the general extraction flow,
    delegating specific steps to subclasses.
    """
    
    def __init__(self, debug: bool = False):
        """
        Initialize the base extractor.
        
        Args:
            debug: Enable debug mode for detailed logging
        """
        self.debug = debug
    
    def extract(self, root, routine_name: str) -> Dict[str, Any]:
        """
        Template method: Defines the complete extraction flow.
        
        Args:
            root: Root of the L5X XML tree
            routine_name: Name of the routine to process
            
        Returns:
            Dictionary with extracted data
        """
        if self.debug:
            logger.debug(f"[{self.__class__.__name__}] Processing routine: {routine_name}")

        # 1. Find items according to the extractor's specific pattern
        items = self.find_items(root, routine_name)

        if self.debug:
            logger.debug(f"[{self.__class__.__name__}] Items found: {len(items)}")
        
        # 2. Validate the found items
        validated_items = self.validate_items(root, items)
        
        # 3. Format the output
        return self.format_output(routine_name, validated_items)
    
    @abstractmethod
    def find_items(self, root, routine_name: str) -> List[Dict[str, Any]]:
        """
        Find and extract specific items from the routine.
        Must be implemented by each concrete extractor.
        
        Args:
            root: XML tree root
            routine_name: Routine name
            
        Returns:
            List of dictionaries with found items
        """
        pass
    
    @abstractmethod
    def get_pattern(self) -> str:
        """
        Returns the extractor's specific regex pattern.
        
        Returns:
            String with the regex pattern
        """
        pass
    
    def validate_items(self, root, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Validate extracted items.
        Can be overridden by subclasses that need specific validation.
        
        Args:
            root: XML tree root
            items: List of items to validate
            
        Returns:
            List of validated items with additional information
        """
        # Default validation: return items unchanged
        return items
    
    def format_output(self, routine_name: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Format extracted data for output.
        Can be overridden by subclasses that need specific formatting.
        
        Args:
            routine_name: Routine name
            items: Validated items
            
        Returns:
            Dictionary with output format
        """
        return {
            'routine_name': routine_name,
            'extractor_type': self.__class__.__name__,
            'items': items,
            'count': len(items)
        }