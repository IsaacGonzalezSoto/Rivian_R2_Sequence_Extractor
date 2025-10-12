"""
Extractor for Actuator Group tags (AOI_Actuator).
Extracts tags with DataType='AOI_Actuator' from all programs.
"""
from typing import List, Dict, Any
from ..core.base_extractor import BaseExtractor
from ..core.logger import get_logger

logger = get_logger(__name__)


class ActuatorGroupExtractor(BaseExtractor):
    """
    Extracts Actuator Group tags of type AOI_Actuator.
    Captures tag name (MM number) and group description.
    """

    def get_pattern(self) -> str:
        """
        This extractor doesn't use regex patterns.
        It searches by XML structure and attributes.

        Returns:
            Empty string (not used)
        """
        return ""

    def extract_all_actuator_groups(self, root) -> List[Dict[str, Any]]:
        """
        Extract all actuator group tags from all programs in the L5X.

        Args:
            root: XML tree root

        Returns:
            List of actuator group tags with their information
        """
        actuator_groups = []

        if self.debug:
            logger.debug("[ActuatorGroupExtractor] Searching for AOI_Actuator tags...")

        # Find all programs in the controller
        programs = root.findall('.//Programs/Program')

        if self.debug:
            logger.debug(f"Found {len(programs)} program(s) to search")

        for program in programs:
            program_name = program.get('Name', 'Unknown')

            # Find all tags with DataType="AOI_Actuator"
            tags = program.findall('.//Tags/Tag[@DataType="AOI_Actuator"]')

            if tags and self.debug:
                logger.debug(f"Program '{program_name}': {len(tags)} actuator group tag(s)")

            for tag in tags:
                tag_name = tag.get('Name', '')

                # Extract description
                description_elem = tag.find('Description')
                description = ''
                if description_elem is not None and description_elem.text:
                    description = description_elem.text.strip()

                actuator_group = {
                    'program': program_name,
                    'tag_name': tag_name,
                    'description': description
                }

                actuator_groups.append(actuator_group)

                if self.debug:
                    logger.debug(f"  [{tag_name}] Description: {description}")

        if self.debug:
            logger.debug(f"Total actuator groups found: {len(actuator_groups)}")

        return actuator_groups

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
        # Actuator groups are extracted from all programs at once
        return []

    def format_output(self, actuator_groups: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Format extracted data for output.

        Args:
            actuator_groups: List of actuator group tags

        Returns:
            Dictionary with formatted output
        """
        return {
            'extractor_type': 'ActuatorGroupExtractor',
            'group_count': len(actuator_groups),
            'actuator_groups': actuator_groups
        }
