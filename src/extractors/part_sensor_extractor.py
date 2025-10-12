"""
Extractor for Part-Sensor relationships.
Extracts sensor assignments from Part routines (Cm{digits}_Part{X}).
"""
import re
from typing import List, Dict, Any
from ..core.base_extractor import BaseExtractor
from ..core.logger import get_logger

logger = get_logger(__name__)


class PartSensorExtractor(BaseExtractor):
    """
    Extracts sensor-to-part assignments from Part routines.
    Identifies which sensors belong to which part present detection.
    """

    # Pattern to match Part routines: Cm{digits}_Part{number}
    PART_ROUTINE_PATTERN = r'Cm\d+_Part(\d+)'

    # Pattern to match sensor assignments: XIC(SENSOR_NAME.Out.Value) OTE(PartX.inpSensors.Y)
    SENSOR_ASSIGNMENT_PATTERN = r'XIC\(([A-Za-z0-9_]+)\.Out\.Value\)\s+OTE\(Part(\d+)\.inpSensors\.(\d+)\)'

    def get_pattern(self) -> str:
        """
        Returns the pattern for sensor assignments.

        Returns:
            Regex pattern string
        """
        return self.SENSOR_ASSIGNMENT_PATTERN

    def extract_all_part_sensors(self, root) -> Dict[str, List[str]]:
        """
        Extract all sensor-to-part mappings from Part routines.

        Args:
            root: XML tree root

        Returns:
            Dictionary mapping sensor_name -> list of part names (e.g., {'BG1_BGB1': ['Part1', 'Part2']})
        """
        sensor_to_parts = {}  # sensor_name -> [part_names]

        if self.debug:
            logger.debug("[PartSensorExtractor] Searching for Part routines...")

        # Find all routines in the controller
        routines = root.findall('.//Programs/Program/Routines/Routine')

        part_routines_found = []

        for routine in routines:
            routine_name = routine.get('Name', '')

            # Check if this is a Part routine
            match = re.search(self.PART_ROUTINE_PATTERN, routine_name)
            if match:
                part_number = match.group(1)
                part_name = f"Part{part_number}"
                part_routines_found.append((routine_name, part_name))

                if self.debug:
                    logger.debug(f"Found Part routine: {routine_name} → {part_name}")

                # Extract sensors from this routine
                sensors = self._extract_sensors_from_part_routine(routine, part_name)

                # Map sensors to parts
                for sensor_name in sensors:
                    if sensor_name not in sensor_to_parts:
                        sensor_to_parts[sensor_name] = []
                    if part_name not in sensor_to_parts[sensor_name]:
                        sensor_to_parts[sensor_name].append(part_name)
                        if self.debug:
                            logger.debug(f"  Sensor '{sensor_name}' → {part_name}")

        # Validate against AOI_Part tags
        self._validate_part_counts(root, part_routines_found)

        if self.debug:
            logger.debug(f"Total unique sensors mapped: {len(sensor_to_parts)}")

        return sensor_to_parts

    def _extract_sensors_from_part_routine(self, routine, part_name: str) -> List[str]:
        """
        Extract sensor names from a specific Part routine.
        Uses the standard pattern: XIC(SENSOR.Out.Value) OTE(PartX.inpSensors.Y)

        Args:
            routine: Routine XML element
            part_name: Name of the part (e.g., 'Part1')

        Returns:
            List of sensor names
        """
        sensors = []

        # Get all rungs in the routine (RLL format)
        rungs = routine.findall('.//Rung')

        # Extract part number from part_name (e.g., 'Part1' -> '1')
        part_number = part_name.replace('Part', '')

        for rung in rungs:
            # Get the Text element containing the ladder logic
            text_elem = rung.find('.//Text')
            if text_elem is None or text_elem.text is None:
                continue

            text_content = text_elem.text

            # Search for the standard sensor assignment pattern
            # Pattern: XIC(SENSOR_NAME.Out.Value) OTE(PartX.inpSensors.Y)
            matches = re.finditer(self.SENSOR_ASSIGNMENT_PATTERN, text_content)

            for match in matches:
                sensor_name = match.group(1)
                matched_part_num = match.group(2)
                sensor_index = match.group(3)

                # Verify that the part number matches
                if f"Part{matched_part_num}" == part_name:
                    sensors.append(sensor_name)
                    if self.debug:
                        logger.debug(f"    Found sensor: {sensor_name} at index {sensor_index}")

        return sensors

    def _validate_part_counts(self, root, part_routines_found: List[tuple]):
        """
        Validate that the number of Part routines matches the number of AOI_Part tags.

        Args:
            root: XML tree root
            part_routines_found: List of (routine_name, part_name) tuples
        """
        # Find all AOI_Part tags
        aoi_part_tags = root.findall('.//Tag[@DataType="AOI_Part"]')

        routine_count = len(part_routines_found)
        tag_count = len(aoi_part_tags)

        if self.debug:
            logger.debug(f"Validation: {routine_count} Part routines vs {tag_count} AOI_Part tags")

        if routine_count != tag_count:
            logger.warning(
                f"Part count mismatch! Found {routine_count} Part routines but {tag_count} AOI_Part tags. "
                f"Routines: {[r[0] for r in part_routines_found]}"
            )
        else:
            if self.debug:
                logger.debug("✓ Part counts match")

    def find_items(self, root, routine_name: str) -> List[Dict[str, Any]]:
        """
        Not used for this extractor - we extract from all Part routines at once.

        Args:
            root: XML tree root
            routine_name: Not used

        Returns:
            Empty list
        """
        return []

    def format_output(self, sensor_to_parts: Dict[str, List[str]]) -> Dict[str, Any]:
        """
        Format extracted data for output.

        Args:
            sensor_to_parts: Dictionary mapping sensor names to part lists

        Returns:
            Dictionary with formatted output
        """
        return {
            'extractor_type': 'PartSensorExtractor',
            'sensor_count': len(sensor_to_parts),
            'sensor_to_parts': sensor_to_parts
        }
