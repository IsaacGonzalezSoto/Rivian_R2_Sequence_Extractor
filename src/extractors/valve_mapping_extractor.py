"""
Extractor for valve mappings from MapIo program.
Maps MM groups to physical valve manifolds and valve positions.
"""
import re
from typing import Dict, List, Any, Optional
from ..core.base_extractor import BaseExtractor
from ..core.logger import get_logger

logger = get_logger(__name__)


class ValveMappingExtractor(BaseExtractor):
    """
    Extracts valve mapping information from MapIo program.

    Supports AOI_ValveManifold_V4, V8, V12, and V16 (generic version detection).

    Workflow:
    1. Extract MM commands (MM1_ToWork, MM1_ToHome) from fixture program
    2. Find MapIo program in L5X
    3. Parse AOI_ValveManifold_V* calls (V4, V8, V12, V16)
    4. Match commands to valve positions
    5. Return mapping: {MM1: {manifold: '...', work: '1A', home: '1B'}}
    """

    # Pattern to find MM routines
    MM_ROUTINE_PATTERN = r'Cm\d+_MM(\d+)'

    # Pattern to find MM commands in ladder logic
    # Example: ,XIC(MM1.outWork) OTE(MM1_ToWork.Inp.Value
    COMMAND_PATTERN = r',XIC\(MM(\d+)\.(out(?:Work|Home))\)\s+OTE\(([A-Za-z0-9_]+)'

    # Pattern to find AOI_ValveManifold calls (supports V4, V8, V12, V16, etc.)
    AOI_PATTERN = r'AOI_ValveManifold_V\d+\(([^)]+)\)'

    def get_pattern(self) -> str:
        """
        Returns the regex pattern (not used for this extractor).

        Returns:
            Empty string
        """
        return ""

    def find_items(self, root, routine_name: str = None, program_name: str = None) -> List[Dict[str, Any]]:
        """
        Find valve mappings for a specific fixture program.

        Args:
            root: XML tree root
            routine_name: Not used (inherited from base class)
            program_name: Name of the fixture program

        Returns:
            List with valve mapping information
        """
        if not program_name:
            if self.debug:
                logger.warning("No program_name provided, cannot extract valve mappings")
            return []

        # Step 1: Extract MM commands from fixture program
        mm_commands = self._extract_mm_commands_from_fixture(root, program_name)

        if not mm_commands:
            if self.debug:
                logger.debug(f"No MM commands found in {program_name}")
            return []

        # Step 2: Extract valve mappings from MapIo
        valve_mappings = self._extract_valve_mappings_from_mapio(root, program_name, mm_commands)

        return valve_mappings

    def _extract_mm_commands_from_fixture(self, root, program_name: str) -> Dict[str, Dict[str, str]]:
        """
        Extract MM command names from fixture program's MM routines.

        Args:
            root: XML tree root
            program_name: Fixture program name

        Returns:
            Dictionary: {MM1: {work_cmd: 'MM1_ToWork', home_cmd: 'MM1_ToHome'}, ...}
        """
        mm_commands = {}

        # Find the fixture program
        programs = root.findall('.//Programs/Program')
        fixture_program = None

        for prog in programs:
            if prog.get('Name') == program_name:
                fixture_program = prog
                break

        if not fixture_program:
            if self.debug:
                logger.warning(f"Fixture program not found: {program_name}")
            return mm_commands

        # Find MM routines (Cm{digits}_MM{N})
        routines = fixture_program.findall('.//Routines/Routine')

        for routine in routines:
            routine_name = routine.get('Name', '')
            match = re.search(self.MM_ROUTINE_PATTERN, routine_name)

            if match:
                mm_number = match.group(1)
                mm_key = f"MM{mm_number}"

                # Initialize commands dict for this MM
                if mm_key not in mm_commands:
                    mm_commands[mm_key] = {'work_cmd': None, 'home_cmd': None}

                # Search for command patterns in rungs
                rungs = routine.findall('.//Rung')

                for rung in rungs:
                    text_elem = rung.find('.//Text')
                    if text_elem is not None and text_elem.text:
                        text = text_elem.text

                        # Find all command patterns
                        cmd_matches = re.finditer(self.COMMAND_PATTERN, text)

                        for cmd_match in cmd_matches:
                            cmd_mm_num = cmd_match.group(1)
                            cmd_type = cmd_match.group(2)  # outWork or outHome
                            cmd_name = cmd_match.group(3)  # MM1_ToWork

                            if cmd_mm_num == mm_number:
                                if cmd_type == 'outWork':
                                    mm_commands[mm_key]['work_cmd'] = cmd_name
                                elif cmd_type == 'outHome':
                                    mm_commands[mm_key]['home_cmd'] = cmd_name

                if self.debug:
                    logger.debug(f"Found {mm_key}: Work={mm_commands[mm_key]['work_cmd']}, Home={mm_commands[mm_key]['home_cmd']}")

        return mm_commands

    def _extract_valve_mappings_from_mapio(self, root, program_name: str, mm_commands: Dict) -> List[Dict[str, Any]]:
        """
        Extract valve mappings from MapIo program's AOI_ValveManifold_V* calls.

        Supports all AOI versions: V4, V8, V12, V16.

        Args:
            root: XML tree root
            program_name: Fixture program name (used to filter relevant mappings)
            mm_commands: MM commands dict from fixture program

        Returns:
            List of valve mappings: [{mm_number: 'MM1', manifold: '...', valve_work: '1A', valve_home: '1B'}, ...]
        """
        valve_mappings = []

        # Find MapIo program
        programs = root.findall('.//Programs/Program')
        mapio_program = None

        for prog in programs:
            if prog.get('Name') == 'MapIo':
                mapio_program = prog
                break

        if not mapio_program:
            if self.debug:
                logger.debug("MapIo program not found, no valve mappings available")
            return valve_mappings

        # Search all routines in MapIo for AOI_ValveManifold_V8
        routines = mapio_program.findall('.//Routines/Routine')

        for routine in routines:
            rungs = routine.findall('.//Rung')

            for rung in rungs:
                text_elem = rung.find('.//Text')
                if text_elem is not None and text_elem.text:
                    text = text_elem.text

                    # Check if this rung contains our fixture's commands
                    # Handle backslash escapes in ladder logic text (e.g., \_010UA1_Fixture_Em0105)
                    if program_name not in text and f"\\{program_name}" not in text:
                        continue

                    # Find AOI_ValveManifold_V8 calls
                    aoi_matches = re.finditer(self.AOI_PATTERN, text)

                    for aoi_match in aoi_matches:
                        params_str = aoi_match.group(1)

                        # Parse AOI parameters
                        mappings = self._parse_aoi_valvemanifold(params_str, program_name, mm_commands)
                        valve_mappings.extend(mappings)

        return valve_mappings

    def _parse_aoi_valvemanifold(self, params_str: str, program_name: str, mm_commands: Dict) -> List[Dict[str, Any]]:
        """
        Parse AOI_ValveManifold_V8 parameters and extract valve mappings.

        Args:
            params_str: Parameter string from AOI call
            program_name: Fixture program name
            mm_commands: MM commands dict

        Returns:
            List of valve mappings for this AOI call
        """
        mappings = []

        # Split parameters by comma
        params = [p.strip() for p in params_str.split(',')]

        if len(params) < 7:
            if self.debug:
                logger.warning(f"AOI has too few parameters: {len(params)}")
            return mappings

        # Extract manifold name (parameter 3, index 2)
        manifold_name = params[2]

        if self.debug:
            logger.debug(f"Processing AOI with manifold: {manifold_name}")

        # Parameters 6+ (index 5+) contain valve assignments
        # Each pair represents one valve: (Work, Home)
        for i in range(5, len(params), 2):
            work_param = params[i] if i < len(params) else None
            home_param = params[i + 1] if i + 1 < len(params) else None

            if not work_param or not home_param:
                break

            # Skip Spare.DO entries
            if 'Spare.DO' in work_param and 'Spare.DO' in home_param:
                continue

            # Calculate valve index (1-based)
            valve_index = ((i - 5) // 2) + 1

            # Find which MM this valve belongs to by matching command names
            for mm_key, commands in mm_commands.items():
                work_cmd = commands.get('work_cmd')
                home_cmd = commands.get('home_cmd')

                # Check if work command matches
                work_match = work_cmd and work_cmd in work_param
                home_match = home_cmd and home_cmd in home_param

                # Handle monoestable valves (only Work or only Home)
                if work_match or home_match:
                    valve_work = f"{valve_index}A" if work_match and 'Spare.DO' not in work_param else "N/A"
                    valve_home = f"{valve_index}B" if home_match and 'Spare.DO' not in home_param else "N/A"

                    mapping = {
                        'mm_number': mm_key,
                        'manifold': manifold_name,
                        'valve_work': valve_work,
                        'valve_home': valve_home
                    }

                    mappings.append(mapping)

                    if self.debug:
                        logger.debug(f"Mapped {mm_key}: Manifold={manifold_name}, Work={valve_work}, Home={valve_home}")

                    break  # Found the MM, move to next valve

        return mappings

    def format_output(self, routine_name: str = None, items: List[Dict[str, Any]] = None) -> Dict[str, Any]:
        """
        Format valve mappings for output.

        Args:
            routine_name: Not used
            items: List of valve mappings

        Returns:
            Formatted output dictionary
        """
        if items is None:
            items = []

        # Convert list to dictionary keyed by MM number for easy lookup
        mappings_dict = {}
        for item in items:
            mm_num = item['mm_number']
            mappings_dict[mm_num] = {
                'manifold': item['manifold'],
                'valve_work': item['valve_work'],
                'valve_home': item['valve_home']
            }

        return {
            'valve_mappings': mappings_dict,
            'mapping_count': len(mappings_dict)
        }
