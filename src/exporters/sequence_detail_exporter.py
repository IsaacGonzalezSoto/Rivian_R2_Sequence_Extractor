"""
Sequence Detail exporter for generating the simplified Start Conditions format.
This exporter creates a single-sheet Excel file showing the initial state of all cylinders and sensors.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from typing import Dict, Any, List
from ..core.constants import ExcelColors, ExcelFontSizes, MAX_COLUMN_WIDTH, DEFAULT_COLUMN_PADDING


class SequenceDetailExporter:
    """
    Exports sequence detail data showing Start Conditions with cylinders and sensors.
    Creates a single Excel file with one sheet showing the initial state.
    """

    def __init__(self):
        """Initialize the Sequence Detail exporter."""
        # Header styles
        self.header_fill = PatternFill(start_color=ExcelColors.HEADER_FILL, end_color=ExcelColors.HEADER_FILL, fill_type="solid")
        self.header_font = Font(bold=True, color=ExcelColors.HEADER_FONT)
        self.header_alignment = Alignment(horizontal="center", vertical="center")

        # Data cell background - soft beige for eye comfort
        self.data_fill = PatternFill(start_color=ExcelColors.DATA_FILL, end_color=ExcelColors.DATA_FILL, fill_type="solid")

    def export(self, common_data: Dict[str, Any], model_data: Dict[str, Any],
               digital_inputs_data: Dict[str, Any], actuator_groups_data: Dict[str, Any],
               valve_mappings_data: Dict[str, Any], all_actuators_data: Dict[str, Any], output_path: str):
        """
        Export sequence detail showing Start Conditions (cylinders + sensors).

        The logic is:
        1. Common sequences define the base state (usually all cylinders to HOME)
        2. Model's first sequence may override some cylinder positions
        3. Result = Common state + Model first sequence overrides
        4. All sensors are listed with their part assignments

        Args:
            common_data: Sequences data from Common routine
            model_data: Sequences data from model routine (e.g., R2S)
            digital_inputs_data: Digital inputs data with part assignments
            actuator_groups_data: Actuator groups data for MM descriptions
            valve_mappings_data: Valve mappings data for valve nomenclature
            all_actuators_data: Complete list of ALL actuators from MM routines
            output_path: Path for the output Excel file
        """
        wb = Workbook()

        # Remove default sheet
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])

        # Create the main Sequence Detail sheet
        self._create_sequence_detail_sheet(wb, common_data, model_data, digital_inputs_data, actuator_groups_data, valve_mappings_data, all_actuators_data)

        # Save workbook
        wb.save(output_path)

    def _create_sequence_detail_sheet(self, wb: Workbook, common_data: Dict[str, Any],
                                     model_data: Dict[str, Any], digital_inputs_data: Dict[str, Any],
                                     actuator_groups_data: Dict[str, Any], valve_mappings_data: Dict[str, Any],
                                     all_actuators_data: Dict[str, Any]):
        """
        Create the Sequence Detail sheet with Start Conditions.

        Args:
            wb: Workbook object
            common_data: Common sequences data
            model_data: Model sequences data
            digital_inputs_data: Digital inputs data
            actuator_groups_data: Actuator groups data
            valve_mappings_data: Valve mappings data
            all_actuators_data: Complete list of ALL actuators from MM routines
        """
        ws = wb.create_sheet("Sequence Detail")

        # Headers matching Perfect_Output structure
        headers = [
            'ID',
            'Style',
            'Row Type',
            'Actor Type',
            'Actor/Unit',
            'Custom Description',
            'Custom \nDuration',
            'Standard \nAction/\nStatus',
            'Standard Duration',
            'Actor/Group \nDescription',
            'Actor/\nGroup \nName',
            'Valve \nName'
        ]

        # Write headers
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.header_alignment

        # Build MM description mapping
        mm_to_description = {}
        if actuator_groups_data and actuator_groups_data.get('actuator_groups'):
            for group in actuator_groups_data['actuator_groups']:
                mm_to_description[group['tag_name']] = group['description']

        # Build valve mapping dictionary: {MM1: {manifold: '...', valve_work: '1A', valve_home: '1B'}, ...}
        mm_to_valve = {}
        if valve_mappings_data and valve_mappings_data.get('valve_mappings'):
            mm_to_valve = valve_mappings_data['valve_mappings']

        # Build initial cylinder state from ALL actuators extracted from MM routines
        # Key: (mm_number, index) for uniqueness, Value: {actuator_name, state, mm_number, description, manifold, valve_work, valve_home}
        cylinder_state = {}

        # Initialize ALL cylinders with HOME state
        if all_actuators_data and all_actuators_data.get('actuators_by_mm'):
            for mm_number, actuators in all_actuators_data['actuators_by_mm'].items():
                mm_description = mm_to_description.get(mm_number, '')

                # Get valve mapping information
                manifold = ''
                valve_work = ''
                valve_home = ''
                if mm_number in mm_to_valve:
                    valve_info = mm_to_valve[mm_number]
                    manifold = valve_info.get('manifold', '')
                    valve_work = valve_info.get('valve_work', '')
                    valve_home = valve_info.get('valve_home', '')

                # Add all actuators with default HOME state
                for actuator in actuators:
                    actuator_name = actuator['description']
                    actuator_index = actuator['index']
                    # Use (mm_number, index) as key to ensure uniqueness
                    unique_key = (mm_number, actuator_index)
                    cylinder_state[unique_key] = {
                        'actuator_name': actuator_name,
                        'state': 'TO HOME',
                        'mm_number': mm_number,
                        'description': mm_description,
                        'manifold': manifold,
                        'valve_work': valve_work,
                        'valve_home': valve_home
                    }

        # Apply overrides from Common sequences
        if common_data and common_data.get('sequences'):
            for sequence in common_data['sequences']:
                for step in sequence['steps']:
                    for action in step['actions']:
                        state = action.get('state', '').upper()
                        state_formatted = f"TO {state}" if state else "TO HOME"
                        mm_number = action.get('mm_number', '')

                        for actuator in action.get('actuators', []):
                            actuator_index = actuator['index']
                            unique_key = (mm_number, actuator_index)
                            # Override state if this actuator exists
                            if unique_key in cylinder_state:
                                cylinder_state[unique_key]['state'] = state_formatted

        # Apply overrides from Model's FIRST sequence
        if model_data and model_data.get('sequences'):
            first_sequence = model_data['sequences'][0] if model_data['sequences'] else None
            if first_sequence:
                for step in first_sequence['steps']:
                    for action in step['actions']:
                        state = action.get('state', '').upper()
                        state_formatted = f"TO {state}" if state else "TO HOME"
                        mm_number = action.get('mm_number', '')

                        for actuator in action.get('actuators', []):
                            actuator_index = actuator['index']
                            unique_key = (mm_number, actuator_index)
                            # Override state if this actuator exists
                            if unique_key in cylinder_state:
                                cylinder_state[unique_key]['state'] = state_formatted

        # Sort cylinders by MM group and index
        sorted_cylinders = sorted(
            cylinder_state.items(),
            key=lambda x: (
                self._extract_mm_number_from_key(x[0][0]),  # Sort by MM number from key tuple
                x[0][1]  # Then by index
            )
        )

        # Write cylinder rows (Start Conditions, CylinderUnits)
        row_num = 2
        id_counter = 1

        for unique_key, info in sorted_cylinders:
            actuator_name = info['actuator_name']

            # Calculate valve name based on state
            valve_name = None
            if info.get('manifold'):
                kj_name = self._extract_kj_name(info['manifold'])
                # Select valve position based on state
                if 'WORK' in info['state']:
                    valve_name = self._build_valve_diagram_name(kj_name, info.get('valve_work', ''))
                else:  # HOME
                    valve_name = self._build_valve_diagram_name(kj_name, info.get('valve_home', ''))

            row_data = [
                id_counter,                      # ID
                'Common',                        # Style (always "Common" for start conditions)
                'Start Conditions',              # Row Type
                'CylinderUnits',                 # Actor Type
                None,                            # Actor/Unit (empty)
                None,                            # Custom Description (empty)
                None,                            # Custom Duration (empty)
                info['state'],                   # Standard Action/Status
                None,                            # Standard Duration (empty)
                info['description'],             # Actor/Group Description
                actuator_name,                   # Actor/Group Name
                valve_name                       # Valve Name
            ]

            for col_num, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_num, column=col_num, value=value)
                cell.fill = self.data_fill
            row_num += 1
            id_counter += 1

        # Extract and write sensor rows
        # Only include sensors that have part assignments (not 'N/A')
        if digital_inputs_data and digital_inputs_data.get('digital_inputs'):
            # Group sensors by part assignment
            sensors_by_part = {}
            for sensor in digital_inputs_data['digital_inputs']:
                tag_name = sensor.get('tag_name', '')
                part_assignment = sensor.get('part_assignment', 'N/A')

                # Only include sensors with actual part assignments and starting with BG (part sensors)
                if part_assignment != 'N/A' and tag_name.startswith('BG'):
                    if part_assignment not in sensors_by_part:
                        sensors_by_part[part_assignment] = []
                    sensors_by_part[part_assignment].append(tag_name)

            # Sort by part number (Part1, Part2, Part3, ...)
            sorted_parts = sorted(sensors_by_part.keys(), key=lambda x: int(x.replace('Part', '')))

            # Write sensor rows
            for part_name in sorted_parts:
                sensors = sorted(sensors_by_part[part_name])
                for sensor_name in sensors:
                    row_data = [
                        id_counter,              # ID
                        'Common',                # Style
                        'Start Conditions',      # Row Type
                        'SensorUnits',           # Actor Type
                        None,                    # Actor/Unit (empty)
                        None,                    # Custom Description (empty)
                        None,                    # Custom Duration (empty)
                        'OFF',                   # Standard Action/Status (sensors start OFF)
                        None,                    # Standard Duration (empty)
                        part_name,               # Actor/Group Description (Part1, Part2, etc.)
                        sensor_name,             # Actor/Group Name
                        None                     # Valve Name (empty)
                    ]

                    for col_num, value in enumerate(row_data, 1):
                        cell = ws.cell(row=row_num, column=col_num, value=value)
                        cell.fill = self.data_fill
                    row_num += 1
                    id_counter += 1

        # Auto-adjust column widths
        self._adjust_column_widths(ws)

    def _extract_mm_number(self, actuator_name: str) -> int:
        """
        Extract MM number from actuator name for sorting.

        Args:
            actuator_name: Actuator name like 'MM1_MMB1', 'MM12_MMB3'

        Returns:
            MM number as integer (e.g., 1, 12)
        """
        import re
        match = re.match(r'MM(\d+)_', actuator_name)
        return int(match.group(1)) if match else 999

    def _extract_mm_number_from_key(self, mm_number: str) -> int:
        """
        Extract MM number from MM tag name for sorting.

        Args:
            mm_number: MM tag name like 'MM1', 'MM12'

        Returns:
            MM number as integer (e.g., 1, 12)
        """
        import re
        match = re.match(r'MM(\d+)', mm_number)
        return int(match.group(1)) if match else 999

    def _extract_kj_name(self, manifold: str) -> str:
        """
        Extract KJ{x} pattern from manifold name.

        Args:
            manifold: Manifold name (e.g., '_010_UA1_KJ1_Hw', '_010UA1KJ1_KEB1_Hw')

        Returns:
            KJ name (e.g., 'KJ1') or 'N/A' if not found
        """
        import re
        if not manifold:
            return 'N/A'

        match = re.search(r'KJ(\d{1,2})', manifold)
        return f"KJ{match.group(1)}" if match else 'N/A'

    def _build_valve_diagram_name(self, kj_name: str, valve_position: str) -> str:
        """
        Build electrical diagram valve nomenclature.

        Args:
            kj_name: Controls manifold name (e.g., 'KJ1')
            valve_position: Valve position (e.g., '1A', '2B')

        Returns:
            Diagram name (e.g., '=KJ1-QMB1') or 'N/A'
        """
        if kj_name == 'N/A' or not valve_position or valve_position == 'N/A':
            return 'N/A'

        # Extract valve number from '1A' -> '1' or '2B' -> '2'
        valve_num = valve_position.rstrip('AB')

        # Prefix with single quote to force Excel to treat as text (not formula)
        return f"'={kj_name}-QMB{valve_num}"

    def _adjust_column_widths(self, ws):
        """
        Auto-adjust column widths based on content.

        Args:
            ws: Worksheet object
        """
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass

            adjusted_width = min(max_length + DEFAULT_COLUMN_PADDING, MAX_COLUMN_WIDTH)
            ws.column_dimensions[column_letter].width = adjusted_width
