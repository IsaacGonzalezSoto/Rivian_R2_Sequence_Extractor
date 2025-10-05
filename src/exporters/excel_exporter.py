"""
Excel exporter for sequences and transitions data.
Uses openpyxl to create .xlsx files with multiple sheets.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from typing import Dict, Any, List


class ExcelExporter:
    """
    Exports data to Excel format (.xlsx) with multiple sheets.
    One file per routine with separate sheets for sequences and transitions.
    """
    
    def __init__(self):
        """Initialize the Excel exporter."""
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.header_font = Font(bold=True, color="FFFFFF")
        self.header_alignment = Alignment(horizontal="center", vertical="center")
    
    def export(self, sequences_data: Dict[str, Any], transitions_data: Dict[str, Any], output_path: str):
        """
        Export sequences and transitions to a single Excel file with multiple sheets.
        
        Args:
            sequences_data: Dictionary with sequence and actuator data
            transitions_data: Dictionary with transition permission data
            output_path: Path for the output Excel file
        """
        wb = Workbook()
        
        # Remove default sheet
        if 'Sheet' in wb.sheetnames:
            wb.remove(wb['Sheet'])
        
        # Create sequences sheet
        self._create_sequences_sheet(wb, sequences_data)
        
        # Create transitions sheet if there's data
        if transitions_data and transitions_data.get('transitions'):
            self._create_transitions_sheet(wb, transitions_data)
        
        # Save workbook
        wb.save(output_path)
    
    def _create_sequences_sheet(self, wb: Workbook, data: Dict[str, Any]):
        """
        Create the Sequences_Actuators sheet.
        
        Args:
            wb: Workbook object
            data: Sequences data
        """
        ws = wb.create_sheet("Sequences_Actuators")
        
        # Headers
        headers = [
            'Routine',
            'Sequence',
            'Step',
            'Action_Index',
            'Action_Name',
            'MM_Number',
            'State',
            'Actuator_Count',
            'Actuators',
            'Validation_Status',
            'Missing_Indices'
        ]
        
        # Write headers
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.header_alignment
        
        # Write data
        row_num = 2
        routine_name = data['routine_name']
        
        for sequence in data['sequences']:
            seq_idx = sequence['sequence_index']
            
            for step in sequence['steps']:
                step_idx = step['step_index']
                
                for action in step['actions']:
                    # Determine validation status
                    validation_status = 'N/A'
                    missing_indices = ''
                    
                    if 'validation' in action:
                        val = action['validation']
                        if val['is_valid'] == True:
                            validation_status = 'OK'
                        elif val['is_valid'] == False:
                            validation_status = 'WARNING'
                            missing_indices = ','.join(map(str, val['missing_indices']))
                        elif val['is_valid'] is None:
                            validation_status = 'UNKNOWN'
                    
                    # Format state
                    state_formatted = ''
                    if action['state']:
                        state_formatted = f"TO {action['state'].upper()}"
                    
                    # Write one row per actuator
                    if not action['actuators']:
                        # No actuators, write one row
                        row_data = [
                            routine_name,
                            seq_idx,
                            step_idx,
                            action['action_index'],
                            action['action_name'],
                            action['mm_number'] or '',
                            state_formatted,
                            action['actuator_count'],
                            '',
                            validation_status,
                            missing_indices
                        ]
                        
                        for col_num, value in enumerate(row_data, 1):
                            ws.cell(row=row_num, column=col_num, value=value)
                        row_num += 1
                    else:
                        # One row per actuator
                        for actuator in action['actuators']:
                            row_data = [
                                routine_name,
                                seq_idx,
                                step_idx,
                                action['action_index'],
                                action['action_name'],
                                action['mm_number'] or '',
                                state_formatted,
                                action['actuator_count'],
                                actuator['description'],
                                validation_status,
                                missing_indices
                            ]
                            
                            for col_num, value in enumerate(row_data, 1):
                                ws.cell(row=row_num, column=col_num, value=value)
                            row_num += 1
        
        # Auto-adjust column widths
        self._adjust_column_widths(ws)
    
    def _create_transitions_sheet(self, wb: Workbook, data: Dict[str, Any]):
        """
        Create the Transitions sheet.
        
        Args:
            wb: Workbook object
            data: Transitions data
        """
        ws = wb.create_sheet("Transitions")
        
        # Headers
        headers = [
            'Routine',
            'Transition_Index',
            'Permission_Count',
            'Permission_Index',
            'Permission_Value',
            'Comment'
        ]
        
        # Write headers
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = self.header_fill
            cell.font = self.header_font
            cell.alignment = self.header_alignment
        
        # Write data
        row_num = 2
        routine_name = data['routine_name']
        
        for transition in data['transitions']:
            trans_idx = transition['transition_index']
            permission_count = transition['permission_count']
            
            for permission in transition['permissions']:
                row_data = [
                    routine_name,
                    trans_idx,
                    permission_count,
                    permission['permission_index'],
                    permission['permission_value'],
                    permission['comment']
                ]
                
                for col_num, value in enumerate(row_data, 1):
                    ws.cell(row=row_num, column=col_num, value=value)
                row_num += 1
        
        # Auto-adjust column widths
        self._adjust_column_widths(ws)
    
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
            
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width