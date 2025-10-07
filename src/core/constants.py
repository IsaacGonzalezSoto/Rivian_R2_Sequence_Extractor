"""
Constants and configuration values for R2_Sequence_Extractor.
Centralizes magic values, patterns, and styling configurations.
"""

# ============================================================================
# REGEX PATTERNS
# ============================================================================

# Sequence and action patterns
PATTERN_ACTION_ASSIGNMENT = r'EmSeqList\[(\d+)\]\.Step\[(\d+)\]\.ActionNumber\[(\d+)\]\s*:=\s*(\w+)\.outActionNum'
PATTERN_SEQUENCE_REGION = r'#region\s+Sequence\s+(\d+)\s+-\s+(.+)'
PATTERN_SEQUENCE_NAME = r"EmSeqList\[(\d+)\]\.Name\s*:=\s*'([^']+)'"
PATTERN_ACTION_NAME = r'Action(MM\d+)(\w+)'

# Transition patterns
PATTERN_TRANSITION_PERMISSION = r'EmTransitionStates\[(\d+)\]\.AutoStartPerms\.(\d+)\s*:=\s*([^;]+);\s*(?://(.*))?'
PATTERN_TRANSITION_REGION = r'#region\s+Transition\s+State\s+(\d+)\s+-\s+(.+)'

# Actuator patterns
PATTERN_ACTUATOR_MOVE = r"MOVE\('([^']+)',\s*{mm_number}Cyls\[(\d+)\]\.Stg\.Name\)"
PATTERN_MM_NUMBER = r'(MM\d+)'
PATTERN_MM_ROUTINE = r'_{mm_number}$|_{mm_number}_'

# ============================================================================
# ROUTINE PREFIXES
# ============================================================================

ROUTINE_PREFIX_SEQUENCES = 'EmStatesAndSequences'

# ============================================================================
# TAG DATA TYPES
# ============================================================================

TAG_TYPE_DIGITAL_INPUT = 'UDT_DigitalInputHal'

# ============================================================================
# EXCEL STYLING COLORS (HEX without #)
# ============================================================================

class ExcelColors:
    """Excel cell colors for styling sheets."""

    # Header colors
    HEADER_FILL = "366092"
    HEADER_FONT = "FFFFFF"

    # Complete Flow sheet colors
    TRANSITION_FILL = "4472C4"
    TRANSITION_FONT = "FFFFFF"
    SEQUENCE_FILL = "70AD47"
    SEQUENCE_FONT = "FFFFFF"
    STEP_FILL = "FFC000"
    STEP_FONT = "000000"
    ACTION_FILL = "E7E6E6"
    ACTION_FONT = "000000"

    # Data cell background - soft beige for eye comfort
    DATA_FILL = "F5F5DC"

# ============================================================================
# EXCEL FONT SIZES
# ============================================================================

class ExcelFontSizes:
    """Excel font sizes for different elements."""

    HEADER = 11
    TRANSITION = 12
    SEQUENCE = 11
    STEP = 10
    ACTION = 10
    DATA = 10

# ============================================================================
# EXCEL CONFIGURATION
# ============================================================================

MAX_COLUMN_WIDTH = 50  # Maximum column width in characters
DEFAULT_COLUMN_PADDING = 2  # Extra padding for column width

# ============================================================================
# OUTPUT CONFIGURATION
# ============================================================================

DEFAULT_OUTPUT_FOLDER = 'output'
EXCEL_FILE_PREFIX = 'complete_'
EXCEL_FILE_EXTENSION = '.xlsx'

# Sheet names
SHEET_COMPLETE_FLOW = 'Complete_Flow'
SHEET_SEQUENCES_ACTUATORS = 'Sequences_Actuators'
SHEET_TRANSITIONS = 'Transitions'
SHEET_DIGITAL_INPUTS = 'Digital_Inputs'
