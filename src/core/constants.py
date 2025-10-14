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
    """Excel cell colors for styling sheets - Dark theme for eye comfort."""

    # Header colors - Dark theme
    HEADER_FILL = "1A1A1A"  # Almost black
    HEADER_FONT = "00D9FF"  # Bright cyan

    # Complete Flow sheet colors - Dark theme
    TRANSITION_FILL = "1E3A5F"  # Dark navy blue
    TRANSITION_FONT = "66B3FF"  # Light blue
    SEQUENCE_FILL = "1F4E2D"  # Dark green
    SEQUENCE_FONT = "90EE90"  # Light green
    STEP_FILL = "4A3B00"  # Dark gold
    STEP_FONT = "FFD966"  # Light gold
    ACTION_FILL = "3A3A3A"  # Dark gray
    ACTION_FONT = "D0D0D0"  # Light gray

    # Data cell background - Dark theme for eye comfort
    DATA_FILL = "2D2D2D"  # Charcoal gray
    DATA_FONT = "E0E0E0"  # Light gray text

    # Duplicate warning - Dark theme
    DUPLICATE_FILL = "5A1A1A"  # Dark red
    DUPLICATE_FONT = "FF6B6B"  # Light red

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
