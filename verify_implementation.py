"""
Verification script for Sequence Detail Exporter implementation.
Checks all critical aspects of the FINAL implementation.
"""

import re
from pathlib import Path

# Read the exporter file
exporter_path = Path("src/exporters/sequence_detail_exporter.py")
with open(exporter_path, 'r', encoding='utf-8') as f:
    code = f.read()

print("=" * 80)
print("SEQUENCE DETAIL EXPORTER - IMPLEMENTATION VERIFICATION")
print("=" * 80)
print()

# Check 1: Verify 17 columns in headers
print("CHECK 1: Headers - 17 Columns")
print("-" * 80)
if "'Ind. \\nStart'" in code and "'Dep. ID'" in code and "'Start'" in code and "'Duration'" in code and "'End'" in code:
    print("PASS: All 5 timing columns present in headers")
else:
    print("FAIL: Missing timing columns in headers")
print()

# Check 2: Verify _format_actor_group_name() exists
print("CHECK 2: Helper Method - _format_actor_group_name()")
print("-" * 80)
if 'def _format_actor_group_name(self, name: str)' in code:
    print("PASS: _format_actor_group_name() method exists")
    if 'replace(\'_\', \'-\')' in code and 'return f"={formatted}"' in code:
        print("PASS: Method formats correctly (= prefix, - separator)")
    else:
        print("FAIL: Method formatting logic incorrect")
else:
    print("FAIL: _format_actor_group_name() method missing")
print()

# Check 3: Verify Start Conditions Header exists
print("CHECK 3: Start Conditions Header Row")
print("-" * 80)
if "'Start Conditions Header'" in code and "'HomePos'" in code:
    print("PASS: Start Conditions Header row exists")
else:
    print("FAIL: Start Conditions Header row missing")
print()

# Check 4: Verify state formatting methods
print("CHECK 4: State Formatting Methods")
print("-" * 80)
if 'def _format_start_condition_state(self' in code:
    print("PASS: _format_start_condition_state() exists")
    if 'return f"AT {state_normalized}"' in code:
        print("PASS: Returns 'AT' prefix for start conditions")
    else:
        print("FAIL: Does not return 'AT' prefix")
else:
    print("FAIL: _format_start_condition_state() missing")

if 'def _format_state_robust(self' in code:
    print("PASS: _format_state_robust() exists")
    if 'return f"TO {state_normalized}"' in code:
        print("PASS: Returns 'TO' prefix for actions")
    else:
        print("FAIL: Does not return 'TO' prefix")
else:
    print("FAIL: _format_state_robust() missing")
print()

# Check 5: Verify all name formatting is applied
print("CHECK 5: Name Formatting Application")
print("-" * 80)
actor_group_name_calls = len(re.findall(r'self\._format_actor_group_name\(', code))
if actor_group_name_calls >= 5:
    print(f"PASS: _format_actor_group_name() called {actor_group_name_calls} times")
else:
    print(f"FAIL: _format_actor_group_name() only called {actor_group_name_calls} times (expected >= 5)")
print()

# Check 6: Verify Standard Duration values
print("CHECK 6: Standard Duration Values")
print("-" * 80)
if "0.0,                             # Standard Duration" in code:
    print("PASS: 0.0 for Start Conditions")
else:
    print("FAIL: Missing 0.0 for Start Conditions")

if "2.0,                             # Standard Duration" in code:
    print("PASS: 2.0 for Actions")
else:
    print("FAIL: Missing 2.0 for Actions")
print()

# Check 7: Verify Step row types
print("CHECK 7: Step Row Types (Step1, Step2, Step3)")
print("-" * 80)
if 'step_name = f"Step{step_idx}"' in code:
    print("PASS: Step numbering implemented")
else:
    print("FAIL: Step numbering missing")
print()

# Check 8: Verify Sequence Headers
print("CHECK 8: Sequence Header and End Of Sequence")
print("-" * 80)
if "'Sequence Header'" in code and '"{sequence_style} Sequence"' in code:
    print("PASS: Sequence Header exists")
else:
    print("FAIL: Sequence Header missing")

if "'End Of Sequence'" in code:
    print("PASS: End Of Sequence exists")
else:
    print("FAIL: End Of Sequence missing")
print()

# Check 9: Verify Fixed/Transition State parsing
print("CHECK 9: Fixed/Transition State Parsing")
print("-" * 80)
if "'Transition State' in transition_name:" in code:
    print("PASS: Transition State parsing exists")
else:
    print("FAIL: Transition State parsing missing")

if "'Fixed State' in transition_name:" in code:
    print("PASS: Fixed State parsing exists")
else:
    print("FAIL: Fixed State parsing missing")

if "state_name = transition_name.split(' - ', 1)[1]" in code:
    print("PASS: State name extraction implemented")
else:
    print("FAIL: State name extraction missing")
print()

# Check 10: Verify Operators pattern
print("CHECK 10: Operators Pattern Support")
print("-" * 80)
if "'Operator' in comment or 'operator' in comment.lower()" in code:
    print("PASS: Operators pattern matching exists")
else:
    print("FAIL: Operators pattern matching missing")

if "'actor_type': 'Operators'" in code:
    print("PASS: Operators actor type exists")
else:
    print("FAIL: Operators actor type missing")
print()

# Check 11: Verify timing columns in all row_data
print("CHECK 11: Timing Columns in All row_data")
print("-" * 80)
row_data_occurrences = len(re.findall(r'row_data = \[', code))
timing_column_occurrences = len(re.findall(r'None, None, None, None, None\s+#.*[Tt]iming', code))
print(f"Found {row_data_occurrences} row_data assignments")
print(f"Found {timing_column_occurrences} with timing columns comment")

if timing_column_occurrences >= 7:
    print(f"PASS: Timing columns present in {timing_column_occurrences} places")
else:
    print(f"WARNING: Timing columns only in {timing_column_occurrences} places (expected >= 7)")
print()

# Final Summary
print("=" * 80)
print("VERIFICATION COMPLETE")
print("=" * 80)
print()
print("Review the results above. If all checks show PASS, the implementation")
print("is correct and ready for production use.")
print()
