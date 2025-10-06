"""
Main entry point for the L5X extraction system.
Extracts sequences and transitions from Rockwell L5X files and generates Excel reports.
"""
import sys
import os
import time
from src.pipeline.extraction_pipeline import ExtractionPipeline


def print_summary(routines_data, elapsed_time):
    """
    Print a summary of extraction results.
    
    Args:
        routines_data: List of processed routines
        elapsed_time: Total execution time in seconds
    """
    print("\n" + "="*60)
    print("EXTRACTION SUMMARY")
    print("="*60)
    
    for routine_info in routines_data:
        print(f"\n  {routine_info['routine_name']}")
        print(f"   Excel File: {routine_info['excel_file']}")
        print(f"   - Sequences: {routine_info['sequences_count']}")
        print(f"   - Transitions: {routine_info['transitions_count']}")
    
    print("\n" + "="*60)
    print(f"Total execution time: {elapsed_time:.2f} seconds")
    print("="*60)


def main():
    """
    Main program function.
    Processes L5X file and generates Excel output.
    """
    # Start timing
    start_time = time.time()
    
    # Configuration
    l5x_file = "_010UA1_Fixture_Em0105_Program.L5X"
    output_folder = "output"
    debug = True
    
    # Validate that the file exists
    if not os.path.exists(l5x_file):
        print(f"Error: File not found: {l5x_file}")
        print("Please update the L5X file path in main.py")
        sys.exit(1)
    
    try:
        # Create and execute pipeline
        print("="*60)
        print("L5X SEQUENCE EXTRACTOR")
        print("="*60)
        print(f"Input:  {l5x_file}")
        print(f"Output: {output_folder}/")
        print("="*60)
        
        pipeline = ExtractionPipeline(
            l5x_file_path=l5x_file,
            output_folder=output_folder,
            debug=debug
        )
        
        # Execute extraction
        routines_data = pipeline.run()
        
        # Calculate elapsed time
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        # Show summary with timing
        print_summary(routines_data, elapsed_time)
        
        print("\nProcess completed successfully")
        
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error during processing: {str(e)}")
        if debug:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()