"""
Main entry point for the L5X extraction system.
"""
import sys
import os
from src.pipeline.extraction_pipeline import ExtractionPipeline


def print_summary(routines_data):
    """
    Print a summary of extraction results.
    
    Args:
        routines_data: List of processed routines
    """
    print("\n" + "="*60)
    print("FINAL SUMMARY")
    print("="*60)
    
    for routine_info in routines_data:
        print(f"\n📄 {routine_info['routine_name']}")
        print(f"   Excel File: {routine_info['excel_file']}")
        print(f"   - Sequences: {routine_info['sequences_count']}")
        print(f"   - Transitions: {routine_info['transitions_count']}")


def main():
    """
    Main program function.
    """
    # Configuration
    l5x_file = "_010UA1_Fixture_Em0105_Program.L5X"
    output_folder = "output"
    debug = True
    
    # Validate that the file exists
    if not os.path.exists(l5x_file):
        print(f"❌ Error: File not found {l5x_file}")
        print("Please update the L5X file path.")
        sys.exit(1)
    
    try:
        # Create and execute pipeline
        pipeline = ExtractionPipeline(
            l5x_file_path=l5x_file,
            output_folder=output_folder,
            debug=debug
        )
        
        # Execute extraction
        routines_data = pipeline.run()
        
        # Show summary
        print_summary(routines_data)
        
        print("\n✅ Process completed successfully")
        
    except FileNotFoundError as e:
        print(f"❌ Error: File not found - {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error during processing: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()