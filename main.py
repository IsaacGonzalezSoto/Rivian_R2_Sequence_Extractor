"""
Interactive L5X extraction system.
Discovers L5X files and provides user-friendly file selection.
"""
import sys
import os
import time
from pathlib import Path
from typing import List, Tuple
from src.pipeline.extraction_pipeline import ExtractionPipeline


def find_l5x_files() -> List[Path]:
    """
    Find all .L5X files in the current directory (case-insensitive).
    
    Returns:
        List of Path objects for found L5X files
    """
    current_dir = Path.cwd()
    l5x_files = list(current_dir.glob("*.[lL]5[xX]"))
    return sorted(l5x_files)


def display_file_list(files: List[Path]) -> None:
    """
    Display numbered list of L5X files.
    
    Args:
        files: List of L5X file paths
    """
    print("\nFound {} L5X file(s):".format(len(files)))
    for idx, file in enumerate(files, 1):
        print(f"  [{idx}] {file.name}")


def get_user_choice(num_files: int) -> str:
    """
    Display menu and get user's processing choice.
    
    Args:
        num_files: Number of available files
        
    Returns:
        User's choice as uppercase letter
    """
    print("\nWhat would you like to process?")
    print(f"  [A] All files ({num_files} file(s))")
    print("  [O] One file")
    print("  [S] Select multiple files")
    print("  [Q] Quit")
    
    while True:
        choice = input("\nYour choice: ").strip().upper()
        if choice in ['A', 'O', 'S', 'Q']:
            return choice
        print("Invalid choice. Please enter A, O, S, or Q.")


def get_single_file_selection(num_files: int) -> int:
    """
    Get user selection for a single file.
    
    Args:
        num_files: Total number of available files
        
    Returns:
        Selected file index (0-based)
    """
    while True:
        try:
            choice = input(f"\nEnter file number (1-{num_files}): ").strip()
            file_num = int(choice)
            if 1 <= file_num <= num_files:
                return file_num - 1
            print(f"Please enter a number between 1 and {num_files}.")
        except ValueError:
            print("Invalid input. Please enter a number.")


def get_multiple_file_selection(num_files: int) -> List[int]:
    """
    Get user selection for multiple files.
    
    Args:
        num_files: Total number of available files
        
    Returns:
        List of selected file indices (0-based)
    """
    while True:
        try:
            choice = input(f"\nEnter file numbers (comma-separated, e.g., 1,3,5): ").strip()
            file_nums = [int(x.strip()) for x in choice.split(',')]
            
            # Validate all numbers are in range
            if all(1 <= num <= num_files for num in file_nums):
                # Remove duplicates and convert to 0-based indices
                return sorted(list(set(num - 1 for num in file_nums)))
            
            print(f"All numbers must be between 1 and {num_files}.")
        except ValueError:
            print("Invalid input. Please enter numbers separated by commas.")


def get_output_folder_name(l5x_file: Path) -> str:
    """
    Generate output folder name from L5X filename.
    
    Args:
        l5x_file: Path to L5X file
        
    Returns:
        Output folder name (without .L5X extension)
    """
    return l5x_file.stem


def process_file(l5x_file: Path, output_base: str = "output", debug: bool = True) -> Tuple[bool, str, float]:
    """
    Process a single L5X file.
    
    Args:
        l5x_file: Path to L5X file
        output_base: Base output directory
        debug: Enable debug output
        
    Returns:
        Tuple of (success, output_folder, elapsed_time)
    """
    start_time = time.time()
    
    # Create specific output folder for this file
    output_folder = os.path.join(output_base, get_output_folder_name(l5x_file))
    
    try:
        pipeline = ExtractionPipeline(
            l5x_file_path=str(l5x_file),
            output_folder=output_folder,
            debug=debug
        )
        
        pipeline.run()
        
        elapsed_time = time.time() - start_time
        return True, output_folder, elapsed_time
        
    except Exception as e:
        elapsed_time = time.time() - start_time
        print(f"      Error: {str(e)}")
        return False, output_folder, elapsed_time


def process_files(files: List[Path], selected_indices: List[int], debug: bool = True) -> None:
    """
    Process selected L5X files.
    
    Args:
        files: List of all available L5X files
        selected_indices: Indices of files to process
        debug: Enable debug output
    """
    total_files = len(selected_indices)
    successful = 0
    failed = 0
    total_start_time = time.time()
    
    print(f"\nProcessing {total_files} file(s)...")
    print("="*60)
    
    for count, idx in enumerate(selected_indices, 1):
        l5x_file = files[idx]
        print(f"\n[{count}/{total_files}] Processing {l5x_file.name}...")
        
        success, output_folder, elapsed_time = process_file(l5x_file, debug=debug)
        
        if success:
            print(f"      Output: {output_folder}/")
            print(f"      Complete ({elapsed_time:.2f}s)")
            successful += 1
        else:
            print(f"      Failed ({elapsed_time:.2f}s)")
            failed += 1
    
    # Final summary
    total_elapsed = time.time() - total_start_time
    print("\n" + "="*60)
    print("PROCESSING SUMMARY")
    print("="*60)
    print(f"Total files processed: {total_files}")
    print(f"  Successful: {successful}")
    if failed > 0:
        print(f"  Failed: {failed}")
    print(f"Total execution time: {total_elapsed:.2f} seconds")
    print("="*60)


def main():
    """
    Main program function with interactive file selection.
    """
    print("="*60)
    print("L5X SEQUENCE EXTRACTOR")
    print("="*60)
    
    # Find all L5X files in current directory
    l5x_files = find_l5x_files()
    
    if not l5x_files:
        print("\nNo L5X files found in current directory.")
        print("Please ensure .L5X files are in the same folder as this program.")
        print("\nPress Enter to exit...")
        input()
        sys.exit(0)
    
    # Display available files
    display_file_list(l5x_files)
    
    # Get user choice
    choice = get_user_choice(len(l5x_files))
    
    if choice == 'Q':
        print("\nExiting...")
        sys.exit(0)
    
    # Determine which files to process
    selected_indices = []
    
    if choice == 'A':
        # Process all files
        selected_indices = list(range(len(l5x_files)))
    
    elif choice == 'O':
        # Process one file
        idx = get_single_file_selection(len(l5x_files))
        selected_indices = [idx]
    
    elif choice == 'S':
        # Process multiple selected files
        selected_indices = get_multiple_file_selection(len(l5x_files))
    
    # Process the selected files
    try:
        process_files(l5x_files, selected_indices, debug=True)
        print("\nAll processing complete!")
        print("\nPress Enter to exit...")
        input()
        
    except KeyboardInterrupt:
        print("\n\nProcessing interrupted by user.")
        print("\nPress Enter to exit...")
        input()
        sys.exit(1)
    except Exception as e:
        print(f"\nUnexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        print("\nPress Enter to exit...")
        input()
        sys.exit(1)


if __name__ == "__main__":
    main()