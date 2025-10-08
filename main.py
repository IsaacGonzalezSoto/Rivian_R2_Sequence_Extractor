"""
Interactive L5X extraction system.
Discovers L5X files and provides user-friendly file selection.
"""
import sys
import os
import shutil
import argparse
import time
from pathlib import Path
from typing import List, Tuple, Optional
from datetime import datetime
from src.pipeline.extraction_pipeline import ExtractionPipeline
from src.core.logger import get_logger

# Initialize logger
logger = get_logger(__name__)

# Constants
OUTPUT_BASE_DIR = "output"
DEBUG_DEFAULT = False
MIN_FREE_SPACE_MB = 100


def find_l5x_files() -> List[Path]:
    """
    Find all .L5X files in the current directory (case-insensitive).
    
    Returns:
        List of Path objects for found L5X files
    """
    current_dir = Path.cwd()
    l5x_files = [
        f for f in current_dir.iterdir()
        if f.is_file() and f.suffix.lower() == '.l5x'
    ]
    return sorted(l5x_files)


def check_disk_space(path: Path, required_mb: int = MIN_FREE_SPACE_MB) -> bool:
    """
    Check if sufficient disk space is available.
    
    Args:
        path: Path to check disk space for
        required_mb: Minimum required space in MB
        
    Returns:
        True if sufficient space available
    """
    try:
        stat = shutil.disk_usage(path)
        free_mb = stat.free / (1024 * 1024)
        return free_mb >= required_mb
    except Exception as e:
        logger.warning(f"Could not check disk space: {e}")
        return True  # Proceed anyway if check fails


def check_write_permission(path: Path) -> bool:
    """
    Check if we have write permission to the directory.
    
    Args:
        path: Directory path to check
        
    Returns:
        True if writable
    """
    try:
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)
        test_file = path / '.write_test'
        test_file.touch()
        test_file.unlink()
        return True
    except Exception as e:
        logger.error(f"No write permission for {path}: {e}")
        return False


def display_file_list(files: List[Path]) -> None:
    """
    Display numbered list of L5X files.
    
    Args:
        files: List of L5X file paths
    """
    print(f"\nFound {len(files)} L5X file(s):")
    for idx, file in enumerate(files, 1):
        size_mb = file.stat().st_size / (1024 * 1024)
        print(f"  [{idx}] {file.name} ({size_mb:.2f} MB)")


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
        try:
            choice = input("\nYour choice: ").strip().upper()
            if choice in ['A', 'O', 'S', 'Q']:
                return choice
            print("Invalid choice. Please enter A, O, S, or Q.")
        except (EOFError, KeyboardInterrupt):
            print("\n\nExiting...")
            sys.exit(0)


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
        except (EOFError, KeyboardInterrupt):
            print("\n\nExiting...")
            sys.exit(0)


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
        except (EOFError, KeyboardInterrupt):
            print("\n\nExiting...")
            sys.exit(0)


def get_output_folder_path(l5x_file: Path, output_base: Path, add_timestamp: bool = False) -> Path:
    """
    Generate output folder path from L5X filename.
    
    Args:
        l5x_file: Path to L5X file
        output_base: Base output directory
        add_timestamp: Whether to add timestamp to folder name
        
    Returns:
        Complete output folder path
    """
    folder_name = l5x_file.stem
    
    if add_timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder_name = f"{folder_name}_{timestamp}"
    
    return output_base / folder_name


def check_output_folder_exists(output_path: Path) -> Optional[str]:
    """
    Check if output folder exists and get user decision.
    
    Args:
        output_path: Path to check
        
    Returns:
        'overwrite', 'skip', or 'timestamp' based on user choice
    """
    if not output_path.exists():
        return None
    
    print(f"\n      Warning: Output folder already exists: {output_path}")
    print("      [O] Overwrite")
    print("      [S] Skip this file")
    print("      [T] Create new folder with timestamp")
    
    while True:
        try:
            choice = input("      Your choice: ").strip().upper()
            if choice in ['O', 'S', 'T']:
                return {'O': 'overwrite', 'S': 'skip', 'T': 'timestamp'}[choice]
            print("      Invalid choice. Please enter O, S, or T.")
        except (EOFError, KeyboardInterrupt):
            return 'skip'


def process_file(
    l5x_file: Path,
    output_base: Path,
    debug: bool = True
) -> Tuple[bool, Optional[Path], float, Optional[str]]:
    """
    Process a single L5X file.
    
    Args:
        l5x_file: Path to L5X file
        output_base: Base output directory
        debug: Enable debug output
        
    Returns:
        Tuple of (success, output_folder, elapsed_time, skip_reason)
    """
    start_time = time.time()
    
    # Validate file still exists
    if not l5x_file.exists():
        elapsed_time = time.time() - start_time
        return False, None, elapsed_time, "File no longer exists"
    
    # Create output folder path
    output_folder = get_output_folder_path(l5x_file, output_base)
    
    # Check if folder exists and handle collision
    collision_action = check_output_folder_exists(output_folder)
    
    if collision_action == 'skip':
        elapsed_time = time.time() - start_time
        return False, output_folder, elapsed_time, "Skipped by user"
    
    if collision_action == 'timestamp':
        output_folder = get_output_folder_path(l5x_file, output_base, add_timestamp=True)
    
    if collision_action == 'overwrite' and output_folder.exists():
        try:
            shutil.rmtree(output_folder)
        except Exception as e:
            elapsed_time = time.time() - start_time
            logger.error(f"Could not remove existing folder: {e}")
            return False, output_folder, elapsed_time, f"Could not overwrite: {e}"
    
    try:
        pipeline = ExtractionPipeline(
            l5x_file_path=str(l5x_file),
            output_folder=str(output_folder),
            debug=debug
        )
        
        pipeline.run()
        
        elapsed_time = time.time() - start_time
        return True, output_folder, elapsed_time, None
        
    except Exception as e:
        elapsed_time = time.time() - start_time
        logger.error(f"Error processing file: {str(e)}", exc_info=True)
        return False, output_folder, elapsed_time, str(e)


def process_files(
    files: List[Path],
    selected_indices: List[int],
    output_base: Path,
    debug: bool = True
) -> None:
    """
    Process selected L5X files.
    
    Args:
        files: List of all available L5X files
        selected_indices: Indices of files to process
        output_base: Base output directory
        debug: Enable debug output
    """
    total_files = len(selected_indices)
    successful = 0
    failed = 0
    skipped = 0
    total_start_time = time.time()
    
    print(f"\nProcessing {total_files} file(s)...")
    print("="*60)
    
    for count, idx in enumerate(selected_indices, 1):
        l5x_file = files[idx]
        print(f"\n[{count}/{total_files}] Processing {l5x_file.name}...")
        
        success, output_folder, elapsed_time, error_msg = process_file(
            l5x_file,
            output_base,
            debug=debug
        )
        
        if success:
            print(f"      ✓ Output: {output_folder}/")
            print(f"      ✓ Complete ({elapsed_time:.2f}s)")
            successful += 1
        elif error_msg and "skip" in error_msg.lower():
            print(f"      - Skipped ({elapsed_time:.2f}s)")
            skipped += 1
        else:
            print(f"      ✗ Failed: {error_msg or 'Unknown error'}")
            print(f"      ✗ Failed ({elapsed_time:.2f}s)")
            failed += 1
    
    # Final summary
    total_elapsed = time.time() - total_start_time
    print("\n" + "="*60)
    print("PROCESSING SUMMARY")
    print("="*60)
    print(f"Total files: {total_files}")
    print(f"  ✓ Successful: {successful}")
    if skipped > 0:
        print(f"  - Skipped: {skipped}")
    if failed > 0:
        print(f"  ✗ Failed: {failed}")
    print(f"Total execution time: {total_elapsed:.2f} seconds")
    print("="*60)


def pause_if_interactive(args: argparse.Namespace) -> None:
    """
    Pause for user input if running interactively.
    
    Args:
        args: Parsed command line arguments
    """
    if not args.no_pause and sys.stdin.isatty():
        try:
            input("\nPress Enter to exit...")
        except (EOFError, KeyboardInterrupt):
            pass


def parse_arguments() -> argparse.Namespace:
    """
    Parse command line arguments.
    
    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description="Interactive L5X file extraction system"
    )
    parser.add_argument(
        '--debug',
        action='store_true',
        default=DEBUG_DEFAULT,
        help='Enable debug output'
    )
    parser.add_argument(
        '--no-debug',
        action='store_false',
        dest='debug',
        help='Disable debug output'
    )
    parser.add_argument(
        '--output-dir',
        type=str,
        default=OUTPUT_BASE_DIR,
        help=f'Base output directory (default: {OUTPUT_BASE_DIR})'
    )
    parser.add_argument(
        '--no-pause',
        action='store_true',
        help='Do not pause before exit (useful for automation)'
    )
    
    return parser.parse_args()


def main():
    """
    Main program function with interactive file selection.
    """
    args = parse_arguments()
    
    print("="*60)
    print("L5X SEQUENCE EXTRACTOR")
    print("="*60)
    
    # Convert output directory to Path
    output_base = Path(args.output_dir)
    
    # Check write permissions
    if not check_write_permission(output_base):
        print(f"\n✗ Error: No write permission for output directory: {output_base}")
        pause_if_interactive(args)
        sys.exit(1)
    
    # Check disk space
    if not check_disk_space(output_base):
        print(f"\n⚠ Warning: Low disk space (< {MIN_FREE_SPACE_MB} MB)")
        try:
            choice = input("Continue anyway? [y/N]: ").strip().lower()
            if choice != 'y':
                sys.exit(0)
        except (EOFError, KeyboardInterrupt):
            print("\nExiting...")
            sys.exit(0)
    
    # Find all L5X files in current directory
    l5x_files = find_l5x_files()
    
    if not l5x_files:
        print("\n✗ No L5X files found in current directory.")
        print("Please ensure .L5X files are in the same folder as this program.")
        pause_if_interactive(args)
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
        process_files(l5x_files, selected_indices, output_base, debug=args.debug)
        print("\n✓ All processing complete!")
        pause_if_interactive(args)
        
    except KeyboardInterrupt:
        print("\n\n⚠ Processing interrupted by user.")
        pause_if_interactive(args)
        sys.exit(1)
    except Exception as e:
        print(f"\n✗ Unexpected error: {str(e)}")
        logger.exception("Unexpected error in main")
        pause_if_interactive(args)
        sys.exit(1)


if __name__ == "__main__":
    main()