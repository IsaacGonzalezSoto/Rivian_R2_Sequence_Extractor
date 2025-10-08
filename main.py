"""
Interactive L5X extraction system using Typer and Rich.
Discovers L5X files and provides user-friendly file selection.
"""
import sys
import shutil
import time
from pathlib import Path
from typing import List, Optional
from datetime import datetime

import typer
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.prompt import Prompt, Confirm
from rich.progress import Progress, SpinnerColumn, TextColumn, BarColumn, TimeElapsedColumn
from rich import box

from src.pipeline.extraction_pipeline import ExtractionPipeline
from src.core.logger import get_logger

# Initialize
app = typer.Typer(help="Interactive L5X file extraction system")
console = Console()
logger = get_logger(__name__)

# Constants
OUTPUT_BASE_DIR = "output"
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
    Display numbered list of L5X files using Rich table.

    Args:
        files: List of L5X file paths
    """
    table = Table(title=f"Found {len(files)} L5X file(s)", box=box.ROUNDED)
    table.add_column("#", style="cyan", justify="right")
    table.add_column("Filename", style="green")
    table.add_column("Size", style="yellow", justify="right")

    for idx, file in enumerate(files, 1):
        size_mb = file.stat().st_size / (1024 * 1024)
        table.add_row(str(idx), file.name, f"{size_mb:.2f} MB")

    console.print(table)


def get_user_choice(num_files: int) -> str:
    """
    Display menu and get user's processing choice.

    Args:
        num_files: Number of available files

    Returns:
        User's choice as uppercase letter
    """
    console.print("\n[bold]What would you like to process?[/bold]")
    console.print(f"  [cyan][A][/cyan] All files ({num_files} file(s))")
    console.print("  [cyan][O][/cyan] One file")
    console.print("  [cyan][S][/cyan] Select multiple files")
    console.print("  [cyan][Q][/cyan] Quit")

    while True:
        try:
            choice = Prompt.ask("\nYour choice", choices=["A", "O", "S", "Q", "a", "o", "s", "q"]).upper()
            return choice
        except (EOFError, KeyboardInterrupt):
            console.print("\n[yellow]Exiting...[/yellow]")
            raise typer.Exit(0)


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
            choice = Prompt.ask(f"\nEnter file number", default="1")
            file_num = int(choice)
            if 1 <= file_num <= num_files:
                return file_num - 1
            console.print(f"[red]Please enter a number between 1 and {num_files}.[/red]")
        except ValueError:
            console.print("[red]Invalid input. Please enter a number.[/red]")
        except (EOFError, KeyboardInterrupt):
            console.print("\n[yellow]Exiting...[/yellow]")
            raise typer.Exit(0)


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
            choice = Prompt.ask("\nEnter file numbers (comma-separated, e.g., 1,3,5)")
            file_nums = [int(x.strip()) for x in choice.split(',')]

            # Validate all numbers are in range
            if all(1 <= num <= num_files for num in file_nums):
                # Remove duplicates and convert to 0-based indices
                return sorted(list(set(num - 1 for num in file_nums)))

            console.print(f"[red]All numbers must be between 1 and {num_files}.[/red]")
        except ValueError:
            console.print("[red]Invalid input. Please enter numbers separated by commas.[/red]")
        except (EOFError, KeyboardInterrupt):
            console.print("\n[yellow]Exiting...[/yellow]")
            raise typer.Exit(0)


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

    console.print(f"\n[yellow]⚠ Warning: Output folder already exists:[/yellow] [cyan]{output_path}[/cyan]")
    console.print("  [cyan][O][/cyan] Overwrite")
    console.print("  [cyan][S][/cyan] Skip this file")
    console.print("  [cyan][T][/cyan] Create new folder with timestamp")

    while True:
        try:
            choice = Prompt.ask("Your choice", choices=["O", "S", "T", "o", "s", "t"]).upper()
            return {'O': 'overwrite', 'S': 'skip', 'T': 'timestamp'}[choice]
        except (EOFError, KeyboardInterrupt):
            return 'skip'


def process_file(
    l5x_file: Path,
    output_base: Path,
    debug: bool = True
) -> tuple[bool, Optional[Path], float, Optional[str]]:
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
    Process selected L5X files with Rich progress display.

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

    console.print(f"\n[bold]Processing {total_files} file(s)...[/bold]")
    console.rule()

    for count, idx in enumerate(selected_indices, 1):
        l5x_file = files[idx]
        console.print(f"\n[bold cyan][{count}/{total_files}][/bold cyan] Processing [green]{l5x_file.name}[/green]...")

        success, output_folder, elapsed_time, error_msg = process_file(
            l5x_file,
            output_base,
            debug=debug
        )

        if success:
            console.print(f"  [green]✓[/green] Output: [cyan]{output_folder}/[/cyan]")
            console.print(f"  [green]✓[/green] Complete ([yellow]{elapsed_time:.2f}s[/yellow])")
            successful += 1
        elif error_msg and "skip" in error_msg.lower():
            console.print(f"  [yellow]−[/yellow] Skipped ([yellow]{elapsed_time:.2f}s[/yellow])")
            skipped += 1
        else:
            console.print(f"  [red]✗[/red] Failed: {error_msg or 'Unknown error'}")
            console.print(f"  [red]✗[/red] Failed ([yellow]{elapsed_time:.2f}s[/yellow])")
            failed += 1

    # Final summary
    total_elapsed = time.time() - total_start_time

    summary_table = Table(title="Processing Summary", box=box.DOUBLE)
    summary_table.add_column("Metric", style="cyan")
    summary_table.add_column("Value", style="bold", justify="right")

    summary_table.add_row("Total files", str(total_files))
    summary_table.add_row("✓ Successful", f"[green]{successful}[/green]")
    if skipped > 0:
        summary_table.add_row("− Skipped", f"[yellow]{skipped}[/yellow]")
    if failed > 0:
        summary_table.add_row("✗ Failed", f"[red]{failed}[/red]")
    summary_table.add_row("Total time", f"{total_elapsed:.2f} seconds")

    console.print()
    console.print(summary_table)


@app.command()
def main(
    debug: bool = typer.Option(
        False,
        "--debug/--no-debug",
        help="Enable or disable debug output"
    ),
    output_dir: str = typer.Option(
        OUTPUT_BASE_DIR,
        "--output-dir",
        help="Base output directory"
    ),
    no_pause: bool = typer.Option(
        False,
        "--no-pause",
        help="Do not pause before exit (useful for automation)"
    )
):
    """
    Interactive L5X file extraction system.

    Discovers L5X files in the current directory and extracts sequences,
    actions, actuators, transitions, and digital inputs to Excel reports.
    """
    try:
        # Display header
        console.print(Panel.fit(
            "[bold cyan]L5X SEQUENCE EXTRACTOR[/bold cyan]",
            border_style="cyan"
        ))

        # Convert output directory to Path
        output_base = Path(output_dir)

        # Check write permissions
        if not check_write_permission(output_base):
            console.print(f"\n[red]✗ Error: No write permission for output directory:[/red] [cyan]{output_base}[/cyan]")
            if not no_pause and sys.stdin.isatty():
                Prompt.ask("\nPress Enter to exit")
            raise typer.Exit(1)

        # Check disk space
        if not check_disk_space(output_base):
            console.print(f"\n[yellow]⚠ Warning: Low disk space (< {MIN_FREE_SPACE_MB} MB)[/yellow]")
            try:
                if not Confirm.ask("Continue anyway?", default=False):
                    raise typer.Exit(0)
            except (EOFError, KeyboardInterrupt):
                console.print("\n[yellow]Exiting...[/yellow]")
                raise typer.Exit(0)

        # Find all L5X files in current directory
        l5x_files = find_l5x_files()

        if not l5x_files:
            console.print("\n[red]✗ No L5X files found in current directory.[/red]")
            console.print("Please ensure .L5X files are in the same folder as this program.")
            if not no_pause and sys.stdin.isatty():
                Prompt.ask("\nPress Enter to exit")
            raise typer.Exit(0)

        # Display available files
        display_file_list(l5x_files)

        # Get user choice
        choice = get_user_choice(len(l5x_files))

        if choice == 'Q':
            console.print("\n[yellow]Exiting...[/yellow]")
            raise typer.Exit(0)

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
        process_files(l5x_files, selected_indices, output_base, debug=debug)
        console.print("\n[bold green]✓ All processing complete![/bold green]")

        if not no_pause and sys.stdin.isatty():
            Prompt.ask("\nPress Enter to exit")

    except KeyboardInterrupt:
        console.print("\n\n[yellow]⚠ Processing interrupted by user.[/yellow]")
        if not no_pause and sys.stdin.isatty():
            try:
                Prompt.ask("\nPress Enter to exit")
            except (EOFError, KeyboardInterrupt):
                pass
        raise typer.Exit(1)
    except typer.Exit:
        raise
    except Exception as e:
        console.print(f"\n[red]✗ Unexpected error: {str(e)}[/red]")
        logger.exception("Unexpected error in main")
        if not no_pause and sys.stdin.isatty():
            try:
                Prompt.ask("\nPress Enter to exit")
            except (EOFError, KeyboardInterrupt):
                pass
        raise typer.Exit(1)


if __name__ == "__main__":
    app()
