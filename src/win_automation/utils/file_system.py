"""File system utilities for Windows automation."""

import shutil
from pathlib import Path


class FileSystemUtils:
    """Utility class for file system operations."""

    @staticmethod
    def ensure_directory(path: str | Path) -> Path:
        """Ensure a directory exists, creating it if necessary.

        Args:
            path: Directory path to ensure

        Returns:
            Path object of the directory
        """
        dir_path = Path(path)
        dir_path.mkdir(parents=True, exist_ok=True)
        return dir_path

    @staticmethod
    def copy_file(source: str | Path, destination: str | Path) -> bool:
        """Copy a file from source to destination.

        Args:
            source: Source file path
            destination: Destination file path

        Returns:
            True if successful, False otherwise
        """
        try:
            shutil.copy2(source, destination)
            return True
        except (OSError, shutil.Error):
            return False

    @staticmethod
    def move_file(source: str | Path, destination: str | Path) -> bool:
        """Move a file from source to destination.

        Args:
            source: Source file path
            destination: Destination file path

        Returns:
            True if successful, False otherwise
        """
        try:
            shutil.move(str(source), str(destination))
            return True
        except (OSError, shutil.Error):
            return False

    @staticmethod
    def find_files(
        directory: str | Path,
        pattern: str = "*",
        recursive: bool = True
    ) -> list[Path]:
        """Find files matching a pattern in a directory.

        Args:
            directory: Directory to search in
            pattern: File pattern to match (e.g., "*.py", "*.txt")
            recursive: Whether to search recursively

        Returns:
            List of matching file paths
        """
        dir_path = Path(directory)
        if not dir_path.exists():
            return []

        if recursive:
            return list(dir_path.rglob(pattern))
        else:
            return list(dir_path.glob(pattern))

    @staticmethod
    def get_file_size(file_path: str | Path) -> int | None:
        """Get the size of a file in bytes.

        Args:
            file_path: Path to the file

        Returns:
            File size in bytes, or None if file doesn't exist
        """
        path = Path(file_path)
        try:
            return path.stat().st_size
        except OSError:
            return None
