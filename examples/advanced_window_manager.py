#!/usr/bin/env python3
"""
Advanced Window Manager - Enhanced pywinauto example
Intelligent window management: layouts, workspace organization, productivity features
"""

import json
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

# Add the src directory to the path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


try:
    import pywinauto
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import find_windows
    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


@dataclass
class WindowInfo:
    """Information about a window."""
    handle: int
    title: str
    process_name: str
    class_name: str
    rect: tuple[int, int, int, int]  # left, top, right, bottom
    is_visible: bool
    is_minimized: bool


@dataclass
class Layout:
    """Window layout configuration."""
    name: str
    windows: list[dict[str, Any]]  # List of window configurations


class AdvancedWindowManager:
    """Advanced window management with intelligent features."""

    def __init__(self):
        """Initialize advanced window manager."""
        if not PYWINAUTO_AVAILABLE:
            raise RuntimeError("pywinauto not available for window management")

        self.desktop = Desktop(backend="uia")
        self.layouts = {}
        self.saved_layouts = {}

    def get_all_windows(self, visible_only: bool = True) -> list[WindowInfo]:
        """Get information about all windows."""
        windows = []

        try:
            # Find all top-level windows
            window_handles = find_windows()

            for handle in window_handles:
                try:
                    window = self.desktop.window(handle=handle)

                    # Skip if not visible and we only want visible windows
                    if visible_only and not window.is_visible():
                        continue

                    # Get window information
                    rect = window.rectangle()

                    window_info = WindowInfo(
                        handle=handle,
                        title=window.window_text(),
                        process_name=window.process_id(),
                        class_name=window.class_name(),
                        rect=(rect.left, rect.top, rect.right, rect.bottom),
                        is_visible=window.is_visible(),
                        is_minimized=window.is_minimized()
                    )

                    windows.append(window_info)

                except Exception:
                    # Skip windows that can't be accessed
                    continue

        except Exception as e:
            print(f"Error getting window list: {e}")

        return windows

    def filter_windows(self, windows: list[WindowInfo],
                      title_contains: str = None,
                      process_name: str = None,
                      exclude_system: bool = True) -> list[WindowInfo]:
        """Filter windows based on criteria."""
        filtered = []

        # System windows to exclude
        system_classes = {
            "Shell_TrayWnd", "DV2ControlHost", "MsgrIMEWindowClass",
            "SysShadow", "Button", "Windows.UI.Core.CoreWindow"
        }

        for window in windows:
            # Skip empty titles
            if not window.title.strip():
                continue

            # Skip system windows if requested
            if exclude_system and window.class_name in system_classes:
                continue

            # Filter by title
            if title_contains and title_contains.lower() not in window.title.lower():
                continue

            # Filter by process name
            if process_name and process_name.lower() not in str(window.process_name).lower():
                continue

            filtered.append(window)

        return filtered

    def arrange_windows_grid(self, windows: list[WindowInfo], rows: int, cols: int) -> None:
        """Arrange windows in a grid layout."""
        if not windows:
            return

        screen_rect = self.desktop.rectangle()
        screen_width = screen_rect.width()
        screen_height = screen_rect.height() - 40  # Leave space for taskbar

        window_width = screen_width // cols
        window_height = screen_height // rows

        for i, window_info in enumerate(windows[:rows * cols]):
            row = i // cols
            col = i % cols

            x = col * window_width
            y = row * window_height

            try:
                window = self.desktop.window(handle=window_info.handle)
                window.move_window(x, y, window_width, window_height)
                print(f"Moved '{window_info.title}' to grid position ({row}, {col})")
            except Exception as e:
                print(f"Failed to move window '{window_info.title}': {e}")

    def arrange_windows_cascade(self, windows: list[WindowInfo], offset: int = 30) -> None:
        """Arrange windows in cascade layout."""
        if not windows:
            return

        base_width = 800
        base_height = 600

        for i, window_info in enumerate(windows):
            x = i * offset
            y = i * offset

            try:
                window = self.desktop.window(handle=window_info.handle)
                window.move_window(x, y, base_width, base_height)
                print(f"Cascaded '{window_info.title}' at position ({x}, {y})")
            except Exception as e:
                print(f"Failed to cascade window '{window_info.title}': {e}")

    def arrange_windows_side_by_side(self, windows: list[WindowInfo]) -> None:
        """Arrange windows side by side."""
        if not windows:
            return

        screen_rect = self.desktop.rectangle()
        screen_width = screen_rect.width()
        screen_height = screen_rect.height() - 40

        window_width = screen_width // len(windows)

        for i, window_info in enumerate(windows):
            x = i * window_width
            y = 0

            try:
                window = self.desktop.window(handle=window_info.handle)
                window.move_window(x, y, window_width, screen_height)
                print(f"Positioned '{window_info.title}' at x={x}")
            except Exception as e:
                print(f"Failed to position window '{window_info.title}': {e}")

    def create_workspace_layout(self, layout_name: str, window_configs: list[dict]) -> None:
        """Create a custom workspace layout."""
        layout = Layout(name=layout_name, windows=window_configs)
        self.layouts[layout_name] = layout
        print(f"Created layout '{layout_name}' with {len(window_configs)} window configurations")

    def apply_workspace_layout(self, layout_name: str) -> None:
        """Apply a saved workspace layout."""
        if layout_name not in self.layouts:
            raise ValueError(f"Layout '{layout_name}' not found")

        layout = self.layouts[layout_name]
        current_windows = self.get_all_windows()

        for window_config in layout.windows:
            title_pattern = window_config.get("title_pattern", "")
            target_rect = window_config.get("rect", (0, 0, 800, 600))

            # Find matching window
            matching_windows = self.filter_windows(
                current_windows,
                title_contains=title_pattern
            )

            if matching_windows:
                window_info = matching_windows[0]  # Use first match
                try:
                    window = self.desktop.window(handle=window_info.handle)
                    window.move_window(target_rect[0], target_rect[1],
                                     target_rect[2] - target_rect[0],
                                     target_rect[3] - target_rect[1])
                    print(f"Applied layout to '{window_info.title}'")
                except Exception as e:
                    print(f"Failed to apply layout to '{window_info.title}': {e}")
            else:
                print(f"No window found matching pattern '{title_pattern}'")

    def save_current_layout(self, layout_name: str) -> None:
        """Save current window positions as a layout."""
        current_windows = self.get_all_windows()
        window_configs = []

        for window_info in current_windows:
            # Only save windows with meaningful titles
            if len(window_info.title.strip()) > 3:
                config = {
                    "title_pattern": window_info.title[:20],  # First 20 chars
                    "rect": window_info.rect,
                    "process_name": window_info.process_name
                }
                window_configs.append(config)

        self.create_workspace_layout(layout_name, window_configs)
        self.saved_layouts[layout_name] = window_configs
        print(f"Saved current layout as '{layout_name}' with {len(window_configs)} windows")

    def export_layouts(self, filename: str) -> None:
        """Export layouts to JSON file."""
        layout_data = {
            name: {"windows": layout.windows}
            for name, layout in self.layouts.items()
        }

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(layout_data, f, indent=2)

        print(f"Exported {len(layout_data)} layouts to {filename}")

    def import_layouts(self, filename: str) -> None:
        """Import layouts from JSON file."""
        try:
            with open(filename, encoding='utf-8') as f:
                layout_data = json.load(f)

            for layout_name, layout_config in layout_data.items():
                self.create_workspace_layout(layout_name, layout_config["windows"])

            print(f"Imported {len(layout_data)} layouts from {filename}")

        except Exception as e:
            print(f"Failed to import layouts from {filename}: {e}")

    def minimize_all_except(self, keep_titles: list[str]) -> None:
        """Minimize all windows except those with specified titles."""
        current_windows = self.get_all_windows()
        minimized_count = 0

        for window_info in current_windows:
            should_keep = any(title.lower() in window_info.title.lower()
                            for title in keep_titles)

            if not should_keep and not window_info.is_minimized:
                try:
                    window = self.desktop.window(handle=window_info.handle)
                    window.minimize()
                    minimized_count += 1
                except Exception:
                    pass

        print(f"Minimized {minimized_count} windows")

    def focus_window_by_title(self, title_pattern: str) -> bool:
        """Focus on first window matching title pattern."""
        current_windows = self.get_all_windows()
        matching_windows = self.filter_windows(current_windows, title_contains=title_pattern)

        if matching_windows:
            window_info = matching_windows[0]
            try:
                window = self.desktop.window(handle=window_info.handle)
                window.set_focus()
                print(f"Focused on window: '{window_info.title}'")
                return True
            except Exception as e:
                print(f"Failed to focus window: {e}")

        print(f"No window found matching '{title_pattern}'")
        return False

    def get_window_statistics(self) -> dict[str, Any]:
        """Get statistics about current windows."""
        windows = self.get_all_windows()

        stats = {
            "total_windows": len(windows),
            "visible_windows": len([w for w in windows if w.is_visible]),
            "minimized_windows": len([w for w in windows if w.is_minimized]),
            "processes": len(set(w.process_name for w in windows)),
            "window_titles": [w.title for w in windows if w.title.strip()]
        }

        return stats


def demo_window_organization() -> None:
    """Demonstrate various window organization features."""
    manager = AdvancedWindowManager()

    print("=== Advanced Window Management Demo ===")

    # Get current windows
    windows = manager.get_all_windows()
    print(f"Found {len(windows)} windows")

    # Filter meaningful windows
    meaningful_windows = manager.filter_windows(windows, exclude_system=True)
    meaningful_windows = [w for w in meaningful_windows if len(w.title.strip()) > 3]

    if not meaningful_windows:
        print("No suitable windows found for demo. Please open some applications first.")
        return

    print(f"Working with {len(meaningful_windows)} meaningful windows:")
    for w in meaningful_windows:
        print(f"  - {w.title}")

    # Save current layout
    manager.save_current_layout("before_demo")

    try:
        # Demo 1: Arrange in grid
        print("\n--- Demo 1: Grid Layout ---")
        manager.arrange_windows_grid(meaningful_windows[:4], rows=2, cols=2)
        time.sleep(3)

        # Demo 2: Side by side
        print("\n--- Demo 2: Side by Side ---")
        manager.arrange_windows_side_by_side(meaningful_windows[:3])
        time.sleep(3)

        # Demo 3: Cascade
        print("\n--- Demo 3: Cascade Layout ---")
        manager.arrange_windows_cascade(meaningful_windows[:4])
        time.sleep(3)

        # Demo 4: Focus specific window
        print("\n--- Demo 4: Focus Management ---")
        if meaningful_windows:
            target_title = meaningful_windows[0].title[:10]
            manager.focus_window_by_title(target_title)

    except Exception as e:
        print(f"Demo error: {e}")

    # Restore original layout
    print("\n--- Restoring Original Layout ---")
    if "before_demo" in manager.layouts:
        manager.apply_workspace_layout("before_demo")


def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python advanced_window_manager.py <command> [arguments]")
        print("Commands:")
        print("  demo                    - Run organization demo")
        print("  list                    - List all windows")
        print("  grid <rows> <cols>      - Arrange windows in grid")
        print("  cascade                 - Arrange windows in cascade")
        print("  side                    - Arrange windows side by side")
        print("  save <layout_name>      - Save current layout")
        print("  load <layout_name>      - Load saved layout")
        print("  focus <title_pattern>   - Focus window by title")
        print("  minimize_except <title> - Minimize all except specified")
        print("  stats                   - Show window statistics")
        return 1

    command = sys.argv[1]
    manager = AdvancedWindowManager()

    try:
        if command == "demo":
            demo_window_organization()

        elif command == "list":
            windows = manager.get_all_windows()
            filtered = manager.filter_windows(windows, exclude_system=True)
            for i, w in enumerate(filtered):
                print(f"{i+1}. {w.title} (PID: {w.process_name})")

        elif command == "grid" and len(sys.argv) >= 4:
            rows = int(sys.argv[2])
            cols = int(sys.argv[3])
            windows = manager.get_all_windows()
            filtered = manager.filter_windows(windows, exclude_system=True)
            manager.arrange_windows_grid(filtered[:rows*cols], rows, cols)

        elif command == "cascade":
            windows = manager.get_all_windows()
            filtered = manager.filter_windows(windows, exclude_system=True)
            manager.arrange_windows_cascade(filtered)

        elif command == "side":
            windows = manager.get_all_windows()
            filtered = manager.filter_windows(windows, exclude_system=True)
            manager.arrange_windows_side_by_side(filtered)

        elif command == "save" and len(sys.argv) >= 3:
            layout_name = sys.argv[2]
            manager.save_current_layout(layout_name)

        elif command == "load" and len(sys.argv) >= 3:
            layout_name = sys.argv[2]
            manager.apply_workspace_layout(layout_name)

        elif command == "focus" and len(sys.argv) >= 3:
            title_pattern = sys.argv[2]
            manager.focus_window_by_title(title_pattern)

        elif command == "minimize_except" and len(sys.argv) >= 3:
            keep_title = sys.argv[2]
            manager.minimize_all_except([keep_title])

        elif command == "stats":
            stats = manager.get_window_statistics()
            print("Window Statistics:")
            for key, value in stats.items():
                if key == "window_titles":
                    print(f"  {key}: {len(value)} titles")
                else:
                    print(f"  {key}: {value}")

        else:
            print(f"Unknown command: {command}")
            return 1

    except Exception as e:
        print(f"Error executing command: {e}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
