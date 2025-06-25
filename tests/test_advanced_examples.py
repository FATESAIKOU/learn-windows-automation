"""Tests for advanced example scripts."""

import sys
from pathlib import Path
from unittest.mock import MagicMock, call, patch

import pytest

# Add examples to path for testing
examples_path = Path(__file__).parent.parent / "examples"
sys.path.insert(0, str(examples_path))


class TestExcelDataProcessor:
    """Test cases for Excel data processor."""

    @patch('excel_data_processor.WIN32_AVAILABLE', True)
    @patch('excel_data_processor.win32')
    def test_excel_processor_initialization(self, mock_win32):
        """Test Excel processor initialization."""
        from excel_data_processor import ExcelDataProcessor

        processor = ExcelDataProcessor()
        assert processor.excel_app is None
        assert processor.workbook is None
        assert processor.worksheet is None

    @patch('excel_data_processor.WIN32_AVAILABLE', False)
    def test_excel_processor_no_win32(self):
        """Test Excel processor when win32 not available."""
        from excel_data_processor import ExcelDataProcessor

        with pytest.raises(RuntimeError, match="pywin32 not available"):
            ExcelDataProcessor()

    @patch('excel_data_processor.WIN32_AVAILABLE', True)
    @patch('excel_data_processor.win32')
    def test_start_excel(self, mock_win32):
        """Test starting Excel application."""
        from excel_data_processor import ExcelDataProcessor

        # Mock Excel application
        mock_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_app

        processor = ExcelDataProcessor()
        processor.start_excel(visible=True)

        mock_win32.Dispatch.assert_called_once_with("Excel.Application")
        assert mock_app.Visible is True
        assert mock_app.DisplayAlerts is False
        assert processor.excel_app == mock_app

    @patch('excel_data_processor.WIN32_AVAILABLE', True)
    @patch('excel_data_processor.win32')
    def test_write_data(self, mock_win32):
        """Test writing data to worksheet."""
        from excel_data_processor import ExcelDataProcessor

        # Mock worksheet
        mock_worksheet = MagicMock()
        mock_cell = MagicMock()
        mock_worksheet.Cells.return_value = mock_cell

        processor = ExcelDataProcessor()
        processor.worksheet = mock_worksheet

        test_data = [["A1", "B1"], ["A2", "B2"]]
        processor.write_data(test_data)

        # Verify cells were set
        expected_calls = [
            call(1, 1),  # A1
            call(1, 2),  # B1
            call(2, 1),  # A2
            call(2, 2),  # B2
        ]
        mock_worksheet.Cells.assert_has_calls(expected_calls)

    @patch('excel_data_processor.WIN32_AVAILABLE', True)
    @patch('excel_data_processor.win32')
    def test_create_chart(self, mock_win32):
        """Test creating chart in worksheet."""
        from excel_data_processor import ExcelDataProcessor

        # Mock worksheet and chart
        mock_worksheet = MagicMock()
        mock_chart = MagicMock()
        mock_shapes = MagicMock()
        mock_shapes.AddChart2.return_value.Chart = mock_chart
        mock_worksheet.Shapes = mock_shapes
        mock_worksheet.Range.return_value = "mock_range"

        processor = ExcelDataProcessor()
        processor.worksheet = mock_worksheet

        processor.create_chart("A1:B5", "Column", "Test Chart")

        # Verify chart creation
        mock_shapes.AddChart2.assert_called_once()
        mock_chart.SetSourceData.assert_called_once_with("mock_range")
        assert mock_chart.ChartTitle.Text == "Test Chart"


class TestPowerPointBatchProcessor:
    """Test cases for PowerPoint batch processor."""

    @patch('powerpoint_batch_processor.WIN32_AVAILABLE', True)
    @patch('powerpoint_batch_processor.win32')
    def test_powerpoint_initialization(self, mock_win32):
        """Test PowerPoint processor initialization."""
        from powerpoint_batch_processor import PowerPointBatchProcessor

        processor = PowerPointBatchProcessor()
        assert processor.ppt_app is None
        assert processor.presentation is None

    @patch('powerpoint_batch_processor.WIN32_AVAILABLE', False)
    def test_powerpoint_no_win32(self):
        """Test PowerPoint processor when win32 not available."""
        from powerpoint_batch_processor import PowerPointBatchProcessor

        with pytest.raises(RuntimeError, match="pywin32 not available"):
            PowerPointBatchProcessor()

    @patch('powerpoint_batch_processor.WIN32_AVAILABLE', True)
    @patch('powerpoint_batch_processor.win32')
    def test_start_powerpoint(self, mock_win32):
        """Test starting PowerPoint application."""
        from powerpoint_batch_processor import PowerPointBatchProcessor

        mock_app = MagicMock()
        mock_win32.Dispatch.return_value = mock_app

        processor = PowerPointBatchProcessor()
        processor.start_powerpoint(visible=True)

        mock_win32.Dispatch.assert_called_once_with("PowerPoint.Application")
        assert mock_app.Visible == 1
        assert processor.ppt_app == mock_app

    @patch('powerpoint_batch_processor.WIN32_AVAILABLE', True)
    @patch('powerpoint_batch_processor.win32')
    def test_add_slide(self, mock_win32):
        """Test adding slide to presentation."""
        from powerpoint_batch_processor import PowerPointBatchProcessor

        mock_presentation = MagicMock()
        mock_slides = MagicMock()
        mock_slides.Count = 2
        mock_slide = MagicMock()
        mock_slides.Add.return_value = mock_slide
        mock_presentation.Slides = mock_slides

        processor = PowerPointBatchProcessor()
        processor.presentation = mock_presentation

        result = processor.add_slide(layout_type=2)

        mock_slides.Add.assert_called_once_with(Index=3, Layout=2)
        assert result == mock_slide

    @patch('powerpoint_batch_processor.WIN32_AVAILABLE', True)
    @patch('powerpoint_batch_processor.win32')
    def test_batch_replace_text(self, mock_win32):
        """Test batch text replacement."""
        from powerpoint_batch_processor import PowerPointBatchProcessor

        # Mock presentation structure
        mock_text_range = MagicMock()
        mock_text_range.Text = "Hello World"

        mock_text_frame = MagicMock()
        mock_text_frame.TextRange = mock_text_range

        mock_shape = MagicMock()
        mock_shape.HasTextFrame = True
        mock_shape.TextFrame = mock_text_frame

        mock_slide = MagicMock()
        mock_slide.Shapes = [mock_shape]

        mock_presentation = MagicMock()
        mock_presentation.Slides = [mock_slide]

        processor = PowerPointBatchProcessor()
        processor.presentation = mock_presentation

        result = processor.batch_replace_text("World", "Universe")

        assert mock_text_range.Text == "Hello Universe"
        assert result == 1


class TestMultiAppCoordinator:
    """Test cases for multi-app coordinator."""

    @patch('multi_app_coordinator.PYWINAUTO_AVAILABLE', True)
    @patch('multi_app_coordinator.Desktop')
    def test_coordinator_initialization(self, mock_desktop):
        """Test coordinator initialization."""
        from multi_app_coordinator import MultiAppCoordinator

        coordinator = MultiAppCoordinator()
        assert coordinator.apps == {}
        assert coordinator.windows == {}
        mock_desktop.assert_called_once_with(backend="uia")

    @patch('multi_app_coordinator.PYWINAUTO_AVAILABLE', False)
    def test_coordinator_no_pywinauto(self):
        """Test coordinator when pywinauto not available."""
        from multi_app_coordinator import MultiAppCoordinator

        with pytest.raises(RuntimeError, match="pywinauto not available"):
            MultiAppCoordinator()

    @patch('multi_app_coordinator.PYWINAUTO_AVAILABLE', True)
    @patch('multi_app_coordinator.Desktop')
    @patch('multi_app_coordinator.Application')
    @patch('multi_app_coordinator.time.sleep')
    def test_launch_application(self, mock_sleep, mock_app_class, mock_desktop):
        """Test launching application."""
        from multi_app_coordinator import MultiAppCoordinator

        # Mock application and window
        mock_app = MagicMock()
        mock_window = MagicMock()
        mock_app.top_window.return_value = mock_window
        mock_app_class.return_value.start.return_value = mock_app

        coordinator = MultiAppCoordinator()
        coordinator.launch_application("notepad", "notepad.exe")

        mock_app_class.assert_called_with(backend="uia")
        mock_app_class.return_value.start.assert_called_with("notepad.exe")
        mock_window.wait.assert_called_with("visible", timeout=10)

        assert coordinator.apps["notepad"] == mock_app
        assert coordinator.windows["notepad"] == mock_window

    @patch('multi_app_coordinator.PYWINAUTO_AVAILABLE', True)
    @patch('multi_app_coordinator.Desktop')
    @patch('multi_app_coordinator.send_keys')
    def test_send_text_to_app(self, mock_send_keys, mock_desktop):
        """Test sending text to application."""
        from multi_app_coordinator import MultiAppCoordinator

        mock_window = MagicMock()

        coordinator = MultiAppCoordinator()
        coordinator.windows["notepad"] = mock_window

        coordinator.send_text_to_app("notepad", "Hello World")

        mock_window.set_focus.assert_called_once()
        mock_send_keys.assert_called_once_with("Hello World")


class TestAdvancedWindowManager:
    """Test cases for advanced window manager."""

    @patch('advanced_window_manager.PYWINAUTO_AVAILABLE', True)
    @patch('advanced_window_manager.Desktop')
    def test_window_manager_initialization(self, mock_desktop):
        """Test window manager initialization."""
        from advanced_window_manager import AdvancedWindowManager

        manager = AdvancedWindowManager()
        assert manager.layouts == {}
        assert manager.saved_layouts == {}
        mock_desktop.assert_called_once_with(backend="uia")

    @patch('advanced_window_manager.PYWINAUTO_AVAILABLE', False)
    def test_window_manager_no_pywinauto(self):
        """Test window manager when pywinauto not available."""
        from advanced_window_manager import AdvancedWindowManager

        with pytest.raises(RuntimeError, match="pywinauto not available"):
            AdvancedWindowManager()

    @patch('advanced_window_manager.PYWINAUTO_AVAILABLE', True)
    @patch('advanced_window_manager.Desktop')
    def test_create_workspace_layout(self, mock_desktop):
        """Test creating workspace layout."""
        from advanced_window_manager import AdvancedWindowManager, Layout

        manager = AdvancedWindowManager()

        window_configs = [
            {"title_pattern": "notepad", "rect": (0, 0, 800, 600)},
            {"title_pattern": "calculator", "rect": (800, 0, 800, 600)}
        ]

        manager.create_workspace_layout("test_layout", window_configs)

        assert "test_layout" in manager.layouts
        layout = manager.layouts["test_layout"]
        assert isinstance(layout, Layout)
        assert layout.name == "test_layout"
        assert len(layout.windows) == 2

    @patch('advanced_window_manager.PYWINAUTO_AVAILABLE', True)
    @patch('advanced_window_manager.Desktop')
    def test_arrange_windows_grid(self, mock_desktop):
        """Test arranging windows in grid."""
        from advanced_window_manager import AdvancedWindowManager, WindowInfo

        # Mock desktop rectangle
        mock_rect = MagicMock()
        mock_rect.width.return_value = 1920
        mock_rect.height.return_value = 1080
        mock_desktop.return_value.rectangle.return_value = mock_rect

        # Mock windows
        mock_window1 = MagicMock()
        mock_window2 = MagicMock()
        mock_desktop.return_value.window.side_effect = [mock_window1, mock_window2]

        manager = AdvancedWindowManager()

        # Create test window info objects
        windows = [
            WindowInfo(1, "Window 1", "proc1", "class1", (0, 0, 100, 100), True, False),
            WindowInfo(2, "Window 2", "proc2", "class2", (0, 0, 100, 100), True, False)
        ]

        manager.arrange_windows_grid(windows, rows=1, cols=2)

        # Verify windows were moved
        mock_window1.move_window.assert_called_once_with(0, 0, 960, 1040)
        mock_window2.move_window.assert_called_once_with(960, 0, 960, 1040)

    @patch('advanced_window_manager.PYWINAUTO_AVAILABLE', True)
    @patch('advanced_window_manager.Desktop')
    def test_filter_windows(self, mock_desktop):
        """Test filtering windows by criteria."""
        from advanced_window_manager import AdvancedWindowManager, WindowInfo

        manager = AdvancedWindowManager()

        windows = [
            WindowInfo(1, "Notepad", "notepad.exe", "Notepad", (0, 0, 100, 100), True, False),
            WindowInfo(2, "Calculator", "calc.exe", "Calculator", (0, 0, 100, 100), True, False),
            WindowInfo(3, "", "system.exe", "Shell_TrayWnd", (0, 0, 100, 100), True, False),
            WindowInfo(4, "Visual Studio Code", "code.exe", "Chrome", (0, 0, 100, 100), True, False),
        ]

        # Test filtering by title
        filtered = manager.filter_windows(windows, title_contains="note")
        assert len(filtered) == 1
        assert filtered[0].title == "Notepad"

        # Test excluding system windows
        filtered = manager.filter_windows(windows, exclude_system=True)
        assert len(filtered) == 3  # Excludes Shell_TrayWnd and empty title

        # Test filtering by process
        filtered = manager.filter_windows(windows, process_name="calc")
        assert len(filtered) == 1
        assert filtered[0].title == "Calculator"
