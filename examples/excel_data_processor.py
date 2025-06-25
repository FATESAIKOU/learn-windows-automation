#!/usr/bin/env python3
"""
Excel Data Processor - Advanced pywin32 example
Processes Excel files: reads data, creates charts, formats cells, generates reports
"""

import sys
from pathlib import Path
from typing import Any

# Add the src directory to the path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


class ExcelDataProcessor:
    """Advanced Excel data processing using pywin32."""

    def __init__(self):
        """Initialize Excel processor."""
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 not available for Excel automation")

        self.excel_app = None
        self.workbook = None
        self.worksheet = None

    def start_excel(self, visible: bool = True) -> None:
        """Start Excel application."""
        try:
            self.excel_app = win32.Dispatch("Excel.Application")
            self.excel_app.Visible = visible
            self.excel_app.DisplayAlerts = False
            print("Excel application started successfully")
        except Exception as e:
            raise RuntimeError(f"Failed to start Excel: {e}")

    def create_workbook(self, filename: str = None) -> None:
        """Create a new workbook."""
        if not self.excel_app:
            raise RuntimeError("Excel application not started")

        self.workbook = self.excel_app.Workbooks.Add()
        self.worksheet = self.workbook.ActiveSheet

        if filename:
            self.workbook.SaveAs(filename)
            print(f"Workbook created and saved as: {filename}")
        else:
            print("New workbook created")

    def open_workbook(self, filename: str) -> None:
        """Open existing workbook."""
        if not self.excel_app:
            raise RuntimeError("Excel application not started")

        try:
            self.workbook = self.excel_app.Workbooks.Open(filename)
            self.worksheet = self.workbook.ActiveSheet
            print(f"Workbook opened: {filename}")
        except Exception as e:
            raise RuntimeError(f"Failed to open workbook {filename}: {e}")

    def write_data(self, data: list[list[Any]], start_row: int = 1, start_col: int = 1) -> None:
        """Write data to worksheet."""
        if not self.worksheet:
            raise RuntimeError("No active worksheet")

        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                self.worksheet.Cells(start_row + row_idx, start_col + col_idx).Value = cell_value

        print(f"Data written to worksheet: {len(data)} rows, {len(data[0]) if data else 0} columns")

    def format_range(self, start_cell: str, end_cell: str, **formatting) -> None:
        """Format a range of cells."""
        if not self.worksheet:
            raise RuntimeError("No active worksheet")

        cell_range = self.worksheet.Range(f"{start_cell}:{end_cell}")

        # Apply formatting options
        if "bold" in formatting:
            cell_range.Font.Bold = formatting["bold"]
        if "font_size" in formatting:
            cell_range.Font.Size = formatting["font_size"]
        if "background_color" in formatting:
            cell_range.Interior.Color = formatting["background_color"]
        if "border" in formatting and formatting["border"]:
            cell_range.Borders.LineStyle = 1  # Continuous line

        print(f"Formatted range {start_cell}:{end_cell}")

    def create_chart(self, data_range: str, chart_type: str = "Column", title: str = "Chart") -> None:
        """Create a chart from data range."""
        if not self.worksheet:
            raise RuntimeError("No active worksheet")

        # Chart type mapping
        chart_types = {
            "Column": 51,  # xlColumnClustered
            "Line": 4,     # xlLine
            "Pie": 5,      # xlPie
            "Bar": 57,     # xlBarClustered
        }

        chart_type_value = chart_types.get(chart_type, 51)

        # Create chart
        chart = self.worksheet.Shapes.AddChart2(
            Style=-1,
            XlChartType=chart_type_value,
            Left=300,
            Top=50,
            Width=400,
            Height=300
        ).Chart

        # Set data source
        chart.SetSourceData(self.worksheet.Range(data_range))
        chart.ChartTitle.Text = title

        print(f"Chart created: {chart_type} chart with title '{title}'")

    def calculate_formulas(self) -> None:
        """Calculate all formulas in the workbook."""
        if not self.workbook:
            raise RuntimeError("No active workbook")

        self.workbook.Calculate()
        print("All formulas calculated")

    def save_workbook(self, filename: str = None) -> None:
        """Save the workbook."""
        if not self.workbook:
            raise RuntimeError("No active workbook")

        if filename:
            self.workbook.SaveAs(filename)
            print(f"Workbook saved as: {filename}")
        else:
            self.workbook.Save()
            print("Workbook saved")

    def close_workbook(self) -> None:
        """Close the workbook."""
        if self.workbook:
            self.workbook.Close(SaveChanges=False)
            self.workbook = None
            self.worksheet = None
            print("Workbook closed")

    def quit_excel(self) -> None:
        """Quit Excel application."""
        if self.excel_app:
            self.excel_app.Quit()
            self.excel_app = None
            print("Excel application closed")


def generate_sample_report(output_file: str) -> None:
    """Generate a sample sales report."""
    processor = ExcelDataProcessor()

    try:
        # Start Excel
        processor.start_excel(visible=True)
        processor.create_workbook()

        # Sample sales data
        headers = ["Month", "Sales", "Expenses", "Profit"]
        data = [
            headers,
            ["Jan", 10000, 7000, 3000],
            ["Feb", 12000, 8000, 4000],
            ["Mar", 15000, 9000, 6000],
            ["Apr", 11000, 7500, 3500],
            ["May", 13000, 8500, 4500],
            ["Jun", 16000, 10000, 6000],
        ]

        # Write data
        processor.write_data(data)

        # Format headers
        processor.format_range("A1", "D1", bold=True, font_size=12, background_color=0xCCCCCC, border=True)

        # Format data
        processor.format_range("A2", "D7", border=True)

        # Add formulas for totals
        processor.worksheet.Cells(8, 1).Value = "Total"
        processor.worksheet.Cells(8, 2).Formula = "=SUM(B2:B7)"
        processor.worksheet.Cells(8, 3).Formula = "=SUM(C2:C7)"
        processor.worksheet.Cells(8, 4).Formula = "=SUM(D2:D7)"

        # Format totals
        processor.format_range("A8", "D8", bold=True, background_color=0xFFFF99)

        # Create chart
        processor.create_chart("A1:D7", "Column", "Monthly Sales Report")

        # Calculate formulas
        processor.calculate_formulas()

        # Save file
        processor.save_workbook(output_file)

        print(f"Sales report generated successfully: {output_file}")

    except Exception as e:
        print(f"Error generating report: {e}")
    finally:
        # Clean up
        processor.close_workbook()
        processor.quit_excel()


def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python excel_data_processor.py <output_file> [operation]")
        print("Operations: report (default)")
        print("Example: python excel_data_processor.py sales_report.xlsx report")
        return 1

    output_file = sys.argv[1]
    operation = sys.argv[2] if len(sys.argv) > 2 else "report"

    if operation == "report":
        generate_sample_report(output_file)
    else:
        print(f"Unknown operation: {operation}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
