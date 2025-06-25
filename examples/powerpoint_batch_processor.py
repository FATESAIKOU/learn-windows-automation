#!/usr/bin/env python3
"""
PowerPoint Batch Processor - Advanced pywin32 example
Batch processes PowerPoint files: updates content, applies templates, exports formats
"""

import json
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


class PowerPointBatchProcessor:
    """Advanced PowerPoint batch processing using pywin32."""

    def __init__(self):
        """Initialize PowerPoint processor."""
        if not WIN32_AVAILABLE:
            raise RuntimeError("pywin32 not available for PowerPoint automation")

        self.ppt_app = None
        self.presentation = None

    def start_powerpoint(self, visible: bool = True) -> None:
        """Start PowerPoint application."""
        try:
            self.ppt_app = win32.Dispatch("PowerPoint.Application")
            self.ppt_app.Visible = 1 if visible else 0
            print("PowerPoint application started successfully")
        except Exception as e:
            raise RuntimeError(f"Failed to start PowerPoint: {e}")

    def create_presentation(self) -> None:
        """Create a new presentation."""
        if not self.ppt_app:
            raise RuntimeError("PowerPoint application not started")

        self.presentation = self.ppt_app.Presentations.Add(WithWindow=1)
        print("New presentation created")

    def open_presentation(self, filename: str) -> None:
        """Open existing presentation."""
        if not self.ppt_app:
            raise RuntimeError("PowerPoint application not started")

        try:
            self.presentation = self.ppt_app.Presentations.Open(filename)
            print(f"Presentation opened: {filename}")
        except Exception as e:
            raise RuntimeError(f"Failed to open presentation {filename}: {e}")

    def add_slide(self, layout_type: int = 1) -> Any:
        """Add a new slide to the presentation."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        # Layout types: 1=Title, 2=Title and Content, 3=Section Header, etc.
        slide = self.presentation.Slides.Add(
            Index=self.presentation.Slides.Count + 1,
            Layout=layout_type
        )
        print(f"Slide added with layout type {layout_type}")
        return slide

    def update_slide_content(self, slide_index: int, title: str = None, content: list[str] = None) -> None:
        """Update content of a specific slide."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        if slide_index > self.presentation.Slides.Count:
            raise ValueError(f"Slide index {slide_index} out of range")

        slide = self.presentation.Slides(slide_index)

        # Update title if provided
        if title:
            for shape in slide.Shapes:
                if shape.HasTextFrame and "Title" in shape.Name:
                    shape.TextFrame.TextRange.Text = title
                    print(f"Updated title on slide {slide_index}")
                    break

        # Update content if provided
        if content:
            for shape in slide.Shapes:
                if shape.HasTextFrame and "Content" in shape.Name:
                    # Clear existing content
                    shape.TextFrame.TextRange.Text = ""

                    # Add bullet points
                    for i, item in enumerate(content):
                        if i == 0:
                            shape.TextFrame.TextRange.Text = item
                        else:
                            # Add new paragraph
                            para = shape.TextFrame.TextRange.Paragraphs().Add()
                            para.Text = item

                    print(f"Updated content on slide {slide_index} with {len(content)} items")
                    break

    def apply_design_template(self, template_path: str) -> None:
        """Apply a design template to the presentation."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        try:
            self.presentation.ApplyTemplate(template_path)
            print(f"Applied design template: {template_path}")
        except Exception as e:
            print(f"Warning: Could not apply template {template_path}: {e}")

    def batch_replace_text(self, old_text: str, new_text: str) -> int:
        """Replace all occurrences of text across all slides."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        replacements = 0

        for slide in self.presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame:
                    text_range = shape.TextFrame.TextRange
                    if old_text in text_range.Text:
                        text_range.Text = text_range.Text.replace(old_text, new_text)
                        replacements += 1

        print(f"Replaced '{old_text}' with '{new_text}' in {replacements} locations")
        return replacements

    def add_footer_to_all_slides(self, footer_text: str) -> None:
        """Add footer text to all slides."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        for slide in self.presentation.Slides:
            # Add text box at bottom of slide
            left = 50
            top = self.presentation.PageSetup.SlideHeight - 50
            width = self.presentation.PageSetup.SlideWidth - 100
            height = 30

            text_box = slide.Shapes.AddTextbox(
                Orientation=1,  # msoTextOrientationHorizontal
                Left=left,
                Top=top,
                Width=width,
                Height=height
            )

            text_box.TextFrame.TextRange.Text = footer_text
            text_box.TextFrame.TextRange.Font.Size = 10
            text_box.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # Center

        print(f"Added footer '{footer_text}' to all slides")

    def export_as_pdf(self, output_file: str) -> None:
        """Export presentation as PDF."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        try:
            self.presentation.SaveAs(output_file, FileFormat=32)  # ppSaveAsPDF
            print(f"Presentation exported as PDF: {output_file}")
        except Exception as e:
            raise RuntimeError(f"Failed to export as PDF: {e}")

    def save_presentation(self, filename: str = None) -> None:
        """Save the presentation."""
        if not self.presentation:
            raise RuntimeError("No active presentation")

        if filename:
            self.presentation.SaveAs(filename)
            print(f"Presentation saved as: {filename}")
        else:
            self.presentation.Save()
            print("Presentation saved")

    def close_presentation(self) -> None:
        """Close the presentation."""
        if self.presentation:
            self.presentation.Close()
            self.presentation = None
            print("Presentation closed")

    def quit_powerpoint(self) -> None:
        """Quit PowerPoint application."""
        if self.ppt_app:
            self.ppt_app.Quit()
            self.ppt_app = None
            print("PowerPoint application closed")


def create_sample_presentation(output_file: str) -> None:
    """Create a sample business presentation."""
    processor = PowerPointBatchProcessor()

    try:
        # Start PowerPoint
        processor.start_powerpoint(visible=True)
        processor.create_presentation()

        # Slide 1: Title slide
        slide1 = processor.add_slide(1)  # Title layout
        processor.update_slide_content(1,
            title="Q1 Business Review",
            content=["Prepared by Automation Team", "Date: 2025-Q1"]
        )

        # Slide 2: Agenda
        slide2 = processor.add_slide(2)  # Title and Content
        processor.update_slide_content(2,
            title="Agenda",
            content=[
                "Executive Summary",
                "Financial Performance",
                "Key Achievements",
                "Challenges & Opportunities",
                "Q2 Outlook"
            ]
        )

        # Slide 3: Financial Performance
        slide3 = processor.add_slide(2)
        processor.update_slide_content(3,
            title="Financial Performance",
            content=[
                "Revenue: $2.5M (+15% YoY)",
                "Profit Margin: 23% (+2% vs Q4)",
                "Operating Expenses: $1.9M (-5% vs Q4)",
                "Cash Flow: Positive $600K"
            ]
        )

        # Slide 4: Key Achievements
        slide4 = processor.add_slide(2)
        processor.update_slide_content(4,
            title="Key Achievements",
            content=[
                "Launched new product line",
                "Expanded to 3 new markets",
                "Improved customer satisfaction by 12%",
                "Reduced operational costs by 8%"
            ]
        )

        # Add footer to all slides
        processor.add_footer_to_all_slides("Confidential - Internal Use Only")

        # Save presentation
        processor.save_presentation(output_file)

        print(f"Sample presentation created successfully: {output_file}")

    except Exception as e:
        print(f"Error creating presentation: {e}")
    finally:
        # Clean up
        processor.close_presentation()
        processor.quit_powerpoint()


def batch_update_presentations(config_file: str) -> None:
    """Batch update multiple presentations based on config."""
    try:
        with open(config_file, encoding='utf-8') as f:
            config = json.load(f)
    except Exception as e:
        print(f"Error reading config file {config_file}: {e}")
        return

    processor = PowerPointBatchProcessor()

    try:
        processor.start_powerpoint(visible=False)  # Run in background for batch processing

        for file_config in config.get("files", []):
            input_file = file_config.get("input")
            output_file = file_config.get("output", input_file)

            print(f"Processing: {input_file}")

            # Open presentation
            processor.open_presentation(input_file)

            # Apply replacements
            replacements = file_config.get("replacements", {})
            for old_text, new_text in replacements.items():
                processor.batch_replace_text(old_text, new_text)

            # Add footer if specified
            if "footer" in file_config:
                processor.add_footer_to_all_slides(file_config["footer"])

            # Save file
            processor.save_presentation(output_file)

            # Export as PDF if requested
            if file_config.get("export_pdf"):
                pdf_file = output_file.replace(".pptx", ".pdf")
                processor.export_as_pdf(pdf_file)

            processor.close_presentation()
            print(f"Completed: {output_file}")

    except Exception as e:
        print(f"Error in batch processing: {e}")
    finally:
        processor.quit_powerpoint()


def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2:
        print("Usage: python powerpoint_batch_processor.py <command> [arguments]")
        print("Commands:")
        print("  create <output_file>     - Create sample presentation")
        print("  batch <config_file>      - Batch process presentations")
        print("Example: python powerpoint_batch_processor.py create sample.pptx")
        return 1

    command = sys.argv[1]

    if command == "create" and len(sys.argv) >= 3:
        output_file = sys.argv[2]
        create_sample_presentation(output_file)
    elif command == "batch" and len(sys.argv) >= 3:
        config_file = sys.argv[2]
        batch_update_presentations(config_file)
    else:
        print(f"Invalid command or missing arguments: {command}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
