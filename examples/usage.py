#!/usr/bin/env python3
"""
Clean Usage Examples for Enhanced Architecture
Supports Aspose.Cells and Aspose.Slides
"""

import sys
import os
import shutil
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from main import AsposeDocumentProcessor, ConversionOptions, OutputFormat


def basic_conversion_example():
    """Basic conversion examples for Cells & Slides"""
    print("=== Basic Conversion Examples ===")

    # Initialize (this also starts JVM and loads ALL jars)
    processor = AsposeDocumentProcessor()

    print("Available processors:")
    for name, info in processor.list_processors().items():
        print(f"  {name}: {', '.join(info['supported_extensions'])}")

    # -----------------------------
    # Aspose.Cells demo
    # -----------------------------
    if not os.path.exists("sample.csv"):
        create_sample_csv()

    try:
        result = processor.convert_document(
            "sample.csv",
            OutputFormat.MARKDOWN,
            "cells_output.md"
        )
        print(f"✅ Cells basic conversion: {result}")

        options = ConversionOptions(
            include_metadata=True,
            preserve_formatting=True,
            include_hidden_sheets=False
        )
        result = processor.convert_document(
            "sample.csv",
            OutputFormat.JSON,
            "cells_output.json",
            options
        )
        print(f"✅ Cells enhanced conversion: {result}")

        info = processor.get_document_info("sample.csv")
        print(f"✅ Cells document info: {info.get('worksheet_count', '?')} worksheets")

    except Exception as e:
        print(f"❌ Cells error: {e}")

    # -----------------------------
    # Aspose.Slides demo
    # -----------------------------
    if 'slides' in processor.processors:
        try:
            sample_pptx = "sample.pptx"
            if not os.path.exists(sample_pptx):
                create_sample_pptx(sample_pptx)

            result = processor.convert_document(
                sample_pptx,
                OutputFormat.MARKDOWN,
                "slides_output.md"
            )
            print(f"✅ Slides basic conversion: {result}")

            result = processor.convert_document(
                sample_pptx,
                OutputFormat.JSON,
                "slides_output.json",
                ConversionOptions(include_metadata=True)
            )
            print(f"✅ Slides enhanced conversion: {result}")

            result = processor.convert_document(
                sample_pptx,
                OutputFormat.HTML,
                "slides_output.html"
            )
            print(f"✅ Slides HTML conversion: {result}")

            info = processor.get_document_info(sample_pptx)
            print(f"✅ Slides document info: {info.get('slide_count', '?')} slides")

        except Exception as e:
            print(f"❌ Slides error: {e}")
    else:
        print("ℹ️ Slides processor not available (missing plugin or JARs).")


def advanced_usage_example():
    """Advanced usage: batch conversions for Cells & Slides"""
    print("\n=== Advanced Usage Examples ===")
    processor = AsposeDocumentProcessor()

    # -----------------------------
    # Batch processing for Cells
    # -----------------------------
    if 'cells' in processor.processors:
        cells_processor = processor.processors['cells']
        print("Cells batch processing example...")
        try:
            test_dir = Path("test_batch")
            out_dir = Path("cells_batch_output")
            test_dir.mkdir(exist_ok=True)
            out_dir.mkdir(exist_ok=True)

            if os.path.exists("sample.csv"):
                shutil.copy("sample.csv", test_dir / "file1.csv")
                shutil.copy("sample.csv", test_dir / "file2.csv")

                results = cells_processor.batch_convert(
                    str(test_dir),
                    str(out_dir),
                    "markdown",
                    ConversionOptions(include_metadata=True)
                )
                print(f"✅ Cells batch converted {len(results)} files")

            shutil.rmtree(test_dir, ignore_errors=True)
        except Exception as e:
            print(f"❌ Cells batch processing error: {e}")
    else:
        print("ℹ️ Cells processor not available.")

    # -----------------------------
    # Batch processing for Slides
    # -----------------------------
    if 'slides' in processor.processors:
        slides_processor = processor.processors['slides']
        print("Slides batch processing example...")
        try:
            test_dir = Path("test_slides_batch")
            out_dir = Path("slides_batch_output")
            test_dir.mkdir(exist_ok=True)
            out_dir.mkdir(exist_ok=True)

            sample_pptx = "sample.pptx"
            if not os.path.exists(sample_pptx):
                create_sample_pptx(sample_pptx)

            shutil.copy(sample_pptx, test_dir / "deck1.pptx")
            shutil.copy(sample_pptx, test_dir / "deck2.pptx")

            results = slides_processor.batch_convert(
                str(test_dir),
                str(out_dir),
                "markdown",
                ConversionOptions(include_metadata=True)
            )
            print(f"✅ Slides batch converted {len(results)} files")

            shutil.rmtree(test_dir, ignore_errors=True)
        except Exception as e:
            print(f"❌ Slides batch processing error: {e}")
    else:
        print("ℹ️ Slides processor not available.")


def create_sample_csv():
    """Create a sample CSV for testing"""
    csv_content = """Name,Age,Department,Salary,Location
John Smith,30,Engineering,75000,New York
Jane Doe,28,Marketing,65000,Los Angeles
Mike Johnson,35,Sales,70000,Chicago
Sarah Wilson,32,HR,60000,Houston
David Brown,29,Engineering,72000,Seattle
Lisa Garcia,31,Marketing,68000,Miami"""
    with open("sample.csv", "w", encoding="utf-8") as f:
        f.write(csv_content)
    print("✅ Created sample.csv")


def create_sample_pptx(path="sample.pptx"):
    """Create a simple sample PPTX using Aspose.Slides Java API"""
    try:
        # JVM is already started by AsposeDocumentProcessor()
        from com.aspose.slides import Presentation, SaveFormat, ShapeType

        pres = Presentation()
        slide = pres.getSlides().get_Item(0)
        shapes = slide.getShapes()

        # Title shape
        title = shapes.addAutoShape(ShapeType.Rectangle, 80, 80, 560, 60)
        title.addTextFrame("Hello from Aspose.Slides")

        # Subtitle
        subtitle = shapes.addAutoShape(ShapeType.Rectangle, 80, 180, 560, 60)
        subtitle.addTextFrame("Sample presentation for testing")

        pres.save(path, SaveFormat.Pptx)
        print(f"✅ Created {path}")
    except Exception as e:
        print(f"❌ Could not create sample PPTX: {e}")


def cli_examples():
    """Show CLI usage examples"""
    print("\n=== CLI Usage Examples ===")
    print("# Cells: CSV to Markdown")
    print("python main.py convert sample.csv markdown output.md\n")
    print("# Slides: PPTX to JSON (with metadata)")
    print("python main.py convert sample.pptx json slides_output.json --include-metadata\n")
    print("# Get info")
    print("python main.py info document.xlsx")
    print("python main.py info document.pptx\n")
    print("# List capabilities")
    print("python main.py list-formats")
    print("python main.py list-processors")


if __name__ == "__main__":
    print("Clean Aspose Document Processor Examples")
    print("========================================")
    basic_conversion_example()
    advanced_usage_example()
    cli_examples()
    print("\n🎉 Examples completed!")
