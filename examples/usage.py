#!/usr/bin/env python3
"""
Clean Usage Examples for Enhanced Architecture
"""

import sys
import os
from pathlib import Path

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from main import AsposeDocumentProcessor, ConversionOptions, OutputFormat

def basic_conversion_example():
    """Basic conversion examples"""
    print("=== Basic Conversion Examples ===")
    
    # Initialize processor
    processor = AsposeDocumentProcessor()
    
    # Check what's available
    print("Available processors:")
    for name, info in processor.list_processors().items():
        print(f"  {name}: {', '.join(info['supported_extensions'])}")
    
    # Create sample data if needed
    if not os.path.exists("sample.csv"):
        create_sample_csv()
    
    # Basic conversions
    try:
        # Method 1: Simple conversion
        result = processor.convert_document(
            "sample.csv", 
            OutputFormat.MARKDOWN, 
            "basic_output.md"
        )
        print(f"✅ Basic conversion: {result}")
        
        # Method 2: With custom options
        options = ConversionOptions(
            include_metadata=True,
            preserve_formatting=True,
            include_hidden_sheets=False
        )
        
        result = processor.convert_document(
            "sample.csv",
            OutputFormat.JSON,
            "enhanced_output.json", 
            options
        )
        print(f"✅ Enhanced conversion: {result}")
        
        # Get document info
        info = processor.get_document_info("sample.csv")
        print(f"✅ Document info: {info['worksheet_count']} worksheets")
        
    except Exception as e:
        print(f"❌ Error: {e}")

def advanced_usage_example():
    """Advanced usage examples"""
    print("\n=== Advanced Usage Examples ===")
    
    processor = AsposeDocumentProcessor()
    
    # Get specific processor for advanced features
    if 'cells' in processor.processors:
        cells_processor = processor.processors['cells']
        
        # Batch processing
        print("Batch processing example...")
        try:
            # Create test directory with files
            test_dir = Path("test_batch")
            test_dir.mkdir(exist_ok=True)
            
            # Copy sample file for batch test
            if os.path.exists("sample.csv"):
                import shutil
                shutil.copy("sample.csv", test_dir / "file1.csv")
                shutil.copy("sample.csv", test_dir / "file2.csv")
                
                # Batch convert
                options = ConversionOptions(include_metadata=True)
                results = cells_processor.batch_convert(
                    str(test_dir), 
                    "batch_output", 
                    "markdown", 
                    options
                )
                print(f"✅ Batch converted {len(results)} files")
                
                # Cleanup
                shutil.rmtree(test_dir)
                
        except Exception as e:
            print(f"❌ Batch processing error: {e}")

def create_sample_csv():
    """Create a sample CSV for testing"""
    csv_content = """Name,Age,Department,Salary,Location
John Smith,30,Engineering,75000,New York
Jane Doe,28,Marketing,65000,Los Angeles  
Mike Johnson,35,Sales,70000,Chicago
Sarah Wilson,32,HR,60000,Houston
David Brown,29,Engineering,72000,Seattle
Lisa Garcia,31,Marketing,68000,Miami"""

    with open("sample.csv", "w") as f:
        f.write(csv_content)
    print("✅ Created sample.csv")

def cli_examples():
    """Show CLI usage examples"""
    print("\n=== CLI Usage Examples ===")
    print("# Basic conversion:")
    print("python main.py convert sample.csv markdown output.md")
    print()
    print("# With options:")
    print("python main.py convert sample.xlsx json --include-metadata --no-formatting")
    print()
    print("# Get info:")
    print("python main.py info document.xlsx")
    print()
    print("# List capabilities:")
    print("python main.py list-formats")
    print("python main.py list-processors")

if __name__ == "__main__":
    print("Clean Aspose Document Processor Examples")
    print("========================================")
    
    basic_conversion_example()
    advanced_usage_example() 
    cli_examples()
    
    print("\n🎉 Examples completed!")