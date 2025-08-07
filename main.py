#!/usr/bin/env python3
"""
Clean Aspose.Total-for-Python-via-Java Document Processor
Focused on Aspose.Cells with clean extensible architecture
"""

import os
import sys
import json
import abc
from pathlib import Path
from typing import Optional, Dict, Any, List
from dataclasses import dataclass
from enum import Enum
import logging

# Add the plugins directory to Python path
PLUGINS_DIR = Path(__file__).parent / "plugins"
sys.path.insert(0, str(PLUGINS_DIR))

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class OutputFormat(Enum):
    """Supported output formats"""
    MARKDOWN = "markdown"
    JSON = "json"
    HTML = "html"

@dataclass
class ConversionOptions:
    """Options for document conversion"""
    include_metadata: bool = True
    preserve_formatting: bool = True
    export_images_as_base64: bool = True
    include_hidden_sheets: bool = False
    custom_options: Optional[Dict[str, Any]] = None
    
    def __post_init__(self):
        if self.custom_options is None:
            self.custom_options = {}

class AsposeJVMManager:
    """Centralized JVM management for Aspose processors"""
    
    _instance = None
    _jvm_started = False
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if not hasattr(self, 'initialized'):
            self.plugins_dir = Path(__file__).parent / "plugins"
            self.initialized = True
    
    def start_jvm_with_jars(self, jar_patterns: List[str]) -> bool:
        """Start JVM with specified JAR patterns"""
        if self._jvm_started:
            logger.info("JVM already started")
            return True
        
        try:
            import jpype
            import jpype.imports
        except ImportError:
            logger.error("JPype1 not found. Install with: pip install JPype1")
            raise
        
        # Find JAR files
        all_jars = []
        for pattern in jar_patterns:
            jars = list(self.plugins_dir.glob(pattern))
            if not jars:
                logger.warning(f"No JAR files found for pattern: {pattern}")
            all_jars.extend(jars)
        
        if not all_jars:
            raise FileNotFoundError(f"No JAR files found in {self.plugins_dir}")
        
        classpath = [str(jar) for jar in all_jars]
        
        try:
            # Mac M4 optimized JVM settings
            jpype.startJVM(
                jpype.getDefaultJVMPath(),
                "-Djava.class.path=" + os.pathsep.join(classpath),
                "-Xmx4g",  # 4GB heap for better performance
                "-Dfile.encoding=UTF-8",
                "-XX:+UseG1GC",  # Better garbage collector for Mac M4
                "-XX:MaxGCPauseMillis=200",
                convertStrings=False
            )
            
            self._jvm_started = True
            logger.info(f"JVM started successfully with JARs: {[jar.name for jar in all_jars]}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to start JVM: {e}")
            raise
    
    def is_started(self) -> bool:
        """Check if JVM is started"""
        return self._jvm_started
    
    def apply_license(self, license_path: str, license_class_name: str) -> bool:
        """Apply Aspose license"""
        if not os.path.exists(license_path):
            logger.warning(f"License file not found: {license_path}")
            return False
        
        try:
            module_parts = license_class_name.split('.')
            class_name = module_parts[-1]
            module_name = '.'.join(module_parts[:-1])
            
            license_module = __import__(module_name, fromlist=[class_name])
            license_class = getattr(license_module, class_name)
            
            license = license_class()
            license.setLicense(license_path)
            
            logger.info(f"License applied successfully")
            return True
            
        except Exception as e:
            logger.error(f"Failed to apply license: {e}")
            return False

class BaseDocumentProcessor(abc.ABC):
    """Base class for document processors with clean interface"""
    
    def __init__(self, jvm_manager: AsposeJVMManager):
        self.jvm_manager = jvm_manager
        self._ensure_jvm_started()
        self._import_classes()
    
    @property
    @abc.abstractmethod
    def name(self) -> str:
        """Processor name"""
        pass
    
    @property
    @abc.abstractmethod
    def supported_extensions(self) -> List[str]:
        """Supported file extensions"""
        pass
    
    @property
    @abc.abstractmethod
    def required_jars(self) -> List[str]:
        """Required JAR file patterns"""
        pass
    
    @property
    @abc.abstractmethod
    def license_class(self) -> str:
        """License class name"""
        pass
    
    def _ensure_jvm_started(self):
        """Ensure JVM is started with required JARs"""
        if not self.jvm_manager.is_started():
            self.jvm_manager.start_jvm_with_jars(self.required_jars)
    
    @abc.abstractmethod
    def _import_classes(self):
        """Import required Java classes"""
        pass
    
    def supports_file(self, file_path: str) -> bool:
        """Check if processor supports the file"""
        return Path(file_path).suffix.lower() in self.supported_extensions
    
    @abc.abstractmethod
    def convert_to_markdown(self, input_path: str, output_path: str, 
                           options: ConversionOptions) -> str:
        """Convert document to Markdown"""
        pass
    
    @abc.abstractmethod
    def convert_to_json(self, input_path: str, output_path: str, 
                       options: ConversionOptions) -> str:
        """Convert document to JSON"""
        pass
    
    @abc.abstractmethod
    def convert_to_html(self, input_path: str, output_path: str, 
                       options: ConversionOptions) -> str:
        """Convert document to HTML"""
        pass
    
    @abc.abstractmethod
    def get_document_info(self, input_path: str) -> Dict[str, Any]:
        """Get document information"""
        pass
    
    def apply_license(self, license_path: str) -> bool:
        """Apply license for this processor"""
        return self.jvm_manager.apply_license(license_path, self.license_class)

class AsposeDocumentProcessor:
    """Main document processor with clean, extensible architecture"""
    
    def __init__(self, license_path: Optional[str] = None):
        self.jvm_manager = AsposeJVMManager()
        self.license_path = license_path
        self.processors: Dict[str, BaseDocumentProcessor] = {}
        self._initialize_processors()
    
    def _initialize_processors(self):
        """Initialize available processors"""
        # Only initialize Aspose.Cells for now
        self._init_cells_processor()
        
        # Future processors can be added here:
        # self._init_pdf_processor()
        # self._init_words_processor()
        # self._init_slides_processor()
    
    def _init_cells_processor(self):
        """Initialize Aspose.Cells processor"""
        try:
            from aspose_cells_processor import AsposeCellsProcessor
            processor = AsposeCellsProcessor(self.jvm_manager)
            self.processors[processor.name] = processor
            logger.info(f"Initialized processor: {processor.name}")
            
            # Apply license if provided
            if self.license_path:
                processor.apply_license(self.license_path)
                
        except ImportError as e:
            logger.error(f"Failed to initialize Aspose.Cells processor: {e}")
            logger.error("Make sure aspose-cells JAR files are in the plugins directory")
        except Exception as e:
            logger.error(f"Error initializing Aspose.Cells processor: {e}")
    
    def get_processor_for_file(self, file_path: str) -> Optional[BaseDocumentProcessor]:
        """Get appropriate processor for file"""
        for processor in self.processors.values():
            if processor.supports_file(file_path):
                return processor
        return None
    
    def detect_document_type(self, file_path: str) -> Optional[str]:
        """Detect document type and return processor name"""
        processor = self.get_processor_for_file(file_path)
        return processor.name if processor else None
    
    def convert_document(self, input_path: str, output_format: OutputFormat, 
                        output_path: Optional[str] = None, 
                        options: Optional[ConversionOptions] = None) -> str:
        """Convert document to specified format"""
        
        # Validate input
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        # Get processor
        processor = self.get_processor_for_file(input_path)
        if not processor:
            supported_formats = self.get_supported_extensions()
            raise ValueError(f"No processor available for file: {input_path}. Supported: {supported_formats}")
        
        # Set default output path
        if not output_path:
            input_file = Path(input_path)
            output_path = str(input_file.with_suffix(f'.{output_format.value}'))
        
        # Set default options
        if options is None:
            options = ConversionOptions()
        
        # Route to appropriate conversion method
        try:
            if output_format == OutputFormat.MARKDOWN:
                return processor.convert_to_markdown(input_path, output_path, options)
            elif output_format == OutputFormat.JSON:
                return processor.convert_to_json(input_path, output_path, options)
            elif output_format == OutputFormat.HTML:
                return processor.convert_to_html(input_path, output_path, options)
            else:
                raise ValueError(f"Unsupported output format: {output_format}")
                
        except Exception as e:
            logger.error(f"Conversion failed: {e}")
            raise
    
    def get_document_info(self, input_path: str) -> Dict[str, Any]:
        """Get document information"""
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        processor = self.get_processor_for_file(input_path)
        if not processor:
            raise ValueError(f"No processor available for file: {input_path}")
        
        return processor.get_document_info(input_path)
    
    def get_supported_extensions(self) -> List[str]:
        """Get all supported file extensions"""
        extensions = []
        for processor in self.processors.values():
            extensions.extend(processor.supported_extensions)
        return sorted(list(set(extensions)))
    
    def list_processors(self) -> Dict[str, Dict[str, Any]]:
        """List available processors and their info"""
        info = {}
        for name, processor in self.processors.items():
            info[name] = {
                "name": processor.name,
                "supported_extensions": processor.supported_extensions,
                "required_jars": processor.required_jars,
                "status": "ready"
            }
        return info

def main():
    """Clean CLI interface"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Aspose Document Processor",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py convert input.xlsx markdown output.md
  python main.py convert input.csv json --include-metadata
  python main.py info document.xlsx
  python main.py list-formats
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Convert command
    convert_parser = subparsers.add_parser('convert', help='Convert document')
    convert_parser.add_argument('input', help='Input file path')
    convert_parser.add_argument('format', choices=['markdown', 'json', 'html'], 
                               help='Output format')
    convert_parser.add_argument('output', nargs='?', help='Output file path (optional)')
    convert_parser.add_argument('--license', help='License file path')
    convert_parser.add_argument('--no-metadata', action='store_true', 
                               help='Exclude metadata')
    convert_parser.add_argument('--no-formatting', action='store_true',
                               help='Don\'t preserve formatting')
    convert_parser.add_argument('--include-hidden', action='store_true',
                               help='Include hidden worksheets')
    
    # Info command
    info_parser = subparsers.add_parser('info', help='Get document information')
    info_parser.add_argument('input', help='Input file path')
    
    # List commands
    subparsers.add_parser('list-formats', help='List supported formats')
    subparsers.add_parser('list-processors', help='List available processors')
    
    # Legacy support for simple usage
    if len(sys.argv) >= 3 and sys.argv[1] in ['markdown', 'json', 'html', 'info']:
        old_format = sys.argv[1]
        if old_format == 'info':
            sys.argv = [sys.argv[0], 'info', sys.argv[2]]
        else:
            new_args = [sys.argv[0], 'convert', sys.argv[2], old_format]
            if len(sys.argv) > 3:
                new_args.append(sys.argv[3])
            sys.argv = new_args
    
    args = parser.parse_args()
    
    if not args.command:
        parser.print_help()
        return
    
    try:
        processor = AsposeDocumentProcessor(license_path=getattr(args, 'license', None))
        
        if args.command == 'convert':
            # Create conversion options
            options = ConversionOptions(
                include_metadata=not getattr(args, 'no_metadata', False),
                preserve_formatting=not getattr(args, 'no_formatting', False),
                include_hidden_sheets=getattr(args, 'include_hidden', False)
            )
            
            output_format = OutputFormat(args.format)
            result = processor.convert_document(args.input, output_format, 
                                              getattr(args, 'output', None), options)
            print(f"✅ Conversion completed: {result}")
        
        elif args.command == 'info':
            info = processor.get_document_info(args.input)
            print(json.dumps(info, indent=2))
        
        elif args.command == 'list-formats':
            extensions = processor.get_supported_extensions()
            print("Supported file extensions:")
            for ext in extensions:
                print(f"  {ext}")
        
        elif args.command == 'list-processors':
            processors = processor.list_processors()
            print("Available processors:")
            for name, info in processors.items():
                print(f"  {name}:")
                print(f"    Extensions: {', '.join(info['supported_extensions'])}")
                print(f"    Status: {info['status']}")
    
    except Exception as e:
        logger.error(f"Error: {e}")
        print(f"❌ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()