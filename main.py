#!/usr/bin/env python3
"""
Aspose.Total-for-Python-via-Java Document Processor (Clean, Extensible)
- Always loads ALL JARs in plugins/ before initializing any processors
"""

import os
import sys
import json
import abc
from pathlib import Path
from typing import Optional, Dict, Any, List
from dataclasses import dataclass
from enum import Enum
import importlib
import importlib.util
import logging

# --- Paths ---
ROOT_DIR = Path(__file__).parent
PLUGINS_DIR = ROOT_DIR / "plugins"

# Make `plugins/` importable for Python-based plugins (e.g., aspose_*_processor.py)
sys.path.insert(0, str(PLUGINS_DIR))

# --- Logging ---
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


# === Public API Types ===
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


# === JVM Manager (Singleton) ===
class AsposeJVMManager:
    """Centralized JVM management for Aspose processors"""

    _instance = None
    _jvm_started = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if not hasattr(self, "_initialized"):
            self.plugins_dir = PLUGINS_DIR
            self._initialized = True

    def start_jvm_with_jars(self, jar_patterns: Optional[List[str]] = None) -> bool:
        """
        Start JVM with ALL JARs found in plugins/.
        (jar_patterns is ignored intentionally; we always load everything to satisfy cross-deps)
        """
        if self._jvm_started:
            logger.info("JVM already started")
            return True

        try:
            import jpype
            import jpype.imports
        except ImportError:
            logger.error("JPype1 not found. Install with: pip install JPype1")
            raise

        # Collect ALL JARs in plugins/
        all_jars = list(self.plugins_dir.glob("*.jar"))
        if not all_jars:
            raise FileNotFoundError(f"No JAR files found in {self.plugins_dir}")

        # Build classpath (dedup + absolute)
        classpath = list({str(j.resolve()) for j in all_jars})

        try:
            jpype.startJVM(
                jpype.getDefaultJVMPath(),
                "-Djava.class.path=" + os.pathsep.join(classpath),
                "-Xmx4g",
                "-Dfile.encoding=UTF-8",
                "-XX:+UseG1GC",
                "-XX:MaxGCPauseMillis=200",
                convertStrings=False,
            )
            self._jvm_started = True
            logger.info(f"✅ JVM started successfully with {len(classpath)} JARs")
            return True
        except Exception as e:
            logger.error(f"Failed to start JVM: {e}")
            raise

    def is_started(self) -> bool:
        return self._jvm_started

    def apply_license(self, license_path: str, license_class_name: str) -> bool:
        """Apply Aspose license via the Java license class"""
        if not os.path.exists(license_path):
            logger.warning(f"License file not found: {license_path}")
            return False

        try:
            module_parts = license_class_name.split(".")
            class_name = module_parts[-1]
            module_name = ".".join(module_parts[:-1])

            license_module = __import__(module_name, fromlist=[class_name])
            license_class = getattr(license_module, class_name)

            license_obj = license_class()
            license_obj.setLicense(license_path)

            logger.info("License applied successfully")
            return True

        except Exception as e:
            logger.error(f"Failed to apply license: {e}")
            return False


# === Base Processor ===
class BaseDocumentProcessor(abc.ABC):
    """Base class for document processors with a clean interface"""

    def __init__(self, jvm_manager: AsposeJVMManager):
        self.jvm_manager = jvm_manager
        self._ensure_jvm_started()
        self._import_classes()

    @property
    @abc.abstractmethod
    def name(self) -> str: ...

    @property
    @abc.abstractmethod
    def supported_extensions(self) -> List[str]: ...

    @property
    @abc.abstractmethod
    def required_jars(self) -> List[str]: ...

    @property
    @abc.abstractmethod
    def license_class(self) -> str: ...

    def _ensure_jvm_started(self):
        """
        Ensure JVM is started with ALL jars.
        (We deliberately do NOT pass required_jars to force loading every jar first.)
        """
        if not self.jvm_manager.is_started():
            self.jvm_manager.start_jvm_with_jars()  # Load ALL jars

    @abc.abstractmethod
    def _import_classes(self): ...

    def supports_file(self, file_path: str) -> bool:
        return Path(file_path).suffix.lower() in self.supported_extensions

    @abc.abstractmethod
    def convert_to_markdown(self, input_path: str, output_path: str,
                            options: ConversionOptions) -> str: ...

    @abc.abstractmethod
    def convert_to_json(self, input_path: str, output_path: str,
                        options: ConversionOptions) -> str: ...

    @abc.abstractmethod
    def convert_to_html(self, input_path: str, output_path: str,
                        options: ConversionOptions) -> str: ...

    @abc.abstractmethod
    def get_document_info(self, input_path: str) -> Dict[str, Any]: ...

    def apply_license(self, license_path: str) -> bool:
        return self.jvm_manager.apply_license(license_path, self.license_class)


# === Central Dispatcher ===
class AsposeDocumentProcessor:
    """Main document processor (loads plugins, routes conversions)"""

    def __init__(self, license_path: Optional[str] = None):
        self.jvm_manager = AsposeJVMManager()
        # Start JVM *before* anything else to satisfy "load all jars first"
        self.jvm_manager.start_jvm_with_jars()

        self.license_path = license_path
        self.processors: Dict[str, BaseDocumentProcessor] = {}
        self._initialize_processors()

    # -- Plugin loading helpers --
    def _import_plugin_module(self, base_name: str):
        """
        Attempt to import plugin by:
          1) standard module name (e.g., 'aspose_cells_processor'),
          2) if that fails, try to load a file with hyphens (e.g., 'aspose-cells-processor.py') from plugins/.
        """
        try:
            return importlib.import_module(base_name)
        except ImportError:
            # Try hyphenated filename as a last resort
            hyphen_name = base_name.replace("_", "-") + ".py"
            candidate = PLUGINS_DIR / hyphen_name
            if candidate.exists():
                spec = importlib.util.spec_from_file_location(base_name, candidate)
                if spec and spec.loader:
                    module = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(module)
                    sys.modules[base_name] = module
                    return module
            raise

    def _initialize_processors(self):
        """Initialize available processors"""
        self._init_cells_processor()
        self._init_slides_processor()
        # Future:
        # self._init_words_processor()
        # self._init_pdf_processor()

    def _init_cells_processor(self):
        try:
            mod = self._import_plugin_module("aspose_cells_processor")
            cls = getattr(mod, "AsposeCellsProcessor")
            processor = cls(self.jvm_manager)
            self.processors[processor.name] = processor
            logger.info(f"Initialized processor: {processor.name}")
            if self.license_path:
                processor.apply_license(self.license_path)
        except ImportError as e:
            logger.warning(f"Cells processor not loaded: {e}")
        except Exception as e:
            logger.error(f"Error initializing Cells processor: {e}")

    def _init_slides_processor(self):
        try:
            mod = self._import_plugin_module("aspose_slides_processor")
            cls = getattr(mod, "AsposeSlidesProcessor")
            processor = cls(self.jvm_manager)
            self.processors[processor.name] = processor
            logger.info(f"Initialized processor: {processor.name}")
            if self.license_path:
                processor.apply_license(self.license_path)
        except ImportError as e:
            logger.warning(f"Slides processor not loaded: {e}")
        except Exception as e:
            logger.error(f"Error initializing Slides processor: {e}")

    # -- Routing & Utilities --
    def get_processor_for_file(self, file_path: str) -> Optional[BaseDocumentProcessor]:
        for processor in self.processors.values():
            if processor.supports_file(file_path):
                return processor
        return None

    def detect_document_type(self, file_path: str) -> Optional[str]:
        proc = self.get_processor_for_file(file_path)
        return proc.name if proc else None

    def convert_document(self, input_path: str, output_format: OutputFormat,
                         output_path: Optional[str] = None,
                         options: Optional[ConversionOptions] = None) -> str:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")

        processor = self.get_processor_for_file(input_path)
        if not processor:
            raise ValueError(f"No processor available for file: {input_path}. "
                             f"Supported: {self.get_supported_extensions()}")

        if output_path is None:
            input_file = Path(input_path)
            output_path = str(input_file.with_suffix(f".{output_format.value}"))

        if options is None:
            options = ConversionOptions()

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
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        processor = self.get_processor_for_file(input_path)
        if not processor:
            raise ValueError(f"No processor available for file: {input_path}")
        return processor.get_document_info(input_path)

    def get_supported_extensions(self) -> List[str]:
        exts: List[str] = []
        for p in self.processors.values():
            exts.extend(p.supported_extensions)
        return sorted(list(set(exts)))

    def list_processors(self) -> Dict[str, Dict[str, Any]]:
        info: Dict[str, Dict[str, Any]] = {}
        for name, p in self.processors.items():
            info[name] = {
                "name": p.name,
                "supported_extensions": p.supported_extensions,
                "required_jars": p.required_jars,
                "status": "ready",
            }
        return info


# === CLI ===
def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Aspose Document Processor (Python ↔ Java via JPype)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py convert input.xlsx markdown output.md
  python main.py convert input.csv json --include-metadata
  python main.py convert deck.pptx html
  python main.py info document.xlsx
  python main.py list-formats
  python main.py list-processors
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Convert
    convert_parser = subparsers.add_parser("convert", help="Convert document")
    convert_parser.add_argument("input", help="Input file path")
    convert_parser.add_argument("format", choices=["markdown", "json", "html"], help="Output format")
    convert_parser.add_argument("output", nargs="?", help="Output file path (optional)")
    convert_parser.add_argument("--license", help="License file path")
    convert_parser.add_argument("--no-metadata", action="store_true", help="Exclude metadata")
    convert_parser.add_argument("--no-formatting", action="store_true", help="Don't preserve formatting")
    convert_parser.add_argument("--include-hidden", action="store_true", help="Include hidden worksheets (Cells)")

    # Info
    info_parser = subparsers.add_parser("info", help="Get document information")
    info_parser.add_argument("input", help="Input file path")

    # Lists
    subparsers.add_parser("list-formats", help="List supported formats")
    subparsers.add_parser("list-processors", help="List available processors")

    # Legacy shim
    if len(sys.argv) >= 3 and sys.argv[1] in ["markdown", "json", "html", "info"]:
        old_format = sys.argv[1]
        if old_format == "info":
            sys.argv = [sys.argv[0], "info", sys.argv[2]]
        else:
            new_args = [sys.argv[0], "convert", sys.argv[2], old_format]
            if len(sys.argv) > 3:
                new_args.append(sys.argv[3])
            sys.argv = new_args

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    try:
        processor = AsposeDocumentProcessor(license_path=getattr(args, "license", None))

        if args.command == "convert":
            options = ConversionOptions(
                include_metadata=not getattr(args, "no_metadata", False),
                preserve_formatting=not getattr(args, "no_formatting", False),
                include_hidden_sheets=getattr(args, "include_hidden", False),
            )
            out_fmt = OutputFormat(args.format)
            result = processor.convert_document(args.input, out_fmt, getattr(args, "output", None), options)
            print(f"✅ Conversion completed: {result}")

        elif args.command == "info":
            info = processor.get_document_info(args.input)
            print(json.dumps(info, indent=2))

        elif args.command == "list-formats":
            exts = processor.get_supported_extensions()
            print("Supported file extensions:")
            for ext in exts:
                print(f"  {ext}")

        elif args.command == "list-processors":
            procs = processor.list_processors()
            print("Available processors:")
            for name, info in procs.items():
                print(f"  {name}:")
                print(f"    Extensions: {', '.join(info['supported_extensions'])}")
                print(f"    Status: {info['status']}")

    except Exception as e:
        logger.error(f"Error: {e}")
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
