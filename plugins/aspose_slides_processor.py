#!/usr/bin/env python3
"""
Aspose.Slides Processor
Handles PowerPoint (PPT/PPTX/ODP) document processing via Aspose.Slides for Java
"""

import os
import json
from pathlib import Path
from typing import Dict, Any, List
import logging

from main import BaseDocumentProcessor, ConversionOptions

logger = logging.getLogger(__name__)


class AsposeSlidesProcessor(BaseDocumentProcessor):
    """Processor for presentations using Aspose.Slides for Java"""

    # ---------- Required Base Properties ----------
    @property
    def name(self) -> str:
        return "slides"

    @property
    def supported_extensions(self) -> List[str]:
        return [".pptx", ".ppt", ".odp"]

    @property
    def required_jars(self) -> List[str]:
        # (Not used now that we always load all jars, but kept for consistency)
        return ["aspose-slides*.jar"]

    @property
    def license_class(self) -> str:
        return "com.aspose.slides.License"

    # ---------- Bootstrap Java Types ----------
    def _import_classes(self):
        """Import required Aspose.Slides classes"""
        try:
            from com.aspose.slides import Presentation, SaveFormat
            self.Presentation = Presentation
            self.SaveFormat = SaveFormat
            logger.info("Aspose.Slides classes imported successfully")
        except Exception as e:
            logger.error(f"Failed to import Aspose.Slides classes: {e}")
            raise

    # ---------- Conversions ----------
    def convert_to_markdown(self, input_path: str, output_path: str,
                            options: ConversionOptions) -> str:
        """
        Convert a presentation to Markdown.
        Strategy:
          - Title line with file name (and metadata if enabled)
          - Each slide outputs as a section with basic text content extracted from text frames
        """
        try:
            pres = self.Presentation(input_path)

            lines = []
            # Header with metadata
            if options.include_metadata:
                props = pres.getDocumentProperties()
                title = props.getTitle() if props.getTitle() else Path(input_path).name
                lines.append(f"# {title}")
                if props.getAuthor():
                    lines.append(f"**Author:** {props.getAuthor()}")
                if props.getCompany():
                    lines.append(f"**Company:** {props.getCompany()}")
                if props.getCreatedTime():
                    lines.append(f"**Created:** {props.getCreatedTime()}")
                lines.extend(["", "---", ""])
            else:
                lines.append(f"# {Path(input_path).name}")
                lines.append("")

            # Slides
            slides = pres.getSlides()
            for i in range(slides.size()):
                slide = slides.get_Item(i)
                lines.append(f"## Slide {i+1}")
                lines.append("")

                shapes = slide.getShapes()
                had_text = False
                for s in range(shapes.size()):
                    shape = shapes.get_Item(s)
                    # Many shape types can contain text; TextFrame property is the common path
                    try:
                        tf = shape.getTextFrame()
                        if tf is not None:
                            text = tf.getText()
                            if text is not None and str(text).strip():
                                lines.append(f"- {text.strip()}")
                                had_text = True
                    except Exception:
                        # Non-text shapes or ones without a text frame
                        continue

                if not had_text:
                    lines.append("_(No textual content detected)_")

                lines.append("")

            with open(output_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines))

            logger.info(f"Converted {input_path} to Markdown: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Slides → Markdown conversion failed: {e}")
            raise

    def convert_to_json(self, input_path: str, output_path: str,
                        options: ConversionOptions) -> str:
        """
        Convert a presentation to a JSON structure:
        {
          "document_type": "presentation",
          "processor": "slides",
          "slide_count": N,
          "metadata": {...},     # optional
          "slides": [
            {"index": 0, "shapes": [{"type": "text", "text": "..."}]}
          ]
        }
        """
        try:
            pres = self.Presentation(input_path)
            data: Dict[str, Any] = {
                "document_type": "presentation",
                "processor": self.name,
                "slide_count": pres.getSlides().size(),
                "slides": []
            }

            if options.include_metadata:
                try:
                    props = pres.getDocumentProperties()
                    data["metadata"] = {
                        "title": str(props.getTitle()) if props.getTitle() else None,
                        "author": str(props.getAuthor()) if props.getAuthor() else None,
                        "company": str(props.getCompany()) if props.getCompany() else None,
                        "created": str(props.getCreatedTime()) if props.getCreatedTime() else None,
                        "modified": str(props.getLastSavedTime()) if props.getLastSavedTime() else None,
                    }
                except Exception:
                    data["metadata"] = {}

            slides = pres.getSlides()
            for i in range(slides.size()):
                slide = slides.get_Item(i)
                shapes_json: List[Dict[str, Any]] = []

                shapes = slide.getShapes()
                for s in range(shapes.size()):
                    shape = shapes.get_Item(s)
                    entry: Dict[str, Any] = {"type": None}

                    # Text
                    try:
                        tf = shape.getTextFrame()
                        if tf is not None:
                            text_val = tf.getText()
                            if text_val is not None:
                                entry["type"] = "text"
                                entry["text"] = str(text_val)
                                shapes_json.append(entry)
                                continue
                    except Exception:
                        pass

                    # Fallback for non-text shapes (very generic)
                    entry["type"] = "shape"
                    shapes_json.append(entry)

                data["slides"].append({
                    "index": i,
                    "shapes": shapes_json
                })

            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            logger.info(f"Converted {input_path} to JSON: {output_path}")
            return output_path

        except Exception as e:
            logger.error(f"Slides → JSON conversion failed: {e}")
            raise

    def convert_to_html(self, input_path: str, output_path: str,
                        options: ConversionOptions) -> str:
        """Export presentation to HTML using native Aspose renderer."""
        try:
            pres = self.Presentation(input_path)
            pres.save(output_path, self.SaveFormat.Html)
            logger.info(f"Converted {input_path} to HTML: {output_path}")
            return output_path
        except Exception as e:
            logger.error(f"Slides → HTML conversion failed: {e}")
            raise

    # ---------- Info ----------
    def get_document_info(self, input_path: str) -> Dict[str, Any]:
        """Return high-level info + metadata"""
        try:
            pres = self.Presentation(input_path)
            slides = pres.getSlides()
            try:
                props = pres.getDocumentProperties()
            except Exception:
                props = None

            info: Dict[str, Any] = {
                "file_path": input_path,
                "file_name": Path(input_path).name,
                "file_size": os.path.getsize(input_path),
                "processor": self.name,
                "slide_count": slides.size(),
            }

            if props:
                info["metadata"] = {
                    "title": str(props.getTitle()) if props.getTitle() else "",
                    "author": str(props.getAuthor()) if props.getAuthor() else "",
                    "company": str(props.getCompany()) if props.getCompany() else "",
                    "created": str(props.getCreatedTime()) if props.getCreatedTime() else None,
                    "modified": str(props.getLastSavedTime()) if props.getLastSavedTime() else None,
                }

            # Per-slide quick stats (count shapes, text shapes)
            slides_info: List[Dict[str, Any]] = []
            for i in range(slides.size()):
                slide = slides.get_Item(i)
                shapes = slide.getShapes()
                text_shapes = 0
                for s in range(shapes.size()):
                    shape = shapes.get_Item(s)
                    try:
                        if shape.getTextFrame() is not None:
                            text_shapes += 1
                    except Exception:
                        pass

                slides_info.append({
                    "index": i,
                    "shapes_count": shapes.size(),
                    "text_shapes_count": text_shapes
                })

            info["slides"] = slides_info
            return info

        except Exception as e:
            logger.error(f"Error getting slides document info: {e}")
            raise

    # ---------- Batch ----------
    def batch_convert(self, input_dir: str, output_dir: str, output_format: str,
                      options: ConversionOptions = None) -> List[str]:
        """Batch convert all supported files in a directory"""
        from main import OutputFormat

        if options is None:
            options = ConversionOptions()

        in_dir = Path(input_dir)
        out_dir = Path(output_dir)
        out_dir.mkdir(exist_ok=True)

        converted: List[str] = []
        for fp in in_dir.iterdir():
            if not fp.is_file():
                continue
            if Path(fp).suffix.lower() not in self.supported_extensions:
                continue

            try:
                out_file = out_dir / f"{fp.stem}.{output_format}"
                if output_format == OutputFormat.MARKDOWN.value:
                    res = self.convert_to_markdown(str(fp), str(out_file), options)
                elif output_format == OutputFormat.JSON.value:
                    res = self.convert_to_json(str(fp), str(out_file), options)
                elif output_format == OutputFormat.HTML.value:
                    res = self.convert_to_html(str(fp), str(out_file), options)
                else:
                    logger.warning(f"Unsupported output format: {output_format}")
                    continue

                converted.append(res)
                logger.info(f"Batch converted: {fp.name} -> {out_file.name}")
            except Exception as e:
                logger.error(f"Failed to convert {fp.name}: {e}")

        return converted
