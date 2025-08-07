#!/usr/bin/env python3
"""
Clean Aspose.Cells Processor
Focused, clean implementation for Excel document processing
"""

import os
import json
from pathlib import Path
from typing import Dict, Any, List
import logging

from main import BaseDocumentProcessor, ConversionOptions, AsposeJVMManager

logger = logging.getLogger(__name__)

class AsposeCellsProcessor(BaseDocumentProcessor):
    """Clean processor for Excel documents using Aspose.Cells for Java"""
    
    @property
    def name(self) -> str:
        return "cells"
    
    @property
    def supported_extensions(self) -> List[str]:
        return ['.xlsx', '.xls', '.csv', '.ods', '.xlsm', '.xlsb']
    
    @property
    def required_jars(self) -> List[str]:
        return ["aspose-cells*.jar"]
    
    @property
    def license_class(self) -> str:
        return "com.aspose.cells.License"
    
    def _import_classes(self):
        """Import required Aspose.Cells classes"""
        try:
            from com.aspose.cells import (
                Workbook, SaveFormat, HtmlSaveOptions, TxtSaveOptions,
                LoadOptions, FileFormatType
            )
            
            self.Workbook = Workbook
            self.SaveFormat = SaveFormat
            self.HtmlSaveOptions = HtmlSaveOptions
            self.TxtSaveOptions = TxtSaveOptions
            self.LoadOptions = LoadOptions
            self.FileFormatType = FileFormatType
            
            logger.info("Aspose.Cells classes imported successfully")
            
        except Exception as e:
            logger.error(f"Failed to import Aspose.Cells classes: {e}")
            raise
    
    def convert_to_markdown(self, input_path: str, output_path: str, 
                           options: ConversionOptions) -> str:
        """Convert Excel file to Markdown format"""
        try:
            workbook = self.Workbook(input_path)
            markdown_content = self._create_markdown(workbook, options)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(markdown_content)
            
            logger.info(f"Converted {input_path} to Markdown: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"Markdown conversion failed: {e}")
            raise
    
    def convert_to_json(self, input_path: str, output_path: str, 
                       options: ConversionOptions) -> str:
        """Convert Excel file to JSON format"""
        try:
            workbook = self.Workbook(input_path)
            json_data = self._create_json(workbook, options)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Converted {input_path} to JSON: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"JSON conversion failed: {e}")
            raise
    
    def convert_to_html(self, input_path: str, output_path: str, 
                       options: ConversionOptions) -> str:
        """Convert Excel file to HTML format"""
        try:
            workbook = self.Workbook(input_path)
            
            # Configure HTML options
            html_options = self.HtmlSaveOptions()
            html_options.setExportGridLines(True)
            html_options.setExportImagesAsBase64(options.export_images_as_base64)
            html_options.setExportHiddenWorksheet(options.include_hidden_sheets)
            
            if options.preserve_formatting:
                html_options.setExportSimilarBorderStyle(True)
                html_options.setAddTooltipText(True)
            
            workbook.save(output_path, html_options)
            
            # Add metadata if requested
            if options.include_metadata:
                self._add_html_metadata(output_path, workbook)
            
            logger.info(f"Converted {input_path} to HTML: {output_path}")
            return output_path
            
        except Exception as e:
            logger.error(f"HTML conversion failed: {e}")
            raise
    
    def get_document_info(self, input_path: str) -> Dict[str, Any]:
        """Get comprehensive Excel document information"""
        try:
            workbook = self.Workbook(input_path)
            
            # Basic info
            info = {
                "file_path": input_path,
                "file_name": Path(input_path).name,
                "file_size": os.path.getsize(input_path),
                "file_format": self._get_format_name(input_path),
                "processor": self.name,
                "worksheet_count": workbook.getWorksheets().getCount(),
                "worksheets": []
            }
            
            # Document properties
            if True:  # Always try to get metadata
                try:
                    properties = workbook.getBuiltInDocumentProperties()
                    info["metadata"] = {
                        "title": str(properties.getTitle()) if properties.getTitle() else "",
                        "author": str(properties.getAuthor()) if properties.getAuthor() else "",
                        "subject": str(properties.getSubject()) if properties.getSubject() else "",
                        "company": str(properties.getCompany()) if properties.getCompany() else "",
                        "created": str(properties.getCreatedTime()) if properties.getCreatedTime() else None,
                        "modified": str(properties.getLastSavedTime()) if properties.getLastSavedTime() else None,
                    }
                except Exception as e:
                    logger.debug(f"Could not read document properties: {e}")
                    info["metadata"] = {}
            
            # Protection status
            try:
                info["is_protected"] = workbook.isProtected()
            except:
                info["is_protected"] = False
            
            # Process worksheets
            total_data_cells = 0
            for i in range(workbook.getWorksheets().getCount()):
                worksheet = workbook.getWorksheets().get(i)
                cells = worksheet.getCells()
                
                max_row = cells.getMaxDataRow()
                max_col = cells.getMaxDataColumn()
                has_data = max_row >= 0
                
                worksheet_info = {
                    "name": worksheet.getName(),
                    "index": i,
                    "has_data": has_data,
                    "is_hidden": not worksheet.isVisible(),
                    "rows": max_row + 1 if has_data else 0,
                    "columns": max_col + 1 if has_data else 0,
                }
                
                if has_data:
                    data_cells = (max_row + 1) * (max_col + 1)
                    total_data_cells += data_cells
                    worksheet_info["data_cells"] = data_cells
                
                # Additional properties
                try:
                    worksheet_info["is_protected"] = worksheet.isProtected()
                    worksheet_info["charts_count"] = worksheet.getCharts().getCount()
                    worksheet_info["images_count"] = worksheet.getPictures().getCount()
                except:
                    pass
                
                info["worksheets"].append(worksheet_info)
            
            info["total_data_cells"] = total_data_cells
            
            logger.info(f"Retrieved document info for {input_path}")
            return info
            
        except Exception as e:
            logger.error(f"Error getting document info: {e}")
            raise
    
    def _get_format_name(self, input_path: str) -> str:
        """Get user-friendly format name"""
        ext = Path(input_path).suffix.lower()
        formats = {
            '.xlsx': 'Excel 2007+ (XLSX)',
            '.xls': 'Excel 97-2003 (XLS)', 
            '.xlsm': 'Excel Macro-Enabled (XLSM)',
            '.xlsb': 'Excel Binary (XLSB)',
            '.csv': 'Comma-Separated Values (CSV)',
            '.ods': 'OpenDocument Spreadsheet (ODS)'
        }
        return formats.get(ext, f'Unknown ({ext})')
    
    def _create_markdown(self, workbook, options: ConversionOptions) -> str:
        """Create Markdown content from workbook"""
        lines = []
        
        # Header with metadata
        if options.include_metadata:
            try:
                properties = workbook.getBuiltInDocumentProperties()
                title = properties.getTitle() if properties.getTitle() else "Excel Document"
                lines.extend([
                    f"# {title}",
                    "",
                ])
                
                if properties.getAuthor():
                    lines.append(f"**Author:** {properties.getAuthor()}")
                if properties.getCreatedTime():
                    lines.append(f"**Created:** {properties.getCreatedTime()}")
                
                lines.extend(["", "---", ""])
            except:
                lines.extend(["# Excel Document", ""])
        else:
            lines.extend(["# Excel Document", ""])
        
        # Process worksheets
        worksheet_count = workbook.getWorksheets().getCount()
        for i in range(worksheet_count):
            worksheet = workbook.getWorksheets().get(i)
            cells = worksheet.getCells()
            
            # Skip hidden sheets unless requested
            if not worksheet.isVisible() and not options.include_hidden_sheets:
                continue
            
            # Worksheet header
            worksheet_name = worksheet.getName()
            if worksheet_count > 1:
                lines.append(f"## {worksheet_name}")
            else:
                lines.append(f"## Data")
            lines.append("")
            
            # Check for data
            if cells.getMaxDataRow() < 0:
                lines.extend(["*No data in this worksheet*", ""])
                continue
            
            # Create table
            max_row = cells.getMaxDataRow() + 1
            max_col = cells.getMaxDataColumn() + 1
            
            table_data = []
            for row in range(max_row):
                row_data = []
                for col in range(max_col):
                    cell = cells.get(row, col)
                    value = ""
                    
                    if cell.getValue() is not None:
                        value = str(cell.getValue())
                        
                        # Apply formatting if requested
                        if options.preserve_formatting:
                            try:
                                style = cell.getStyle()
                                if style.getFont().isBold():
                                    value = f"**{value}**"
                                if style.getFont().isItalic():
                                    value = f"*{value}*"
                            except:
                                pass
                    
                    # Clean value for markdown
                    value = value.replace("|", "\\|").replace("\n", " ").replace("\r", " ")
                    row_data.append(value)
                
                table_data.append(row_data)
            
            # Write table
            if table_data:
                # Header row
                lines.append("| " + " | ".join(table_data[0]) + " |")
                lines.append("| " + " | ".join(["---"] * max_col) + " |")
                
                # Data rows
                for row_data in table_data[1:]:
                    lines.append("| " + " | ".join(row_data) + " |")
                
                lines.append("")
        
        return "\n".join(lines)
    
    def _create_json(self, workbook, options: ConversionOptions) -> Dict[str, Any]:
        """Create JSON structure from workbook"""
        data = {
            "document_type": "excel",
            "processor": self.name,
            "worksheet_count": workbook.getWorksheets().getCount(),
            "worksheets": []
        }
        
        # Add metadata
        if options.include_metadata:
            try:
                properties = workbook.getBuiltInDocumentProperties()
                data["metadata"] = {
                    "title": str(properties.getTitle()) if properties.getTitle() else None,
                    "author": str(properties.getAuthor()) if properties.getAuthor() else None,
                    "subject": str(properties.getSubject()) if properties.getSubject() else None,
                    "created": str(properties.getCreatedTime()) if properties.getCreatedTime() else None,
                    "modified": str(properties.getLastSavedTime()) if properties.getLastSavedTime() else None,
                }
            except:
                data["metadata"] = {}
        
        # Process worksheets
        for i in range(workbook.getWorksheets().getCount()):
            worksheet = workbook.getWorksheets().get(i)
            cells = worksheet.getCells()
            
            # Skip hidden sheets unless requested
            if not worksheet.isVisible() and not options.include_hidden_sheets:
                continue
            
            worksheet_data = {
                "name": worksheet.getName(),
                "index": i,
                "is_hidden": not worksheet.isVisible(),
                "data": []
            }
            
            # Extract data
            if cells.getMaxDataRow() >= 0:
                max_row = cells.getMaxDataRow() + 1
                max_col = cells.getMaxDataColumn() + 1
                
                for row in range(max_row):
                    row_data = []
                    for col in range(max_col):
                        cell = cells.get(row, col)
                        value = None
                        
                        if cell.getValue() is not None:
                            value = str(cell.getValue())
                            
                            # Include formatting if requested
                            if options.preserve_formatting:
                                try:
                                    style = cell.getStyle()
                                    cell_data = {
                                        "value": value,
                                        "bold": style.getFont().isBold(),
                                        "italic": style.getFont().isItalic()
                                    }
                                    value = cell_data
                                except:
                                    pass
                        
                        row_data.append(value)
                    worksheet_data["data"].append(row_data)
            
            data["worksheets"].append(worksheet_data)
        
        return data
    
    def _add_html_metadata(self, html_path: str, workbook):
        """Add metadata as HTML comment"""
        try:
            with open(html_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            properties = workbook.getBuiltInDocumentProperties()
            metadata_comment = f"""<!-- 
Document Metadata:
- Title: {properties.getTitle() if properties.getTitle() else 'N/A'}
- Author: {properties.getAuthor() if properties.getAuthor() else 'N/A'}
- Created: {properties.getCreatedTime() if properties.getCreatedTime() else 'N/A'}
- Modified: {properties.getLastSavedTime() if properties.getLastSavedTime() else 'N/A'}
- Processor: Aspose.Cells via Java
-->
"""
            
            # Insert comment after <html> tag
            if '<html>' in content:
                content = content.replace('<html>', f'<html>\n{metadata_comment}', 1)
            else:
                content = f'{metadata_comment}\n{content}'
            
            with open(html_path, 'w', encoding='utf-8') as f:
                f.write(content)
                
        except Exception as e:
            logger.debug(f"Could not add HTML metadata: {e}")

    def batch_convert(self, input_dir: str, output_dir: str, output_format: str, 
                     options: ConversionOptions = None) -> List[str]:
        """Batch convert all supported files in a directory"""
        from main import OutputFormat
        
        if options is None:
            options = ConversionOptions()
        
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True)
        
        converted_files = []
        
        for file_path in input_path.iterdir():
            if file_path.is_file() and self.supports_file(str(file_path)):
                try:
                    output_file = output_path / f"{file_path.stem}.{output_format}"
                    
                    if output_format == "markdown":
                        result = self.convert_to_markdown(str(file_path), str(output_file), options)
                    elif output_format == "json":
                        result = self.convert_to_json(str(file_path), str(output_file), options)
                    elif output_format == "html":
                        result = self.convert_to_html(str(file_path), str(output_file), options)
                    else:
                        logger.warning(f"Unsupported output format: {output_format}")
                        continue
                    
                    converted_files.append(result)
                    logger.info(f"Batch converted: {file_path.name} -> {output_file.name}")
                    
                except Exception as e:
                    logger.error(f"Failed to convert {file_path.name}: {e}")
        
        return converted_files