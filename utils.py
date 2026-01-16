"""
Utility functions and service layer for XML-Excel conversion
Contains ALL conversion logic - NO dependency on main.py
"""

import os
import sys
from pathlib import Path
from typing import Optional
import logging
import xml.etree.ElementTree as ET
from xml.dom import minidom
from collections import defaultdict

try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    raise ImportError("openpyxl is required. Install it with: pip install openpyxl")

logger = logging.getLogger(__name__)


# =============================================================================
# CUSTOM EXCEPTIONS
# =============================================================================

class ConversionError(Exception):
    """Base exception for conversion errors"""
    pass


class ValidationError(ConversionError):
    """Exception for validation errors"""
    pass


class FileProcessingError(ConversionError):
    """Exception for file processing errors"""
    pass


# =============================================================================
# XML TO EXCEL CONVERSION FUNCTIONS
# =============================================================================

def analyze_xml_structure(root):
    """Analyze XML structure to determine if multiple sheets are needed."""
    children = list(root)
    
    if not children:
        return {"Sheet1": [root]}
    
    tag_groups = defaultdict(list)
    for child in children:
        tag_groups[child.tag].append(child)
    
    if len(tag_groups) == 1:
        tag_name = list(tag_groups.keys())[0]
        return {tag_name: tag_groups[tag_name]}
    
    sheets = {}
    for tag, elements in tag_groups.items():
        sheets[tag] = elements
    
    return sheets


def extract_element_data(element):
    """Extract data from an XML element into a flat dictionary."""
    data = {}
    
    # Add attributes with @ prefix
    for attr, value in element.attrib.items():
        data[f"@{attr}"] = value
    
    # Add child elements
    for child in element:
        child_text = child.text.strip() if child.text else ""
        
        if len(child) > 0:
            nested_data = extract_element_data(child)
            for key, value in nested_data.items():
                data[f"{child.tag}.{key}"] = value
        else:
            if child.tag in data:
                i = 2
                while f"{child.tag}_{i}" in data:
                    i += 1
                data[f"{child.tag}_{i}"] = child_text
            else:
                data[child.tag] = child_text
            
            for attr, value in child.attrib.items():
                data[f"{child.tag}.@{attr}"] = value
    
    if element.text and element.text.strip() and len(element) == 0:
        data["_text"] = element.text.strip()
    
    return data


def get_all_columns(elements):
    """Get all unique column names from a list of elements."""
    columns = []
    seen = set()
    
    for element in elements:
        data = extract_element_data(element)
        for key in data.keys():
            if key not in seen:
                columns.append(key)
                seen.add(key)
    
    return columns


def style_header(ws, num_columns):
    """Apply styling to the header row."""
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for col in range(1, num_columns + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border


def auto_adjust_column_width(ws):
    """Automatically adjust column widths based on content."""
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        
        for cell in column_cells:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = max(adjusted_width, 10)


def xml_to_excel_internal(xml_path, output_path=None):
    """Convert XML to Excel - internal implementation."""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except ET.ParseError as e:
        raise ValueError(f"Invalid XML file: {e}")
    except FileNotFoundError:
        raise FileNotFoundError(f"XML file not found: {xml_path}")
    
    sheets_data = analyze_xml_structure(root)
    
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    for sheet_name, elements in sheets_data.items():
        safe_sheet_name = sheet_name[:31]
        for char in [':', '/', '\\', '?', '*', '[', ']']:
            safe_sheet_name = safe_sheet_name.replace(char, "_")
        
        ws = wb.create_sheet(title=safe_sheet_name)
        columns = get_all_columns(elements)
        
        if not columns:
            ws.cell(row=1, column=1, value="No data found")
            continue
        
        for col_idx, col_name in enumerate(columns, 1):
            ws.cell(row=1, column=col_idx, value=col_name)
        
        style_header(ws, len(columns))
        
        for row_idx, element in enumerate(elements, 2):
            data = extract_element_data(element)
            for col_idx, col_name in enumerate(columns, 1):
                value = data.get(col_name, "")
                ws.cell(row=row_idx, column=col_idx, value=value)
        
        auto_adjust_column_width(ws)
        ws.freeze_panes = "A2"
    
    if output_path is None:
        output_path = Path(xml_path).with_suffix(".xlsx")
    
    wb.save(output_path)
    return str(output_path)


# =============================================================================
# EXCEL TO XML CONVERSION FUNCTIONS
# =============================================================================

def sanitize_tag_name(name):
    """Sanitize a string to be a valid XML tag name."""
    if not name:
        return "element"
    
    name = str(name).strip()
    
    sanitized = ""
    for i, char in enumerate(name):
        if i == 0:
            if char.isalpha() or char == '_':
                sanitized += char
            else:
                sanitized += "_" + char if char.isalnum() else "_"
        else:
            if char.isalnum() or char in ['_', '-', '.']:
                sanitized += char
            else:
                sanitized += "_"
    
    if not sanitized or not (sanitized[0].isalpha() or sanitized[0] == '_'):
        sanitized = "element_" + sanitized
    
    if sanitized.lower().startswith("xml"):
        sanitized = "_" + sanitized
    
    return sanitized


def parse_column_structure(columns):
    """Parse column names to understand nested structure."""
    parsed = []
    for col in columns:
        col = str(col)
        if col.startswith("@"):
            parsed.append((col[1:], True, []))
        elif ".@" in col:
            parts = col.split(".@")
            path = parts[0].split(".")
            attr_name = parts[1]
            parsed.append((attr_name, True, path))
        elif "." in col:
            parts = col.split(".")
            parsed.append((parts[-1], False, parts[:-1]))
        else:
            parsed.append((col, False, []))
    
    return parsed


def row_to_xml_element(row_data, columns, element_tag):
    """Convert a row of data to an XML element."""
    element = ET.Element(sanitize_tag_name(element_tag))
    nested_elements = {}
    
    parsed_columns = parse_column_structure(columns)
    
    for (col_name, is_attr, nested_path), orig_col in zip(parsed_columns, columns):
        value = row_data.get(orig_col, "")
        if value is None or value == "":
            continue
        
        value = str(value)
        
        if not nested_path:
            if is_attr:
                element.set(sanitize_tag_name(col_name), value)
            else:
                child = ET.SubElement(element, sanitize_tag_name(col_name))
                child.text = value
        else:
            path_key = ".".join(nested_path)
            
            if path_key not in nested_elements:
                current = element
                for part in nested_path:
                    found = current.find(sanitize_tag_name(part))
                    if found is None:
                        found = ET.SubElement(current, sanitize_tag_name(part))
                    current = found
                nested_elements[path_key] = current
            
            target = nested_elements[path_key]
            
            if is_attr:
                target.set(sanitize_tag_name(col_name), value)
            else:
                child = ET.SubElement(target, sanitize_tag_name(col_name))
                child.text = value
    
    return element


def prettify_xml(element):
    """Return a pretty-printed XML string."""
    rough_string = ET.tostring(element, encoding='unicode')
    reparsed = minidom.parseString(rough_string)
    return reparsed.toprettyxml(indent="    ")


def detect_regions(ws):
    """Detect contiguous blocks of data in a worksheet."""
    regions = []
    visited = set()
    
    max_row = ws.max_row
    max_col = ws.max_column
    
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if (r, c) not in visited and ws.cell(row=r, column=c).value is not None:
                min_r, max_r, min_c, max_c = r, r, c, c
                
                stack = [(r, c)]
                visited.add((r, c))
                
                while stack:
                    curr_r, curr_c = stack.pop()
                    min_r = min(min_r, curr_r)
                    max_r = max(max_r, curr_r)
                    min_c = min(min_c, curr_c)
                    max_c = max(max_c, curr_c)
                    
                    for dr, dc in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
                        nr, nc = curr_r + dr, curr_c + dc
                        if 1 <= nr <= max_row and 1 <= nc <= max_col:
                            if (nr, nc) not in visited and ws.cell(row=nr, column=nc).value is not None:
                                visited.add((nr, nc))
                                stack.append((nr, nc))
                
                regions.append((min_r, max_r, min_c, max_c))
    
    return sorted(regions, key=lambda x: (x[0], x[2]))


def excel_to_xml_internal(excel_path, output_path=None, root_tag=None):
    """Convert Excel to XML - internal implementation."""
    try:
        wb = load_workbook(excel_path, data_only=True)
    except FileNotFoundError:
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    except Exception as e:
        raise ValueError(f"Invalid Excel file: {e}")
    
    if root_tag is None:
        root_tag = sanitize_tag_name(Path(excel_path).stem)
    
    root = ET.Element(root_tag)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        sheet_container = ET.SubElement(root, sanitize_tag_name(sheet_name))
        
        regions = detect_regions(ws)
        
        for i, (min_r, max_r, min_c, max_c) in enumerate(regions):
            width = max_c - min_c + 1
            height = max_r - min_r + 1
            
            if width == 2 and height <= 5:
                metadata = ET.SubElement(sheet_container, "Metadata")
                for r in range(min_r, max_r + 1):
                    k = str(ws.cell(row=r, column=min_c).value or f"field_{r}")
                    v = str(ws.cell(row=r, column=max_c).value or "")
                    item = ET.SubElement(metadata, sanitize_tag_name(k))
                    item.text = v
                continue
            
            headers = []
            col_indices = []
            for c in range(min_c, max_c + 1):
                val = ws.cell(row=min_r, column=c).value
                if val is not None:
                    headers.append(str(val))
                    col_indices.append(c)
            
            if not headers:
                continue
            
            record_tag = sheet_name.rstrip('s') if sheet_name.endswith('s') else "Item"
            record_tag = sanitize_tag_name(record_tag)
            
            table_container = ET.Element("Table")
            
            for r in range(min_r + 1, max_r + 1):
                row_data = {}
                has_data = False
                for header, c in zip(headers, col_indices):
                    val = ws.cell(row=r, column=c).value
                    if val is not None:
                        row_data[header] = val
                        has_data = True
                
                if has_data:
                    element = row_to_xml_element(row_data, headers, record_tag)
                    table_container.append(element)
            
            if len(table_container) > 0:
                sheet_container.append(table_container)
    
    if output_path is None:
        output_path = Path(excel_path).with_suffix(".xml")
    
    xml_string = prettify_xml(root)
    
    lines = xml_string.split('\n')
    if lines[0].startswith('<?xml'):
        lines = lines[1:]
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
        f.write('\n'.join(lines))
    
    return str(output_path)


# =============================================================================
# CONVERSION SERVICE CLASS
# =============================================================================

class ConversionService:
    """Service layer for handling XML-Excel conversions."""
    
    def __init__(self, max_file_size_mb: int = 50):
        self.max_file_size_bytes = max_file_size_mb * 1024 * 1024
        logger.info(f"ConversionService initialized (max file size: {max_file_size_mb}MB)")
    
    def validate_file(self, file_path: str, expected_ext: Optional[str] = None) -> None:
        """Validate file existence, size, and extension."""
        path = Path(file_path)
        
        if not path.exists():
            raise ValidationError(f"File not found: {file_path}")
        
        if not path.is_file():
            raise ValidationError(f"Path is not a file: {file_path}")
        
        file_size = os.path.getsize(file_path)
        if file_size == 0:
            raise ValidationError("File is empty")
        
        if file_size > self.max_file_size_bytes:
            size_mb = file_size / (1024 * 1024)
            max_mb = self.max_file_size_bytes / (1024 * 1024)
            raise ValidationError(
                f"File too large: {size_mb:.2f}MB (max: {max_mb}MB)"
            )
        
        if expected_ext:
            if not isinstance(expected_ext, (list, tuple)):
                expected_ext = [expected_ext]
            
            if path.suffix.lower() not in [e.lower() for e in expected_ext]:
                raise ValidationError(
                    f"Invalid file extension: {path.suffix}. Expected: {', '.join(expected_ext)}"
                )
        
        logger.info(f"File validated: {file_path} ({file_size} bytes)")
    
    def xml_to_excel(self, xml_path: str, output_path: Optional[str] = None) -> str:
        """Convert XML file to Excel with validation and error handling."""
        try:
            self.validate_file(xml_path, expected_ext='.xml')
            logger.info(f"Starting XML to Excel conversion: {xml_path}")
            
            result_path = xml_to_excel_internal(xml_path, output_path)
            
            if not os.path.exists(result_path):
                raise FileProcessingError("Output file was not created")
            
            logger.info(f"Conversion successful: {result_path}")
            return result_path
            
        except ValidationError:
            raise
        except FileNotFoundError as e:
            raise ValidationError(f"File not found: {str(e)}")
        except ValueError as e:
            raise ConversionError(f"Invalid XML format: {str(e)}")
        except Exception as e:
            logger.error(f"XML to Excel conversion failed: {str(e)}", exc_info=True)
            raise ConversionError(f"Conversion failed: {str(e)}")
    
    def excel_to_xml(self, excel_path: str, output_path: Optional[str] = None, 
                     root_tag: Optional[str] = None) -> str:
        """Convert Excel file to XML with validation and error handling."""
        try:
            self.validate_file(excel_path, expected_ext=['.xlsx', '.xls', '.xlsm'])
            logger.info(f"Starting Excel to XML conversion: {excel_path}")
            
            result_path = excel_to_xml_internal(excel_path, output_path, root_tag)
            
            if not os.path.exists(result_path):
                raise FileProcessingError("Output file was not created")
            
            logger.info(f"Conversion successful: {result_path}")
            return result_path
            
        except ValidationError:
            raise
        except FileNotFoundError as e:
            raise ValidationError(f"File not found: {str(e)}")
        except ValueError as e:
            raise ConversionError(f"Invalid Excel format: {str(e)}")
        except Exception as e:
            logger.error(f"Excel to XML conversion failed: {str(e)}", exc_info=True)
            raise ConversionError(f"Conversion failed: {str(e)}")
    
    def get_file_info(self, file_path: str) -> dict:
        """Get information about a file."""
        path = Path(file_path)
        
        if not path.exists():
            raise ValidationError(f"File not found: {file_path}")
        
        file_size = os.path.getsize(file_path)
        
        return {
            "filename": path.name,
            "extension": path.suffix,
            "size_bytes": file_size,
            "size_mb": round(file_size / (1024 * 1024), 2),
            "is_xml": path.suffix.lower() == '.xml',
            "is_excel": path.suffix.lower() in ['.xlsx', '.xls', '.xlsm']
        }