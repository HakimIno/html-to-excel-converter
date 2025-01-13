import sys
import json
import xlsxwriter
from selectolax.parser import HTMLParser
from typing import Dict, List, Set, Tuple, Union, Optional
import re
import time
import logging
from dataclasses import dataclass
from functools import lru_cache
import html
import gc
from io import BytesIO

logger = logging.getLogger(__name__)

class PerformanceTimer:
    _timings = {}  # Class variable to store all timings
    
    def __init__(self, name, min_duration=0.001):
        self.name = name
        self.start_time = None
        self.min_duration = min_duration
        
    def __enter__(self):
        self.start_time = time.time()
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        duration = time.time() - self.start_time
        if duration >= self.min_duration:
            if self.name not in self._timings:
                self._timings[self.name] = []
            self._timings[self.name].append(duration)
    
    @classmethod
    def print_summary(cls):
        logger.info("Performance Summary:")
        for name, durations in sorted(cls._timings.items()):
            avg_time = sum(durations) / len(durations)
            total_time = sum(durations)
            count = len(durations)
            logger.info(f"  {name}:")
            logger.info(f"    Count: {count}")
            logger.info(f"    Average: {avg_time:.3f} seconds")
            logger.info(f"    Total: {total_time:.3f} seconds")
        cls._timings.clear()  # Clear timings after printing

@dataclass
class CellStyle:
    """Represents Excel cell style properties"""
    font_name: str = None
    font_size: float = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: str = None
    bg_color: str = None
    border_top: int = 0
    border_right: int = 0
    border_bottom: int = 0
    border_left: int = 0
    align: str = 'center'
    valign: str = 'vcenter'
    text_wrap: bool = True
    
    def to_excel_format(self) -> dict:
        """Convert to Excel format dictionary"""
        format_dict = {}
        
        if self.font_name:
            format_dict['font_name'] = self.font_name
        if self.font_size:
            format_dict['font_size'] = self.font_size
        if self.bold:
            format_dict['bold'] = True
        if self.italic:
            format_dict['italic'] = True
        if self.underline:
            format_dict['underline'] = True
        if self.font_color:
            format_dict['font_color'] = self.font_color
        if self.bg_color:
            format_dict['bg_color'] = self.bg_color
        if self.border_top:
            format_dict['top'] = self.border_top
        if self.border_right:
            format_dict['right'] = self.border_right
        if self.border_bottom:
            format_dict['bottom'] = self.border_bottom
        if self.border_left:
            format_dict['left'] = self.border_left
        if self.align:
            format_dict['align'] = self.align
        if self.valign:
            format_dict['valign'] = self.valign
        if self.text_wrap:
            format_dict['text_wrap'] = True
            
        return format_dict

class StyleManager:
    """Manages Excel styles with efficient caching"""
    
    def __init__(self, workbook: xlsxwriter.Workbook):
        self.workbook = workbook
        self._format_cache = {}
        self._style_cache = {}
        self._color_cache = {}
        self._max_cache_size = 2500
        self._cache_hits = 0
        self._cache_misses = 0
        
    def _clean_cache(self):
        """Clean up cache using LRU strategy when exceeds max size"""
        if len(self._format_cache) > self._max_cache_size:
            items = sorted(self._format_cache.items(), key=lambda x: x[1]['last_used'])
            self._format_cache = dict(items[-self._max_cache_size:])
            
        if len(self._style_cache) > self._max_cache_size:
            items = sorted(self._style_cache.items(), key=lambda x: x[1].get('last_used', 0))
            self._style_cache = dict(items[-self._max_cache_size:])
            
        if len(self._color_cache) > self._max_cache_size:
            items = sorted(self._color_cache.items(), key=lambda x: x[1].get('last_used', 0))
            self._color_cache = dict(items[-self._max_cache_size:])
            
    @lru_cache(maxsize=2500)
    def _parse_color(self, color: str) -> Optional[str]:
        """Parse and normalize color values with caching"""
        if not color:
            return None
            
        if color in self._color_cache:
            self._cache_hits += 1
            self._color_cache[color]['last_used'] = time.time()
            return self._color_cache[color]['value']
            
        self._cache_misses += 1
        try:
            if color.startswith('rgb'):
                r, g, b = map(int, re.findall(r'\d+', color))
                hex_color = f'#{r:02x}{g:02x}{b:02x}'.upper()
            elif color.startswith('#'):
                hex_color = color.upper()
            else:
                hex_color = color.upper()
                
            if len(self._color_cache) < self._max_cache_size:
                self._color_cache[color] = {
                    'value': hex_color,
                    'last_used': time.time()
                }
            return hex_color
        except:
            return None
            
    @lru_cache(maxsize=2500)
    def _parse_css_style(self, style: str) -> Dict:
        """Parse CSS style string with caching"""
        if not style:
            return {}
            
        if style in self._style_cache:
            self._cache_hits += 1
            self._style_cache[style]['last_used'] = time.time()
            return self._style_cache[style]['value']
            
        self._cache_misses += 1
        style_dict = {}
        for item in style.split(';'):
            if ':' in item:
                prop, value = item.split(':', 1)
                style_dict[prop.strip()] = value.strip()
                
        if len(self._style_cache) < self._max_cache_size:
            self._style_cache[style] = {
                'value': style_dict,
                'last_used': time.time()
            }
        return style_dict

    def get_cell_style(self, node) -> CellStyle:
        """Extract cell style from HTML node with optimized caching"""
        style = CellStyle()
        
        # Get inline styles
        css = self._parse_css_style(node.attributes.get('style', ''))
        
        # Check for nested table styles
        nested_table = node.css_first('table')
        if nested_table:
            nested_css = self._parse_css_style(nested_table.attributes.get('style', ''))
            # Merge nested table styles with cell styles
            css.update(nested_css)
        
        # Process font properties
        if 'font-family' in css:
            style.font_name = css['font-family'].strip('"\'')
        if 'font-size' in css:
            size = css['font-size']
            if 'px' in size:
                style.font_size = float(size.replace('px', '')) * 0.75  # Convert px to points
        if css.get('font-weight', '') in ('bold', '700', '800', '900'):
            style.bold = True
        if css.get('font-style', '') == 'italic':
            style.italic = True
        if css.get('text-decoration', '') == 'underline':
            style.underline = True
            
        # Process colors with caching
        style.font_color = self._parse_color(css.get('color'))
        style.bg_color = self._parse_color(css.get('background-color'))
        
        # Process alignment
        if 'text-align' in css:
            align = css['text-align']
            if align == 'left':
                style.align = 'left'
            elif align == 'right':
                style.align = 'right'
            else:
                style.align = 'center'
                
        if 'vertical-align' in css:
            valign = css['vertical-align']
            if valign == 'top':
                style.valign = 'top'
            elif valign == 'bottom':
                style.valign = 'bottom'
            else:
                style.valign = 'vcenter'
                
        # Process borders - only apply when specified in style
        border_styles = {
            'border-top': 'border_top',
            'border-right': 'border_right',
            'border-bottom': 'border_bottom',
            'border-left': 'border_left',
            'border': None  # Handle full border
        }
        
        # Check for full border first
        if 'border' in css:
            border_value = css['border']
            border_type = 0
            if 'solid' in border_value:
                border_type = 1
            elif 'double' in border_value:
                border_type = 2
            elif 'dashed' in border_value:
                border_type = 3
                
            if border_type > 0:
                style.border_top = border_type
                style.border_right = border_type
                style.border_bottom = border_type
                style.border_left = border_type
        
        # Then process individual borders (these will override full border)
        for css_prop, style_prop in border_styles.items():
            if style_prop and css_prop in css:  # Skip 'border' as it's already handled
                value = css[css_prop]
                if 'solid' in value:
                    setattr(style, style_prop, 1)
                elif 'double' in value:
                    setattr(style, style_prop, 2)
                elif 'dashed' in value:
                    setattr(style, style_prop, 3)
                    
        # Set text wrap by default for all cells
        style.text_wrap = True
                    
        return style
        
    def get_format(self, style: CellStyle) -> object:
        """Get cached format or create new one with improved caching"""
        format_key = hash(tuple(sorted(style.to_excel_format().items())))
        
        if format_key in self._format_cache:
            self._cache_hits += 1
            self._format_cache[format_key]['last_used'] = time.time()
            return self._format_cache[format_key]['format']
            
        self._cache_misses += 1
        excel_format = self.workbook.add_format(style.to_excel_format())
        
        if len(self._format_cache) < self._max_cache_size:
            self._format_cache[format_key] = {
                'format': excel_format,
                'last_used': time.time()
            }
        
        if len(self._format_cache) >= self._max_cache_size * 0.9:  # Clean at 90% capacity
            self._clean_cache()
            
        return excel_format

    def get_cache_stats(self) -> Dict:
        """Get cache performance statistics"""
        total_requests = self._cache_hits + self._cache_misses
        hit_rate = (self._cache_hits / total_requests * 100) if total_requests > 0 else 0
        
        return {
            'cache_hits': self._cache_hits,
            'cache_misses': self._cache_misses,
            'hit_rate': f"{hit_rate:.2f}%",
            'format_cache_size': len(self._format_cache),
            'style_cache_size': len(self._style_cache),
            'color_cache_size': len(self._color_cache)
        }

class TableMatrix:
    """Manages table cell matrix with merge handling and nested table support"""
    
    def __init__(self, rows: int, cols: int):
        self.rows = rows
        self.cols = cols
        self.matrix = [[None] * cols for _ in range(rows)]
        self.merged_cells = {}  # Maps (row,col) to origin cell
        self.cell_origins = {}  # Maps origin position to cell data
        self.occupied_positions = set()  # Track all occupied positions
        self.header_groups = {}  # Track header group relationships
        self.nested_tables = {}  # Track nested table positions and info
        self.conflict_resolution_matrix = [[set() for _ in range(cols)] for _ in range(rows)]
        
    def is_position_available(self, row: int, col: int) -> bool:
        """Check if position is available for cell placement"""
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return False
        return (row, col) not in self.occupied_positions
        
    def find_next_position(self, start_row: int, start_col: int) -> Tuple[int, int]:
        """Find next available position starting from given coordinates"""
        col = start_col
        while col < self.cols:
            if self.is_position_available(start_row, col):
                return start_row, col
            col += 1
        return start_row, col
        
    def resolve_merge_conflicts(self, row: int, col: int, rowspan: int, colspan: int) -> bool:
        """Resolve conflicts for merged cells"""
        # Check for existing merges in the target area
        for r in range(row, row + rowspan):
            for c in range(col, col + colspan):
                if not (0 <= r < self.rows and 0 <= c < self.cols):
                    return False
                if (r, c) in self.occupied_positions:
                    existing_origin = self.merged_cells.get((r, c))
                    if existing_origin:
                        # If there's a conflict, check if we can adjust the current merge
                        if self._can_adjust_merge(row, col, rowspan, colspan, existing_origin):
                            return True
                        return False
        return True
        
    def _can_adjust_merge(self, row: int, col: int, rowspan: int, colspan: int, existing_origin: Tuple[int, int]) -> bool:
        """Check if merge can be adjusted to avoid conflicts"""
        existing_cell = self.cell_origins.get(existing_origin)
        if not existing_cell:
            return False
            
        # Try to find alternative merge arrangements
        alternatives = [
            (row, col, rowspan - 1, colspan),
            (row, col, rowspan, colspan - 1),
            (row + 1, col, rowspan - 1, colspan),
            (row, col + 1, rowspan, colspan - 1)
        ]
        
        for alt_row, alt_col, alt_rowspan, alt_colspan in alternatives:
            if self._is_valid_merge(alt_row, alt_col, alt_rowspan, alt_colspan):
                return True
                
        return False
        
    def _is_valid_merge(self, row: int, col: int, rowspan: int, colspan: int) -> bool:
        """Check if a merge range is valid"""
        if rowspan < 1 or colspan < 1:
            return False
            
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return False
            
        if row + rowspan > self.rows or col + colspan > self.cols:
            return False
            
        return True
        
    def place_cell(self, row: int, col: int, cell_data: dict) -> bool:
        """Place cell in matrix and handle merging"""
        rowspan = cell_data.get('rowspan', 1)
        colspan = cell_data.get('colspan', 1)
        
        # Handle nested tables
        nested_table = cell_data.get('nested_table')
        if nested_table:
            self.nested_tables[(row, col)] = {
                'table': nested_table,
                'rowspan': rowspan,
                'colspan': colspan,
                'parent_cell': cell_data
            }
            
        # Validate placement and resolve conflicts
        if not self.resolve_merge_conflicts(row, col, rowspan, colspan):
            return False
            
        # Place cell
        origin = (row, col)
        cell_data['origin'] = origin
        self.cell_origins[origin] = cell_data
        
        # Mark all covered positions and track merged cells
        for r in range(row, row + rowspan):
            for c in range(col, col + colspan):
                self.matrix[r][c] = cell_data
                self.occupied_positions.add((r, c))
                if (r, c) != origin:
                    self.merged_cells[(r, c)] = origin
                    
        # Special handling for header groups
        if cell_data.get('is_header'):
            if colspan > 1:  # Header spans multiple columns
                self.header_groups[origin] = {
                    'type': 'group',
                    'start_col': col,
                    'end_col': col + colspan - 1,
                    'subheader_row': row + 1
                }
            elif rowspan > 1:  # Header spans multiple rows
                self.header_groups[origin] = {
                    'type': 'spanning',
                    'start_row': row,
                    'end_row': row + rowspan - 1,
                    'col': col
                }
                
        return True
        
    def get_cell_at(self, row: int, col: int) -> Optional[dict]:
        """Get cell data at given position"""
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return None
            
        cell_data = self.matrix[row][col]
        if cell_data is None:
            return None
            
        # Check for nested table
        nested_info = self.nested_tables.get((row, col))
        if nested_info:
            return {
                **cell_data,
                'nested_table_info': nested_info
            }
            
        # Handle header styling
        if cell_data.get('is_header'):
            origin = cell_data['origin']
            if origin in self.header_groups:
                header_info = self.header_groups[origin]
                if header_info['type'] == 'group' and row == header_info['subheader_row']:
                    # This is a subheader under a header group
                    cell_data = cell_data.copy()
                    cell_data['style'] = self._adjust_subheader_style(cell_data['style'])
                elif header_info['type'] == 'spanning':
                    # This is a row-spanning header
                    cell_data = cell_data.copy()
                    cell_data['style'] = self._adjust_spanning_header_style(cell_data['style'])
                    
        return cell_data
        
    def _adjust_subheader_style(self, style: CellStyle) -> CellStyle:
        """Adjust style for subheaders"""
        new_style = CellStyle()
        for key, value in vars(style).items():
            setattr(new_style, key, value)
        new_style.font_size = (style.font_size or 10) * 0.9  # Slightly smaller font
        return new_style
        
    def _adjust_spanning_header_style(self, style: CellStyle) -> CellStyle:
        """Adjust style for row-spanning headers"""
        new_style = CellStyle()
        for key, value in vars(style).items():
            setattr(new_style, key, value)
        new_style.align = 'center'  # Center align
        new_style.valign = 'vcenter'  # Vertical center
        return new_style
        
    def get_nested_tables(self) -> Dict[Tuple[int, int], Dict]:
        """Get all nested tables with their positions"""
        return self.nested_tables.copy()
        
    def get_merge_ranges(self) -> List[Tuple[int, int, int, int]]:
        """Get all merge ranges as (start_row, start_col, end_row, end_col)"""
        merge_ranges = []
        processed = set()
        
        for row in range(self.rows):
            for col in range(self.cols):
                if (row, col) in processed:
                    continue
                    
                cell_data = self.matrix[row][col]
                if cell_data and cell_data['origin'] == (row, col):
                    rowspan = cell_data.get('rowspan', 1)
                    colspan = cell_data.get('colspan', 1)
                    
                    if rowspan > 1 or colspan > 1:
                        merge_ranges.append((
                            row, col,
                            row + rowspan - 1,
                            col + colspan - 1
                        ))
                        
                    # Mark all cells in this range as processed
                    for r in range(row, row + rowspan):
                        for c in range(col, col + colspan):
                            processed.add((r, c))
                            
        return merge_ranges

class HTMLTableConverter:
    """Converts HTML tables to Excel with high performance"""
    
    DEFAULT_OPTIONS = {
        'chunk_size': 2000,  # Increased chunk size for better performance
        'font': {
            'name': 'TH Sarabun New',
            'size': 10
        },
        'header': {
            'bold': True,
            'bg_color': '#A6A6A6',
            'border': 1
        },
        'cell': {
            'border': 1,
            'valign': 'vcenter'
        },
        'table': {
            'first_row_as_header': True,
            'auto_width': True,
            'min_width': 8,
            'max_width': 100,
            'row_height_multiplier': 1.2,
            'max_nested_level': 3  # Limit nested table processing
        },
        'html': {
            'preserve_formatting': True,
            'parse_entities': True,
            'max_cell_length': 32767  # Excel cell limit
        }
    }
    
    def __init__(self, options: Dict = None):
        self.options = self.DEFAULT_OPTIONS.copy()
        if options:
            self._deep_update(self.options, options)
            
        self._reset_state()
        
    def _deep_update(self, d: Dict, u: Dict):
        """Deep update dictionary"""
        for k, v in u.items():
            if isinstance(v, dict) and k in d:
                self._deep_update(d[k], v)
            else:
                d[k] = v
                
    def _reset_state(self):
        """Reset internal state"""
        self._column_widths = {}
        self._row_heights = {}
        self._current_row = 0
        self._nested_level = 0
        self._processed_tables = set()
        gc.collect()
        
    def _calculate_column_width(self, content: str, style: CellStyle) -> float:
        """Calculate optimal column width based on content and style"""
        if not content:
            return self.options['table']['min_width']
            
        # Split into lines
        lines = str(content).split('\n')
        max_line_length = max(len(line.strip()) for line in lines)
        
        # Base width calculation
        base_width = max_line_length * 1.0  # Reduced from 1.2 to 1.0
        
        # Adjust for font size
        if style.font_size:
            size_factor = style.font_size / 11  # Adjusted base font size
            base_width *= size_factor
            
        # Adjust for font style
        if style.bold:
            base_width *= 1.05  # Reduced from 1.1 to 1.05
            
        # Add minimal padding for borders
        if any([style.border_left, style.border_right]):
            base_width += 0.5  # Reduced from 1 to 0.5
            
        # Add minimal space for alignment
        if style.align in ('center', 'right'):
            base_width += 0.5  # Reduced from 1 to 0.5
            
        # Apply min/max constraints
        return min(
            max(base_width, self.options['table']['min_width']),
            self.options['table']['max_width']
        )
        
    def _get_cell_content(self, node) -> str:
        """Extract and clean cell content"""
        # Get text content first
        content = node.text()
        
        # Handle <br> tags by replacing them with newlines
        if node.css('br'):
            # Get HTML content and replace br tags with newlines
            content = node.html.replace('<br>', '\n').replace('<br/>', '\n')
            # Remove any remaining HTML tags
            content = re.sub('<[^<]+?>', '', content)
            
        # Clean up whitespace
        content = ' '.join(content.split()) if content else ''
        
        # Unescape HTML entities
        if self.options['html']['parse_entities']:
            content = html.unescape(content)
            
        # Truncate if exceeds Excel limit
        if len(content) > self.options['html']['max_cell_length']:
            content = content[:self.options['html']['max_cell_length']]
            
        return content
        
    def _process_table(self, table_node, worksheet: object, style_manager: StyleManager):
        """Process table node and convert to Excel with nested table support"""
        if self._nested_level >= self.options['table']['max_nested_level']:
            logger.warning(f"Skipping nested table at level {self._nested_level} (max level reached)")
            return
            
        self._nested_level += 1
        try:
            # Skip if already processed
            if table_node in self._processed_tables:
                return
                
            # Calculate dimensions
            rows = table_node.css('tr')
            max_rows = len(rows)
            max_cols = self._calculate_max_columns(rows)
            
            # Initialize matrix
            matrix = TableMatrix(max_rows, max_cols)
            current_row = self._current_row
            
            # Process all rows
            row_idx = 0
            for tr in rows:
                if 'display: none' in tr.attributes.get('style', ''):
                    continue
                    
                col_idx = 0
                for cell in tr.css('td, th'):
                    # Find next available position
                    while col_idx < matrix.cols and not matrix.is_position_available(row_idx, col_idx):
                        col_idx += 1
                        
                    if col_idx >= matrix.cols:
                        break
                        
                    # Get cell properties
                    rowspan = int(cell.attributes.get('rowspan', 1))
                    colspan = int(cell.attributes.get('colspan', 1))
                    
                    # Create cell data
                    cell_data = {
                        'content': self._get_cell_content(cell),
                        'style': style_manager.get_cell_style(cell),
                        'rowspan': rowspan,
                        'colspan': colspan,
                        'is_header': cell.tag == 'th',
                        'row': row_idx,
                        'col': col_idx
                    }
                    
                    # Place cell
                    if not matrix.place_cell(row_idx, col_idx, cell_data):
                        logger.warning(f"Failed to place cell at ({row_idx}, {col_idx})")
                        continue
                        
                    col_idx += colspan
                    
                row_idx += 1
                
            # Write to Excel
            self._write_to_excel(matrix, worksheet, max_rows, max_cols, current_row, style_manager)
            
            # Update current row
            self._current_row = current_row + max_rows
            
            self._processed_tables.add(table_node)
            
        finally:
            self._nested_level -= 1
            
    def _write_to_excel(self, matrix: TableMatrix, worksheet: object, max_rows: int, max_cols: int, 
                       current_row: int, style_manager: StyleManager):
        """Write matrix data to Excel worksheet"""
        try:
            # Write cells to Excel
            for row in range(max_rows):
                for col in range(max_cols):
                    cell_data = matrix.get_cell_at(row, col)
                    if cell_data is None or cell_data['origin'] != (row, col):
                        continue
                        
                    try:
                        cell_style = style_manager.get_format(cell_data['style'])
                        content = cell_data['content']
                        
                        if cell_data['rowspan'] > 1 or cell_data['colspan'] > 1:
                            # Check if merge range is already used
                            merge_range = (
                                current_row + row,
                                col,
                                current_row + row + cell_data['rowspan'] - 1,
                                col + cell_data['colspan'] - 1
                            )
                            
                            try:
                                worksheet.merge_range(
                                    *merge_range,
                                    content,
                                    cell_style
                                )
                            except Exception as e:
                                # If merge fails, write as normal cell
                                worksheet.write(
                                    current_row + row,
                                    col,
                                    content,
                                    cell_style
                                )
                        else:
                            worksheet.write(
                                current_row + row,
                                col,
                                content,
                                cell_style
                            )
                            
                        # Update column widths
                        if self.options['table']['auto_width']:
                            content_width = self._calculate_column_width(
                                content,
                                cell_data['style']
                            )
                            width_per_col = content_width / cell_data['colspan']
                            for c in range(col, col + cell_data['colspan']):
                                self._column_widths[c] = max(
                                    self._column_widths.get(c, 0),
                                    width_per_col
                                )
                                
                    except Exception as e:
                        logger.warning(f"Failed to write cell at ({row}, {col}): {str(e)}")
                        continue
                        
                # Set row height
                row_height = 18 * self.options['table']['row_height_multiplier']
                worksheet.set_row(current_row + row, row_height)
                
            # Set final column widths
            if self.options['table']['auto_width']:
                for col, width in self._column_widths.items():
                    adjusted_width = min(
                        max(width * 0.85, self.options['table']['min_width']),
                        self.options['table']['max_width']
                    )
                    worksheet.set_column(col, col, adjusted_width)
                    
        except Exception as e:
            logger.error(f"Error writing to Excel: {str(e)}")
            raise
            
    def convert(self, html_content: str, output_path: str = None) -> Dict:
        """Convert HTML to Excel, returns buffer if no output_path provided"""
        try:
            with PerformanceTimer('Total conversion'):
                # Parse HTML
                parser = HTMLParser(html_content)
                
                # Create workbook - either to file or buffer
                workbook_options = {
                    'constant_memory': True,
                    'default_format_properties': {
                        'font_name': self.options['font']['name'],
                        'font_size': self.options['font']['size']
                    }
                }
                
                if output_path:
                    workbook = xlsxwriter.Workbook(output_path, workbook_options)
                else:
                    output_buffer = BytesIO()
                    workbook_options['in_memory'] = True
                    workbook = xlsxwriter.Workbook(output_buffer, workbook_options)
                
                worksheet = workbook.add_worksheet()
                style_manager = StyleManager(workbook)
                
                # Process tables sequentially with spacing
                current_row = 0
                tables = parser.css('table')
                
                for table in tables:
                    if 'display: none' not in table.attributes.get('style', ''):
                        # Save current row position
                        self._current_row = current_row
                        
                        # Process current table
                        self._process_table(table, worksheet, style_manager)
                        
                        # Update current row for next table (add spacing)
                        current_row = self._current_row + 2  # Add 2 rows spacing between tables
                        
                workbook.close()
                
                # Print performance summary
                PerformanceTimer.print_summary()
                
                if output_path:
                    return {"success": True, "data": ""}
                else:
                    # Get the buffer value
                    excel_data = output_buffer.getvalue()
                    output_buffer.close()
                    return {"success": True, "data": excel_data}
                
        except Exception as e:
            logger.error(f"Conversion failed: {str(e)}")
            return {"success": False, "error": str(e)}
            
        finally:
            self._reset_state()

    def _calculate_max_columns(self, rows) -> int:
        """Calculate maximum number of columns needed"""
        max_cols = 0
        for tr in rows:
            if 'display: none' in tr.attributes.get('style', ''):
                continue
            col_sum = 0
            for cell in tr.css('td, th'):
                colspan = int(cell.attributes.get('colspan', 1))
                col_sum += colspan
            max_cols = max(max_cols, col_sum)
        return max_cols

    def _extract_nested_tables(self, node) -> List[Tuple[object, Dict]]:
        """Extract nested tables with their context"""
        nested_tables = []
        
        for cell in node.css('td, th'):
            tables = cell.css('table')
            if tables:
                # Get cell position info
                parent_row = len(node.css_first('tr').find_previous_siblings('tr'))
                parent_col = len(cell.find_previous_siblings(['td', 'th']))
                
                for table in tables:
                    if table not in self._processed_tables:
                        nested_tables.append((table, {
                            'parent_cell': cell,
                            'parent_row': parent_row,
                            'parent_col': parent_col,
                            'rowspan': int(cell.attributes.get('rowspan', 1)),
                            'colspan': int(cell.attributes.get('colspan', 1))
                        }))
                        self._processed_tables.add(table)
                        
        return nested_tables
        
    def _get_cell_content(self, node) -> str:
        """Extract and clean cell content"""
        # Get text content first
        content = node.text()
        
        # Handle <br> tags by replacing them with newlines
        if node.css('br'):
            # Get HTML content and replace br tags with newlines
            content = node.html.replace('<br>', '\n').replace('<br/>', '\n')
            # Remove any remaining HTML tags
            content = re.sub('<[^<]+?>', '', content)
            
        # Clean up whitespace
        content = ' '.join(content.split()) if content else ''
        
        # Unescape HTML entities
        if self.options['html']['parse_entities']:
            content = html.unescape(content)
            
        # Truncate if exceeds Excel limit
        if len(content) > self.options['html']['max_cell_length']:
            content = content[:self.options['html']['max_cell_length']]
            
        return content

if __name__ == "__main__":
    try:
        # Read input from stdin
        input_data = sys.stdin.read().strip()
        converter = HTMLTableConverter()
        
        try:
            data = json.loads(input_data)
            html_content = data.get('html', '')
            output_to_buffer = data.get('buffer', True) 
            output_file = data.get('output', None) if not output_to_buffer else None
        except json.JSONDecodeError:
            html_content = input_data
            output_file = None
            
        if not html_content:
            print(json.dumps({"error": "No HTML content provided"}), file=sys.stderr)
            sys.exit(1)
            
        result = converter.convert(html_content, output_file)
        
        # If result contains binary data, encode it to base64
        if result["success"] and isinstance(result.get("data"), bytes):
            import base64
            result["data"] = base64.b64encode(result["data"]).decode('utf-8')
            
        print(json.dumps(result))
        sys.exit(0 if result["success"] else 1)
        
    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(1) 