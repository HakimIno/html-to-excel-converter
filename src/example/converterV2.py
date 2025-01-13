import sys
import json
import xlsxwriter
import re
import html
from bs4 import BeautifulSoup, SoupStrainer, Tag
from functools import lru_cache
from typing import Dict, List, Set, Tuple, Union
import gc
import time
import logging

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

class ExcelStyleManager:
    """Manages Excel styles and formatting with efficient caching"""
    
    def __init__(self):
        self._format_cache = {}
        self._style_hash_cache = {}
        self._color_cache = {}
        self._max_cache_size = 2500  # Increased from 2000
        self._cache_hits = 0
        self._cache_misses = 0
        
    def _clean_cache(self):
        """Clean up cache using LRU strategy when exceeds max size"""
        if len(self._format_cache) > self._max_cache_size:
            # Keep only the most recently used items
            items = sorted(self._format_cache.items(), key=lambda x: x[1]['last_used'])
            self._format_cache = dict(items[-self._max_cache_size:])
            
        # Use LRU for style and color caches too instead of clearing
        if len(self._style_hash_cache) > self._max_cache_size:
            items = sorted(self._style_hash_cache.items(), key=lambda x: x[1].get('last_used', 0))
            self._style_hash_cache = dict(items[-self._max_cache_size:])
            
        if len(self._color_cache) > self._max_cache_size:
            items = sorted(self._color_cache.items(), key=lambda x: x[1].get('last_used', 0))
            self._color_cache = dict(items[-self._max_cache_size:])

    @lru_cache(maxsize=2500)  # Increased from 2000
    def _normalize_color(self, color: str) -> str:
        """Normalize color values with caching"""
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

    @lru_cache(maxsize=5000)  # Increased from 2000
    def _parse_style(self, style: str) -> Dict:
        """Parse CSS style string with caching"""
        if not style:
            return {}
            
        if style in self._style_hash_cache:
            self._cache_hits += 1
            self._style_hash_cache[style]['last_used'] = time.time()
            return self._style_hash_cache[style]['value']
            
        self._cache_misses += 1
        style_dict = {}
        for item in style.split(';'):
            if ':' in item:
                prop, value = item.split(':', 1)
                style_dict[prop.strip()] = value.strip()
        
        if len(self._style_hash_cache) < self._max_cache_size:
            self._style_hash_cache[style] = {
                'value': style_dict,
                'last_used': time.time()
            }
        return style_dict

    def get_format(self, workbook, properties: Dict) -> object:
        """Get cached format or create new one with improved caching"""
        format_key = hash(frozenset(properties.items()))
        
        if format_key in self._format_cache:
            self._cache_hits += 1
            self._format_cache[format_key]['last_used'] = time.time()
            return self._format_cache[format_key]['format']
            
        self._cache_misses += 1
        excel_format = workbook.add_format(properties)
        
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
            'style_cache_size': len(self._style_hash_cache),
            'color_cache_size': len(self._color_cache)
        }

class HTMLTableConverter:
    """Converts HTML tables to Excel with high performance and memory efficiency"""
    
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
            
        self.style_manager = ExcelStyleManager()
        self._reset_state()
        self._setup_parsers()

    def _setup_parsers(self):
        """Setup optimized parsers"""
        self._cell_strainer = SoupStrainer(['td', 'th'])
        self._table_strainer = SoupStrainer('table')
        self._row_strainer = SoupStrainer('tr')

    def _reset_state(self):
        """Reset internal state with optimized data structures"""
        self._column_widths = {}
        self._row_heights = {}
        self._merged_ranges = set()
        self._current_row = 0
        self._max_col = 0
        self._chunk_buffer = []
        self._processed_tables = set()
        gc.collect()

    def _deep_update(self, d: Dict, u: Dict):
        """Deep update dictionary"""
        for k, v in u.items():
            if isinstance(v, dict) and k in d:
                self._deep_update(d[k], v)
            else:
                d[k] = v

    def _process_cell_content(self, cell: BeautifulSoup) -> Tuple[str, Dict]:
        """Optimized cell content processing"""
        properties = {
            'font_name': self.options['font']['name'],
            'font_size': self.options['font']['size'],
            'align': 'center'  # Default alignment
        }
        
        # Fast style processing
        style = cell.get('style', '')
        if style:
            inline_style = self.style_manager._parse_style(style)
            
            # Process alignment
            if 'text-align' in inline_style:
                properties['align'] = inline_style['text-align']
            
            # Process colors efficiently
            bg_color = inline_style.get('background-color')
            if bg_color:
                properties['bg_color'] = self.style_manager._normalize_color(bg_color)
                
            font_color = inline_style.get('color')
            if font_color:
                properties['font_color'] = self.style_manager._normalize_color(font_color)
            
            # Process font styles
            if any(s in inline_style.get('font-weight', '') for s in ['bold', '700', '800', '900']):
                properties['bold'] = True
            if 'italic' in inline_style.get('font-style', ''):
                properties['italic'] = True
                
            # Process borders
            if 'border' in inline_style:
                self._process_borders(inline_style, properties)
        
        # Efficient text extraction
        if self.options['html']['preserve_formatting']:
            content = []
            for element in cell.children:
                if isinstance(element, str):
                    content.append(element.strip())
                elif element.name in ['br', 'p']:
                    content.append('\n')
                else:
                    content.append(element.get_text(strip=True))
            text = ' '.join(filter(None, content))
        else:
            text = cell.get_text(strip=True)
            
        if self.options['html']['parse_entities']:
            text = html.unescape(text)
            
        # Truncate if exceeds Excel limit
        if len(text) > self.options['html']['max_cell_length']:
            text = text[:self.options['html']['max_cell_length']]
            
        return text, properties

    def _process_borders(self, style: Dict, properties: Dict):
        """Process border styles efficiently"""
        borders = {
            'border-top': 'top',
            'border-right': 'right',
            'border-bottom': 'bottom',
            'border-left': 'left'
        }
        
        for css_prop, excel_border in borders.items():
            if css_prop in style:
                value = style[css_prop]
                if 'solid' in value:
                    properties[f'border_{excel_border}'] = 1
                elif 'double' in value:
                    properties[f'border_{excel_border}'] = 2
                elif 'dashed' in value:
                    properties[f'border_{excel_border}'] = 3

    def _calculate_column_width(self, content: str, cell_style: Dict) -> float:
        """Calculate optimal column width based on content and style"""
        # Get width from style if specified
        if 'width' in cell_style:
            width_str = cell_style['width']
            if '%' in width_str:
                # Convert percentage to approximate characters
                width_percent = float(width_str.replace('%', ''))
                return (width_percent / 100) * 100  # Max width is 100 characters
            elif 'px' in width_str:
                # Convert pixels to approximate characters
                width_px = float(width_str.replace('px', ''))
                return width_px / 7  # Approximate 7 pixels per character
                
        # Calculate based on content
        if not content:
            return self.options['table']['min_width']
            
        lines = str(content).split('\n')
        max_line_length = max(len(line) for line in lines)
        
        # Adjust for font size
        font_size = float(cell_style.get('font-size', '').replace('px', '').strip() or self.options['font']['size'])
        size_factor = font_size / 10  # Base size is 10px
        
        # Calculate base width
        base_width = max_line_length * size_factor
        
        # Adjust for padding
        padding_left = float(cell_style.get('padding-left', '').replace('px', '').strip() or 0)
        padding_right = float(cell_style.get('padding-right', '').replace('px', '').strip() or 0)
        total_padding = (padding_left + padding_right) / 7  # Convert pixels to characters
        
        # Final width calculation
        width = base_width + total_padding
        
        # Apply min/max constraints
        return min(
            max(width, self.options['table']['min_width']),
            self.options['table']['max_width']
        )

    def _process_table_chunk(self, chunk: List, worksheet: object, workbook: object):
        """Process a chunk of table rows efficiently"""
        # Track merged ranges to prevent overlaps
        merged_ranges = set()
        
        for row_data in chunk:
            row_num, cells = row_data
            col_num = 0
            
            for content, properties, span in cells:
                # Skip columns that are part of previous merges
                while (row_num, col_num) in self._merged_ranges:
                    col_num += 1
                    
                if span:
                    rowspan, colspan = span
                    end_row = row_num + rowspan - 1
                    end_col = col_num + colspan - 1
                    
                    # Check if this merge would overlap with any existing merge
                    can_merge = True
                    merge_cells = set()
                    
                    for r in range(row_num, end_row + 1):
                        for c in range(col_num, end_col + 1):
                            if (r, c) in self._merged_ranges or (r, c) in merged_ranges:
                                can_merge = False
                                break
                            merge_cells.add((r, c))
                        if not can_merge:
                            break
                            
                    if can_merge:
                        # Perform the merge
                        worksheet.merge_range(
                            row_num, col_num,
                            end_row, end_col,
                            content,
                            self.style_manager.get_format(workbook, properties)
                        )
                        # Track merged cells
                        merged_ranges.update(merge_cells)
                        self._merged_ranges.update(merge_cells)
                        
                        # Update column width for merged cells
                        if self.options['table']['auto_width']:
                            # Calculate width per column in merge
                            cell_style = self.style_manager._parse_style(properties.get('style', ''))
                            total_width = self._calculate_column_width(content, cell_style)
                            width_per_col = total_width / colspan
                            for c in range(col_num, end_col + 1):
                                self._column_widths[c] = max(
                                    self._column_widths.get(c, 0),
                                    width_per_col
                                )
                                
                        col_num = end_col + 1
                    else:
                        # If can't merge, write as normal cell
                        worksheet.write(
                            row_num, col_num,
                            content,
                            self.style_manager.get_format(workbook, properties)
                        )
                        # Update column width
                        if self.options['table']['auto_width']:
                            cell_style = self.style_manager._parse_style(properties.get('style', ''))
                            width = self._calculate_column_width(content, cell_style)
                            self._column_widths[col_num] = max(
                                self._column_widths.get(col_num, 0),
                                width
                            )
                        col_num += 1
                else:
                    # Write normal cell
                    worksheet.write(
                        row_num, col_num,
                        content,
                        self.style_manager.get_format(workbook, properties)
                    )
                    # Update column width
                    if self.options['table']['auto_width']:
                        cell_style = self.style_manager._parse_style(properties.get('style', ''))
                        width = self._calculate_column_width(content, cell_style)
                        self._column_widths[col_num] = max(
                            self._column_widths.get(col_num, 0),
                            width
                        )
                    col_num += 1

    def _is_merge_conflict(self, start_row: int, start_col: int, end_row: int, end_col: int) -> bool:
        """Check if a merge range conflicts with existing merges"""
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                if (r, c) in self._merged_ranges:
                    return True
        return False

    def resolve_merge_conflicts(self, table: BeautifulSoup) -> List[List[Dict]]:
        """Resolve merge conflicts in table by creating a matrix representation"""
        # 1. Calculate table dimensions
        max_rows = 0
        max_cols = 0
        for row in table.find_all('tr'):
            if 'display: none' not in row.get('style', ''):
                max_cols = max(max_cols, sum(int(cell.get('colspan', 1)) 
                    for cell in row.find_all(['td', 'th'])))
                max_rows += 1
                
        # 2. Create empty matrix
        matrix = [[None] * max_cols for _ in range(max_rows)]
        
        # 3. Fill matrix and resolve conflicts
        row_idx = 0
        for tr in table.find_all('tr'):
            if 'display: none' in tr.get('style', ''):
                continue
                
            col_idx = 0
            for cell in tr.find_all(['td', 'th']):
                # Skip occupied positions
                while col_idx < max_cols and matrix[row_idx][col_idx] is not None:
                    col_idx += 1
                    
                if col_idx >= max_cols:
                    break
                    
                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))
                
                # Check for overlaps and adjust spans
                actual_rowspan = rowspan
                actual_colspan = colspan
                
                for r in range(row_idx, min(row_idx + rowspan, max_rows)):
                    for c in range(col_idx, min(col_idx + colspan, max_cols)):
                        if matrix[r][c] is not None:
                            actual_rowspan = min(actual_rowspan, r - row_idx)
                            actual_colspan = min(actual_colspan, c - col_idx)
                            break
                    if actual_rowspan != rowspan or actual_colspan != colspan:
                        break
                
                # Ensure minimum spans
                actual_rowspan = max(1, actual_rowspan)
                actual_colspan = max(1, actual_colspan)
                
                # Fill matrix with cell info
                cell_info = {
                    'element': cell,
                    'rowspan': actual_rowspan,
                    'colspan': actual_colspan,
                    'is_header': cell.name == 'th'
                }
                
                for r in range(row_idx, row_idx + actual_rowspan):
                    for c in range(col_idx, col_idx + actual_colspan):
                        if r < max_rows and c < max_cols:
                            matrix[r][c] = cell_info
                            
                col_idx += actual_colspan
            row_idx += 1
            
        return matrix

    def _process_table(self, table: BeautifulSoup, worksheet: object, workbook: object):
        """Process table with optimized nested table handling"""
        self._chunk_buffer = []
        current_row = self._current_row
        
        # Process nested tables efficiently
        nested_tables = self._extract_nested_tables(table)
        
        # Process main table
        resolved_matrix = self.resolve_merge_conflicts(table)
        max_row_used = self._process_matrix(resolved_matrix, current_row, worksheet, workbook)
        
        # Process nested tables with level limit
        if nested_tables and len(nested_tables) <= self.options['table']['max_nested_level']:
            max_row_used = self._process_nested_tables(
                nested_tables, current_row, max_row_used, worksheet, workbook
            )
        
        # Update final position
        self._current_row = current_row + max_row_used + 1

    def _extract_nested_tables(self, table: BeautifulSoup, level: int = 0) -> List[Dict]:
        """Extract nested tables efficiently"""
        if level >= self.options['table']['max_nested_level']:
            return []
            
        nested_tables = []
        processed_positions = set()
        
        for row_idx, row in enumerate(table.find_all('tr', recursive=False)):
            if 'display: none' in row.get('style', ''):
                continue
                
            for col_idx, cell in enumerate(row.find_all(['td', 'th'], recursive=False)):
                inner_tables = cell.find_all('table', recursive=False)
                if inner_tables:
                    abs_row = row_idx
                    abs_col = sum(int(sibling.get('colspan', 1)) 
                                for sibling in cell.find_previous_siblings(['td', 'th']))
                    
                    position_key = (abs_row, abs_col)
                    if position_key not in processed_positions:
                        processed_positions.add(position_key)
                        
                        for inner_table in inner_tables:
                            if inner_table not in self._processed_tables:
                                nested_tables.append({
                                    'table': inner_table,
                                    'position': {'row': abs_row, 'col': abs_col},
                                    'parent_cell': cell,
                                    'rowspan': int(cell.get('rowspan', 1)),
                                    'colspan': int(cell.get('colspan', 1))
                                })
                                
                                # Recursively process deeper nested tables
                                deeper_tables = self._extract_nested_tables(
                                    inner_table, level + 1
                                )
                                nested_tables.extend(deeper_tables)
        
        return nested_tables

    def _process_matrix(self, matrix: List[List[Dict]], current_row: int, 
                       worksheet: object, workbook: object) -> int:
        """Process matrix efficiently"""
        max_row_used = 0
        
        for row_idx, row in enumerate(matrix):
            row_cells = []
            col_idx = 0
            
            while col_idx < len(row):
                cell_info = row[col_idx]
                
                if cell_info and cell_info['element'] is not None:
                    if self._is_cell_start(row_idx, col_idx, matrix):
                        content, properties = self._process_cell_content(cell_info['element'])
                        span = None
                        if cell_info['rowspan'] > 1 or cell_info['colspan'] > 1:
                            span = (cell_info['rowspan'], cell_info['colspan'])
                        row_cells.append((content, properties, span))
                        
                        if span:
                            max_row_used = max(max_row_used, row_idx + span[0])
                        
                    col_idx += cell_info['colspan'] if cell_info['colspan'] > 0 else 1
                else:
                    col_idx += 1
                    
            if row_cells:
                self._chunk_buffer.append((current_row + row_idx, row_cells))
                max_row_used = max(max_row_used, row_idx + 1)
                
                if len(self._chunk_buffer) >= self.options['chunk_size']:
                    self._process_table_chunk(self._chunk_buffer, worksheet, workbook)
                    self._chunk_buffer = []
                    gc.collect()
        
        if self._chunk_buffer:
            self._process_table_chunk(self._chunk_buffer, worksheet, workbook)
            self._chunk_buffer = []
            gc.collect()
            
        return max_row_used

    def _is_cell_start(self, row_idx: int, col_idx: int, matrix: List[List[Dict]]) -> bool:
        """Check if cell is starting position efficiently"""
        cell_info = matrix[row_idx][col_idx]
        
        if col_idx > 0:
            prev_cell = matrix[row_idx][col_idx - 1]
            if prev_cell and prev_cell is cell_info:
                return False
                
        if row_idx > 0:
            above_cell = matrix[row_idx - 1][col_idx]
            if above_cell and above_cell is cell_info:
                return False
                
        return True

    def _process_nested_tables(self, nested_tables: List[Dict], current_row: int,
                             max_row_used: int, worksheet: object, workbook: object) -> int:
        """Process nested tables efficiently"""
        for nested_info in sorted(nested_tables, key=lambda x: (x['position']['row'], x['position']['col'])):
            nested_table = nested_info['table']
            
            if nested_table in self._processed_tables:
                continue
                
            pos = nested_info['position']
            nested_row = current_row + pos['row']
            
            original_row = self._current_row
            self._current_row = nested_row
            self._process_table(nested_table, worksheet, workbook)
            
            max_row_used = max(max_row_used, pos['row'] + (self._current_row - nested_row))
            self._current_row = original_row
            self._processed_tables.add(nested_table)
        
        return max_row_used

    def convert(self, html_content: str, output: Union[str, object]) -> Dict:
        """Convert HTML to Excel with optimized performance"""
        workbook = None
        try:
            with PerformanceTimer('Total conversion'):
                self._reset_state()
                
                # Create workbook with optimized settings
                workbook_options = {
                    'constant_memory': True,
                    'default_format_properties': {
                        'font_name': self.options['font']['name'],
                        'font_size': self.options['font']['size']
                    }
                }
                
                if isinstance(output, str):
                    workbook = xlsxwriter.Workbook(output, workbook_options)
                else:
                    workbook_options['in_memory'] = True
                    workbook = xlsxwriter.Workbook(output, workbook_options)
                
                worksheet = workbook.add_worksheet('Sheet1')
                
                # Parse HTML efficiently
                with PerformanceTimer('HTML parsing'):
                    soup = BeautifulSoup(html_content, 'html.parser', parse_only=self._table_strainer)
                
                # Process tables in parallel for large documents
                with PerformanceTimer('Table processing'):
                    if len(html_content) > 1000000:  # 1MB threshold
                        logger.info('Using parallel processing for large document')
                        self._process_tables_parallel(soup.find_all('table', recursive=False), worksheet, workbook)
                    else:
                        self._process_tables_sequential(soup.find_all('table', recursive=False), worksheet, workbook)
                
                # Apply column widths efficiently
                with PerformanceTimer('Column width adjustment'):
                    if self.options['table']['auto_width']:
                        for col, width in sorted(self._column_widths.items()):
                            adjusted_width = min(
                                max(width, self.options['table']['min_width']),
                                self.options['table']['max_width']
                            )
                            worksheet.set_column(col, col, adjusted_width)
                
                workbook.close()
                
                # Print performance summary at the end
                PerformanceTimer.print_summary()
                
                return {"success": True}
                
        except Exception as e:
            logger.error(f'Conversion failed: {str(e)}')
            if workbook:
                try:
                    workbook.close()
                except:
                    pass
            return {"success": False, "error": str(e)}
        finally:
            self._reset_state()

    def _process_tables_sequential(self, tables: List[Tag], worksheet: object, workbook: object):
        """Process tables sequentially"""
        for table in tables:
            if 'display: none' not in table.get('style', '') and table not in self._processed_tables:
                self._process_table(table, worksheet, workbook)
                self._processed_tables.add(table)

    def _process_tables_parallel(self, tables: List[Tag], worksheet: object, workbook: object):
        """Process tables in parallel for large documents"""
        from concurrent.futures import ThreadPoolExecutor
        import threading
        
        lock = threading.Lock()
        
        def process_table_safe(table):
            if 'display: none' not in table.get('style', '') and table not in self._processed_tables:
                with lock:
                    if table not in self._processed_tables:
                        self._process_table(table, worksheet, workbook)
                        self._processed_tables.add(table)
        
        with ThreadPoolExecutor(max_workers=4) as executor:
            executor.map(process_table_safe, tables)

    @classmethod
    def convert_file(cls, input_path: str, output_path: str, options: Dict = None) -> Dict:
        """Convert HTML file to Excel file"""
        try:
            with open(input_path, 'r', encoding='utf-8') as file:
                html_content = file.read()
            
            converter = cls(options)
            return converter.convert(html_content, output_path)
            
        except Exception as e:
            return {"success": False, "error": str(e)}

if __name__ == "__main__":
    try:
        input_data = sys.stdin.read().strip()
        converter = HTMLTableConverter()
        
        try:
            data = json.loads(input_data)
            html_content = data.get('html', '')
            output_file = data.get('output', 'output.xlsx')
        except json.JSONDecodeError:
            html_content = input_data
            output_file = 'output.xlsx'
        
        if not html_content:
            print(json.dumps({"error": "No HTML content provided"}), file=sys.stderr)
            sys.exit(1)
            
        result = converter.convert(html_content, output_file)
        print(json.dumps(result))
        sys.exit(0 if result["success"] else 1)
        
    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(1) 