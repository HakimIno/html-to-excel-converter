import sys
import json
import xlsxwriter
import re
import html
from bs4 import BeautifulSoup, SoupStrainer
from functools import lru_cache
from typing import Dict, List, Set, Tuple, Union
import gc

class ExcelStyleManager:
    """Manages Excel styles and formatting with efficient caching"""
    
    def __init__(self):
        self._format_cache = {}
        self._style_hash_cache = {}
        self._color_cache = {}
        
    @lru_cache(maxsize=1000)
    def _normalize_color(self, color: str) -> str:
        """Normalize color values with caching"""
        if not color:
            return None
        if color in self._color_cache:
            return self._color_cache[color]
            
        try:
            if color.startswith('rgb'):
                r, g, b = map(int, re.findall(r'\d+', color))
                hex_color = f'#{r:02x}{g:02x}{b:02x}'
            elif color.startswith('#'):
                hex_color = color
            else:
                hex_color = color
            self._color_cache[color] = hex_color.upper()
            return hex_color.upper()
        except:
            return None

    @lru_cache(maxsize=1000)
    def _parse_style(self, style: str) -> Dict:
        """Parse CSS style string with caching"""
        if not style:
            return {}
        if style in self._style_hash_cache:
            return self._style_hash_cache[style]
            
        style_dict = {}
        for item in style.split(';'):
            if ':' in item:
                prop, value = item.split(':', 1)
                style_dict[prop.strip()] = value.strip()
        
        self._style_hash_cache[style] = style_dict
        return style_dict

    def get_format(self, workbook, properties: Dict) -> object:
        """Get cached format or create new one"""
        format_key = hash(frozenset(properties.items()))
        if format_key not in self._format_cache:
            self._format_cache[format_key] = workbook.add_format(properties)
        return self._format_cache[format_key]

class HTMLTableConverter:
    """Converts HTML tables to Excel with high performance and memory efficiency"""
    
    DEFAULT_OPTIONS = {
        'chunk_size': 1000,  # Process rows in chunks
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
            'row_height_multiplier': 1.2
        },
        'html': {
            'preserve_formatting': True,
            'parse_entities': True
        }
    }

    def __init__(self, options: Dict = None):
        self.options = self.DEFAULT_OPTIONS.copy()
        if options:
            self._deep_update(self.options, options)
            
        self.style_manager = ExcelStyleManager()
        self._reset_state()

    def _reset_state(self):
        """Reset internal state for new conversion"""
        self._column_widths = {}
        self._row_heights = {}
        self._merged_ranges = set()
        self._current_row = 0
        self._max_col = 0
        self._chunk_buffer = []
        gc.collect()  # Force garbage collection

    def _deep_update(self, d: Dict, u: Dict):
        """Deep update dictionary"""
        for k, v in u.items():
            if isinstance(v, dict) and k in d:
                self._deep_update(d[k], v)
            else:
                d[k] = v

    def _process_cell_content(self, cell: BeautifulSoup) -> Tuple[str, Dict]:
        """Process cell content and extract styles efficiently"""
        # Extract and combine styles
        styles = []
        classes = cell.get('class', [])
        if isinstance(classes, str):
            classes = classes.split()
            
        # Get inline style
        inline_style = self.style_manager._parse_style(cell.get('style', ''))
        
        # Base properties
        properties = {
            'font_name': self.options['font']['name'],
            'font_size': self.options['font']['size']
        }
        
        # Apply styles in correct order
        if inline_style.get('text-align'):
            properties['align'] = inline_style['text-align']
        elif 'text-right' in classes:
            properties['align'] = 'right'
        elif 'text-left' in classes:
            properties['align'] = 'left'
        else:
            properties['align'] = 'center'
            
        # Font styles
        if any(s in inline_style.get('font-weight', '') for s in ['bold', '700', '800', '900']):
            properties['bold'] = True
        if inline_style.get('font-style') == 'italic':
            properties['italic'] = True
            
        # Colors
        bg_color = inline_style.get('background-color')
        if bg_color:
            properties['bg_color'] = self.style_manager._normalize_color(bg_color)
            
        font_color = inline_style.get('color')
        if font_color:
            properties['font_color'] = self.style_manager._normalize_color(font_color)
            
        # Borders
        if 'border' in inline_style:
            self._process_borders(inline_style, properties)
            
        # Get text content
        if self.options['html']['preserve_formatting']:
            content = []
            for element in cell.children:
                if isinstance(element, str):
                    content.append(element.strip())
                elif element.name in ['br', 'p']:
                    content.append('\n')
                else:
                    content.append(element.get_text().strip())
            text = ' '.join(content).strip()
        else:
            text = cell.get_text(strip=True)
            
        if self.options['html']['parse_entities']:
            text = html.unescape(text)
            
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
        """Process table using conflict resolution matrix and handle nested tables"""
        self._chunk_buffer = []
        current_row = self._current_row
        
        # Create a map to store nested table positions
        nested_table_map = {}
        
        # First pass: Map all nested tables to their exact positions
        for cell in table.find_all(['td', 'th']):
            nested_tables = cell.find_all('table', recursive=False)
            if nested_tables:
                # Calculate cell's position
                row_pos = sum(1 for sibling in cell.parent.find_previous_siblings('tr')
                            if 'display: none' not in sibling.get('style', ''))
                col_pos = sum(1 for sibling in cell.find_previous_siblings(['td', 'th']))
                
                # Store nested tables with their parent cell info
                for idx, nested_table in enumerate(nested_tables):
                    nested_table_map[(row_pos, col_pos, idx)] = {
                        'table': nested_table,
                        'parent_cell': cell,
                        'rowspan': int(cell.get('rowspan', 1)),
                        'colspan': int(cell.get('colspan', 1))
                    }
                    # Replace nested table with placeholder
                    nested_table.replace_with(f'[NESTED_TABLE_{row_pos}_{col_pos}_{idx}]')

        # Resolve merge conflicts for main table
        resolved_matrix = self.resolve_merge_conflicts(table)
        
        # Process resolved matrix
        max_row_used = 0
        for row_idx, row in enumerate(resolved_matrix):
            row_cells = []
            col_idx = 0
            
            while col_idx < len(row):
                cell_info = row[col_idx]
                
                if cell_info and cell_info['element'] is not None:
                    # Only process cells that start here
                    is_start = True
                    if col_idx > 0:
                        prev_cell = row[col_idx - 1]
                        if prev_cell and prev_cell is cell_info:
                            is_start = False
                    if row_idx > 0 and is_start:
                        above_cell = resolved_matrix[row_idx - 1][col_idx]
                        if above_cell and above_cell is cell_info:
                            is_start = False
                            
                    if is_start:
                        content, properties = self._process_cell_content(cell_info['element'])
                        span = (cell_info['rowspan'], cell_info['colspan']) if cell_info['rowspan'] > 1 or cell_info['colspan'] > 1 else None
                        row_cells.append((content, properties, span))
                        
                        # Update max_row_used based on rowspan
                        if span:
                            max_row_used = max(max_row_used, row_idx + span[0])
                        
                    col_idx += cell_info['colspan']
                else:
                    col_idx += 1
                    
            if row_cells:
                self._chunk_buffer.append((current_row + row_idx, row_cells))
                max_row_used = max(max_row_used, row_idx + 1)
                
                if len(self._chunk_buffer) >= self.options['chunk_size']:
                    self._process_table_chunk(self._chunk_buffer, worksheet, workbook)
                    self._chunk_buffer = []
                    gc.collect()
        
        # Process remaining buffer
        if self._chunk_buffer:
            self._process_table_chunk(self._chunk_buffer, worksheet, workbook)
            self._chunk_buffer = []
            gc.collect()

        # Process nested tables in their exact positions
        for (row_pos, col_pos, idx), nested_info in sorted(nested_table_map.items()):
            nested_table = nested_info['table']
            parent_cell = nested_info['parent_cell']
            
            # Calculate absolute position for nested table
            nested_row_offset = current_row + row_pos
            nested_col_offset = col_pos
            
            # Save current state
            original_row = self._current_row
            original_merged = self._merged_ranges.copy()
            
            # Process nested table at exact position
            self._current_row = nested_row_offset
            self._process_table(nested_table, worksheet, workbook)
            
            # Restore merged ranges from parent table
            self._merged_ranges = original_merged
            
            # Update max_row_used if nested table extends beyond parent
            nested_rows_used = self._current_row - nested_row_offset
            max_row_used = max(max_row_used, row_pos + nested_rows_used)
            
            # Restore row counter
            self._current_row = original_row

        # Update final row position
        self._current_row = current_row + max_row_used + 1  # Add spacing

    def convert(self, html_content: str, output: Union[str, object]) -> Dict:
        """Convert HTML to Excel with high performance"""
        workbook = None
        try:
            self._reset_state()
            
            # Create workbook
            if isinstance(output, str):
                workbook = xlsxwriter.Workbook(output, {'constant_memory': True})
            else:
                workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'constant_memory': True})
            
            worksheet = workbook.add_worksheet('Sheet1')
            
            # Parse HTML
            soup = BeautifulSoup(html_content, 'html.parser', parse_only=SoupStrainer('table'))
            
            # Process tables
            for table in soup.find_all('table'):
                if 'display: none' not in table.get('style', ''):
                    self._process_table(table, worksheet, workbook)
            
            # Apply column widths
            if self.options['table']['auto_width']:
                for col, width in self._column_widths.items():
                    adjusted_width = min(
                        max(width, self.options['table']['min_width']),
                        self.options['table']['max_width']
                    )
                    worksheet.set_column(col, col, adjusted_width)
            
            workbook.close()
            return {"success": True}
            
        except Exception as e:
            if workbook:
                try:
                    workbook.close()
                except:
                    pass
            return {"success": False, "error": str(e)}
        finally:
            self._reset_state()

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