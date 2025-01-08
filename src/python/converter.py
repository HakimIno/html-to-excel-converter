import sys
import logging
import re
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("Error: beautifulsoup4 is not installed. Please run: pip install beautifulsoup4")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("Error: pandas is not installed. Please run: pip install pandas")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is not installed. Please run: pip install openpyxl")
    sys.exit(1)

try:
    import cssutils
except ImportError:
    print("Error: cssutils is not installed. Please run: pip install cssutils")
    sys.exit(1)

# ตั้งค่า logging level สำหรับ cssutils
cssutils.log.setLevel(logging.CRITICAL)

class HTMLToExcelConverter:
    def __init__(self):
        self.default_font = Font(name='TH Sarabun New', size=10)
        self.default_alignment = Alignment(horizontal='center', vertical='center')
        self.default_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        self.header_font = Font(name='TH Sarabun New', size=10, bold=True)
        self.current_row = 1

    def get_cell_dimensions(self, cell):
        """Get rowspan and colspan of a cell"""
        rowspan = int(cell.get('rowspan', 1))
        colspan = int(cell.get('colspan', 1))
        return rowspan, colspan

    def process_cell(self, ws, cell, row, col, merged_cells):
        """Process a single cell and handle merging"""
        rowspan, colspan = self.get_cell_dimensions(cell)
        
        # Get cell content
        value = cell.get_text(strip=True)
        excel_cell = ws.cell(row=row, column=col)
        excel_cell.value = value
        
        # Apply styles
        styles = self.parse_style(cell)
        
        # Set default style for header cells
        if cell.name == 'th':
            styles['font-weight'] = 'bold'
        
        self.apply_cell_style(excel_cell, styles)
        
        # Handle merged cells
        if rowspan > 1 or colspan > 1:
            merge_range = f"{get_column_letter(col)}{row}:{get_column_letter(col+colspan-1)}{row+rowspan-1}"
            try:
                ws.merge_cells(merge_range)
                
                # Track merged areas
                for r in range(row, row + rowspan):
                    for c in range(col, col + colspan):
                        merged_cells.add((r, c))
            except:
                print(f"Warning: Could not merge cells {merge_range}")
        
        return colspan

    def css_to_rgb(self, color):
        """Convert CSS color to RGB"""
        if not color:
            return None
            
        if color.startswith('#'):
            color = color[1:]
            if len(color) != 6:
                return None
            try:
                return tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
            except ValueError:
                return None
            
        rgb_match = re.search(r'rgb\((\d+),\s*(\d+),\s*(\d+)\)', color)
        if rgb_match:
            try:
                return tuple(map(int, rgb_match.groups()))
            except ValueError:
                return None
            
        return None

    def parse_style(self, element):
        """Parse CSS styles from element"""
        styles = {}
        
        # Get inline styles
        if element.get('style'):
            try:
                style = cssutils.parseStyle(element['style'])
                for prop in style:
                    styles[prop.name] = prop.value
            except:
                pass
                
        # Get background color from HTML attributes
        if element.get('bgcolor'):
            styles['background-color'] = element['bgcolor']
            
        return styles

    def apply_cell_style(self, cell, styles):
        """Apply styles to Excel cell"""
        try:
            # Background color
            if 'background-color' in styles:
                rgb = self.css_to_rgb(styles['background-color'])
                if rgb:
                    cell.fill = PatternFill(
                        start_color=f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}',
                        end_color=f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}',
                        fill_type='solid'
                    )

            # Font
            font_props = {'name': 'TH Sarabun New', 'size': 10}  # Default font properties
            if 'color' in styles:
                rgb = self.css_to_rgb(styles['color'])
                if rgb:
                    font_props['color'] = f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
            if 'font-size' in styles:
                try:
                    size_match = re.search(r'(\d+)', styles['font-size'])
                    if size_match:
                        size = float(size_match.group(1))
                        font_props['size'] = size
                except:
                    pass
            if 'font-weight' in styles:
                font_props['bold'] = styles['font-weight'] in ('bold', '700')
                
            cell.font = Font(**font_props)

            # Alignment
            align_props = {'wrap_text': True}
            if 'text-align' in styles:
                if styles['text-align'] == 'left':
                    align_props['horizontal'] = 'left'
                elif styles['text-align'] == 'right':
                    align_props['horizontal'] = 'right'
                else:
                    align_props['horizontal'] = 'center'
            else:
                align_props['horizontal'] = 'center'
                
            align_props['vertical'] = 'center'  # Default vertical alignment
            cell.alignment = Alignment(**align_props)

            # Border
            cell.border = self.default_border
        except Exception as e:
            print(f"Error applying style: {str(e)}")

    def process_table(self, wb, ws, table, start_row):
        """Process a single table and return the next row number"""
        current_row = start_row
        max_col = 0
        merged_cells = set()

        # Process each row
        for row in table.find_all(['tr']):
            current_col = 1
            row_height = 0

            # Process each cell in the row
            for cell in row.find_all(['td', 'th']):
                # Skip if cell is in merged area
                while (current_row, current_col) in merged_cells:
                    current_col += 1

                colspan = self.process_cell(ws, cell, current_row, current_col, merged_cells)
                current_col += colspan

            # Update maximum column count
            max_col = max(max_col, current_col - 1)
            
            # Move to next row
            current_row += 1

        return current_row + 1, max_col  # Return next row with 1 row gap

    def convert(self, html_content, output):
        """Convert HTML content to Excel file"""
        try:
            # Create workbook and select active sheet
            wb = openpyxl.Workbook()
            ws = wb.active
            
            # Parse HTML content
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # Find all tables
            tables = soup.find_all('table')
            if not tables:
                raise ValueError("No tables found in HTML content")

            current_row = 1
            max_col = 0

            # Process each table
            for table in tables:
                next_row, table_max_col = self.process_table(wb, ws, table, current_row)
                current_row = next_row  # Update current row for next table
                max_col = max(max_col, table_max_col)

            # Adjust column widths for all processed tables
            for col in range(1, max_col + 1):
                max_length = 0
                column = get_column_letter(col)
                
                for cell in ws[column]:
                    if cell.value:
                        try:
                            max_length = max(max_length, len(str(cell.value)))
                        except:
                            max_length = max(max_length, 30)
                            
                adjusted_width = min(max_length + 2, 30)
                ws.column_dimensions[column].width = adjusted_width

            # Save workbook
            if isinstance(output, str):
                wb.save(output)
            else:
                wb.save(output)  # Save to buffer

        except Exception as e:
            print(f"Error converting HTML to Excel: {str(e)}")
            raise

    @classmethod
    def convert_file(cls, input_path, output_path):
        """Convert HTML file to Excel file"""
        try:
            with open(input_path, 'r', encoding='utf-8') as file:
                html_content = file.read()
            
            converter = cls()
            converter.convert(html_content, output_path)
            
        except Exception as e:
            print(f"Error converting file: {str(e)}")
            raise

if __name__ == "__main__":
    # ตรวจสอบ command line arguments
    if len(sys.argv) < 2:
        print("กรุณาระบุชื่อไฟล์ HTML")
        print("วิธีใช้: python converter.py input.html [output.xlsx]")
        sys.exit(1)

    # รับชื่อไฟล์ input
    input_file = sys.argv[1]
    
    # กำหนดชื่อไฟล์ output (ถ้าไม่ระบุจะใช้ชื่อเดียวกับ input แต่เปลี่ยนนามสกุล)
    output_file = sys.argv[2] if len(sys.argv) > 2 else input_file.rsplit('.', 1)[0] + '.xlsx'

    # แปลงไฟล์
    try:
        HTMLToExcelConverter.convert_file(input_file, output_file)
    except Exception as e:
        print(f"ไม่สามารถแปลงไฟล์ได้: {str(e)}")
        sys.exit(1) 