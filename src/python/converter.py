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
        cssutils.log.setLevel(logging.CRITICAL)
        
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

    def parse_style(self, element):
        """Parse style attribute from HTML element"""
        styles = {}
        style = element.get('style', '')
        if style:
            try:
                parsed = cssutils.parseStyle(style)
                for property in parsed:
                    styles[property.name] = property.value
            except:
                pass
        return styles

    def css_to_rgb(self, color):
        """Convert CSS color to RGB tuple"""
        if not color:
            return None
            
        try:
            if color.startswith('#'):
                # Handle hex colors
                color = color.lstrip('#')
                if len(color) == 3:
                    color = ''.join(c + c for c in color)
                return tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
                
            elif color.startswith('rgb'):
                # Handle rgb/rgba colors
                return tuple(map(int, re.findall(r'\d+', color)[:3]))
        except:
            pass
            
        return None

    def apply_header_style(self, cell):
        """Apply header styles to cell"""
        cell.font = self.header_font
        cell.alignment = self.default_alignment
        cell.border = self.default_border
        cell.fill = PatternFill(
            start_color='D9D9D9',
            end_color='D9D9D9',
            fill_type='solid'
        )

    def apply_cell_style(self, cell, styles):
        """Apply styles to Excel cell"""
        try:
            # Font
            font_props = {'name': 'TH Sarabun New', 'size': 10}
            if 'color' in styles:
                rgb = self.css_to_rgb(styles['color'])
                if rgb:
                    font_props['color'] = f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}'
            if 'font-weight' in styles:
                font_props['bold'] = styles['font-weight'] in ('bold', '700')
            cell.font = Font(**font_props)

            # Alignment
            align_props = {'wrap_text': True}
            if 'text-align' in styles:
                align_props['horizontal'] = styles['text-align']
            else:
                align_props['horizontal'] = 'center'
            align_props['vertical'] = 'center'
            cell.alignment = Alignment(**align_props)

            # Border
            cell.border = self.default_border

            # Background
            if 'background-color' in styles:
                rgb = self.css_to_rgb(styles['background-color'])
                if rgb:
                    cell.fill = PatternFill(
                        start_color=f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}',
                        end_color=f'{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}',
                        fill_type='solid'
                    )
        except Exception as e:
            print(f"Error applying style: {str(e)}")

    def process_table(self, ws, table, start_row):
        """Process a single table and return the next row number"""
        current_row = start_row
        max_col = 0
        merged_cells = set()
        header_rows = []
        footer_rows = []
        body_rows = []

        # แยก header, body, footer rows
        rows = table.find_all('tr')
        for row in rows:
            if row.find('th'):
                header_rows.append(row)
            elif row.parent and row.parent.name == 'tfoot':
                footer_rows.append(row)
            else:
                body_rows.append(row)

        # Process header rows
        for row in header_rows:
            current_col = 1
            for cell in row.find_all(['th']):
                while (current_row, current_col) in merged_cells:
                    current_col += 1

                value = cell.get_text(strip=True)
                excel_cell = ws.cell(row=current_row, column=current_col)
                excel_cell.value = value
                self.apply_header_style(excel_cell)

                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))

                if rowspan > 1 or colspan > 1:
                    ws.merge_cells(
                        start_row=current_row,
                        start_column=current_col,
                        end_row=current_row + rowspan - 1,
                        end_column=current_col + colspan - 1
                    )
                    for r in range(current_row, current_row + rowspan):
                        for c in range(current_col, current_col + colspan):
                            merged_cells.add((r, c))

                current_col += colspan
            max_col = max(max_col, current_col - 1)
            current_row += 1

        # Process body rows
        for row in body_rows:
            current_col = 1
            for cell in row.find_all(['td']):
                while (current_row, current_col) in merged_cells:
                    current_col += 1

                value = cell.get_text(strip=True)
                excel_cell = ws.cell(row=current_row, column=current_col)
                excel_cell.value = value
                
                styles = self.parse_style(cell)
                self.apply_cell_style(excel_cell, styles)

                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))

                if rowspan > 1 or colspan > 1:
                    ws.merge_cells(
                        start_row=current_row,
                        start_column=current_col,
                        end_row=current_row + rowspan - 1,
                        end_column=current_col + colspan - 1
                    )
                    for r in range(current_row, current_row + rowspan):
                        for c in range(current_col, current_col + colspan):
                            merged_cells.add((r, c))

                current_col += colspan
            max_col = max(max_col, current_col - 1)
            current_row += 1

        # Process footer rows
        for row in footer_rows:
            current_col = 1
            for cell in row.find_all(['td']):
                while (current_row, current_col) in merged_cells:
                    current_col += 1

                value = cell.get_text(strip=True)
                excel_cell = ws.cell(row=current_row, column=current_col)
                excel_cell.value = value
                
                styles = self.parse_style(cell)
                self.apply_cell_style(excel_cell, styles)
                excel_cell.font = Font(name='TH Sarabun New', size=10, bold=True)

                rowspan = int(cell.get('rowspan', 1))
                colspan = int(cell.get('colspan', 1))

                if rowspan > 1 or colspan > 1:
                    ws.merge_cells(
                        start_row=current_row,
                        start_column=current_col,
                        end_row=current_row + rowspan - 1,
                        end_column=current_col + colspan - 1
                    )
                    for r in range(current_row, current_row + rowspan):
                        for c in range(current_col, current_col + colspan):
                            merged_cells.add((r, c))

                current_col += colspan
            max_col = max(max_col, current_col - 1)
            current_row += 1

        return current_row + 1, max_col

    def convert(self, html_content, output):
        """Convert HTML content to Excel file"""
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            
            soup = BeautifulSoup(html_content, 'html.parser')
            tables = soup.find_all('table')
            
            if not tables:
                raise ValueError("No tables found in HTML content")

            current_row = 1
            max_col = 0

            for table in tables:
                next_row, table_max_col = self.process_table(ws, table, current_row)
                current_row = next_row
                max_col = max(max_col, table_max_col)

            # Adjust column widths
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

            if isinstance(output, str):
                wb.save(output)
            else:
                wb.save(output)

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