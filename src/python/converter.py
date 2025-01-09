import sys
import logging
import re
import json
import xlsxwriter
import time

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("Error: beautifulsoup4 is not installed. Please run: pip install beautifulsoup4")
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
        self.default_font = {'name': 'TH Sarabun New', 'size': 10}
        self.header_font = {'name': 'TH Sarabun New', 'size': 10, 'bold': True}

    def css_to_rgb(self, color):
        """Convert CSS color to RGB tuple - optimized version"""
        if not color:
            return None
        
        # Cache สำหรับสี
        if not hasattr(self, '_color_cache'):
            self._color_cache = {}
            
        if color in self._color_cache:
            return self._color_cache[color]
            
        try:
            if color.startswith('#'):
                color = color.lstrip('#')
                if len(color) == 3:
                    color = ''.join(c + c for c in color)
                rgb = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
                self._color_cache[color] = rgb
                return rgb
                
            elif color.startswith('rgb'):
                rgb = tuple(map(int, re.findall(r'\d+', color)[:3]))
                self._color_cache[color] = rgb
                return rgb
        except:
            pass
        return None

    def convert(self, html_content, output):
        """Convert HTML to Excel using xlsxwriter for better performance"""
        try:
            start_time = time.time()
            
            # สร้าง workbook ด้วย xlsxwriter
            workbook = xlsxwriter.Workbook(output)
            worksheet = workbook.add_worksheet('Sheet1')

            # Define styles once
            default_format = workbook.add_format({
                'font_name': 'TH Sarabun New',
                'font_size': 10,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            header_format = workbook.add_format({
                'font_name': 'TH Sarabun New',
                'font_size': 10,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })

            customer_format = workbook.add_format({
                'font_name': 'TH Sarabun New',
                'font_size': 10,
                'align': 'left',
                'valign': 'vcenter'
            })

            footer_format = workbook.add_format({
                'font_name': 'TH Sarabun New',
                'font_size': 10,
                'align': 'left'
            })

            current_row = 0
            max_col = 13

            # แยก HTML documents
            html_documents = [doc for doc in html_content.split('</html>') if doc.strip()]
            
            for doc in html_documents:
                # เพิ่ม closing tag กลับเข้าไป
                if not doc.endswith('</html>'):
                    doc = doc + '</html>'
                
                # Parse HTML
                soup = BeautifulSoup(doc, 'html.parser')

                # Process timestamp header
                timestamp_header = soup.find('th', string=lambda x: x and 'Printed' in x)
                if timestamp_header:
                    worksheet.write(current_row, max_col-3, timestamp_header.get_text(strip=True),
                        workbook.add_format({
                            'font_name': 'TH Sarabun New',
                            'font_size': 10,
                            'align': 'right'
                        }))
                    current_row += 2

                # Process company info
                company_table = soup.find('table', attrs={'style': lambda x: x and 'margin-bottom: 5px' in x})
                if company_table:
                    # Company name
                    company_name = company_table.find('th', string=lambda x: x and 'บริษัท' in x)
                    if company_name:
                        worksheet.merge_range(current_row, 0, current_row, max_col-1,
                            company_name.get_text(strip=True), header_format)
                        current_row += 1

                    # Date range
                    date_range = company_table.find('th', string=lambda x: x and 'ระหว่างวันที่' in x)
                    if date_range:
                        worksheet.merge_range(current_row, 0, current_row, max_col-1,
                            date_range.get_text(strip=True), default_format)
                        current_row += 1

                    # Customer info
                    customer_rows = company_table.find_all('tr')
                    for tr in customer_rows[2:]:  # Skip company name and date range
                        cols = tr.find_all(['th'])
                        if not cols:
                            continue
                        
                        col_index = 0
                        for i in range(0, len(cols), 2):
                            if i + 1 < len(cols):
                                label = cols[i].get_text(strip=True)
                                value = cols[i + 1].get_text(strip=True)
                                if label and value:
                                    worksheet.write(current_row, col_index, label, customer_format)
                                    worksheet.write(current_row, col_index + 1, value, customer_format)
                                    col_index += 4
                        current_row += 1
                    
                    current_row += 1  # Add space after customer info

                # Process main table efficiently
                main_table = None
                for table in soup.find_all('table'):
                    # Skip header table
                    if table.get('style') and 'margin-bottom: 5px' in table.get('style'):
                        continue
                    # Find table with header row containing วันที่, เริ่ม, สิ้นสุด
                    header_row = table.find('tr', attrs={'style': lambda x: x and 'text-align: center' in x})
                    if header_row and header_row.find('th', string='วันที่'):
                        main_table = table
                        break
                
                if main_table:
                    # Process headers
                    header_row = main_table.find('tr', attrs={
                        'style': lambda x: x and 'text-align: center' in x and 'font-size: 10px' in x
                    })
                    
                    if header_row:
                        for col, th in enumerate(header_row.find_all('th')):
                            worksheet.write(current_row, col, th.get_text(strip=True), header_format)
                            # Set column width based on content
                            width = len(th.get_text(strip=True)) * 1.2  # Adjust multiplier as needed
                            worksheet.set_column(col, col, max(width, 8))  # Minimum width of 8
                        worksheet.set_row_pixels(current_row, 30)  # Set row height
                        current_row += 1

                    # Process table body efficiently
                    rows = []
                    for tr in main_table.find_all('tr'):
                        if tr == header_row:
                            continue
                        
                        row_data = []
                        cells = tr.find_all(['td'])
                        
                        if not cells:
                            continue
                        
                        for td in cells:
                            value = td.get_text(strip=True)
                            style = td.get('style', '')
                            
                            cell_format = workbook.add_format({
                                'font_name': 'TH Sarabun New',
                                'font_size': 10,
                                'border': 1,
                                'align': 'left' if 'text-align: left' in style else 
                                        'right' if 'text-align: right' in style else 'center'
                            })
                            
                            row_data.append((value, cell_format))
                        
                        rows.append(row_data)

                    # Write all rows at once
                    for row_data in rows:
                        for col, (value, fmt) in enumerate(row_data):
                            worksheet.write(current_row, col, value, fmt)
                        current_row += 1

                # Process footer
                current_row += 1  # Add space before footer
                footer = soup.find('tfoot')
                if footer:
                    for tr in footer.find_all('tr'):
                        col = 0
                        for td in tr.find_all(['td']):
                            text = td.get_text(strip=True)
                            if text:
                                worksheet.write(current_row, col, text, footer_format)
                                col += 4  # Space between footer columns
                        current_row += 1

                # Add extra space between documents
                current_row += 2

            workbook.close()
            
            print(f"Conversion completed in {time.time() - start_time:.2f} seconds")
            print(json.dumps({"success": True}))

        except Exception as e:
            print(json.dumps({"error": str(e)}), file=sys.stderr)
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
    try:
        input_data = sys.stdin.read().strip()
        converter = HTMLToExcelConverter()
        
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
            
        converter.convert(html_content, output_file)
        sys.exit(0)
        
    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(1) 