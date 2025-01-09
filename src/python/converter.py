import sys
import logging
import re
import json
import xlsxwriter
import time
import gc
from itertools import islice

try:
    from bs4 import BeautifulSoup, SoupStrainer
except ImportError as e:
    print(json.dumps({
        "success": False,
        "error": f"Error importing beautifulsoup4: {str(e)}. Please run: pip install beautifulsoup4"
    }))
    sys.exit(1)

try:
    import cssutils
except ImportError as e:
    print(json.dumps({
        "success": False,
        "error": f"Error importing cssutils: {str(e)}. Please run: pip install cssutils"
    }))
    sys.exit(1)

logging.basicConfig(level=logging.INFO, 
                   format='%(message)s',
                   stream=sys.stderr)
logger = logging.getLogger(__name__)
cssutils.log.setLevel(logging.CRITICAL)

class HTMLToExcelConverter:
    def __init__(self, chunk_size=1000):
        """Initialize converter with chunk size for memory management"""
        self.chunk_size = chunk_size
        cssutils.log.setLevel(logging.CRITICAL)
        self.default_font = {'name': 'TH Sarabun New', 'size': 10}
        self.header_font = {'name': 'TH Sarabun New', 'size': 10, 'bold': True}
        self._color_cache = {}
        self._format_cache = {}

    def _get_format(self, workbook, key, properties):
        """Get cached format or create new one"""
        if key not in self._format_cache:
            self._format_cache[key] = workbook.add_format(properties)
        return self._format_cache[key]

    def _get_alignment(self, element):
        """Get text alignment from element's class or style"""
        style = element.get('style', '')
        classes = element.get('class', [])
        
        if isinstance(classes, str):
            classes = classes.split()
            
        if 'text-right' in classes or 'text-align: right' in style:
            return 'right'
        elif 'text-left' in classes or 'text-align: left' in style:
            return 'left'
        return 'center'

    def _process_header_section(self, soup, worksheet, current_row, formats, workbook):
        """Process header section including company info and metadata"""
        # Process timestamp header only once
        timestamp_header = None
        timesheet_text = None
        
        # Find all headers to process only the rightmost timestamp
        all_headers = soup.find_all('th', string=lambda x: x and ('Printed' in x or 'พิมพ์' in x))
        if all_headers:
            timestamp_header = all_headers[-1]  # Use only the last (rightmost) timestamp
            
        if timestamp_header:
            worksheet.write(current_row, 12, timestamp_header.get_text(strip=True),
                self._get_format(workbook, 'timestamp', {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'align': 'right'
                }))
            current_row += 2

        processed_info = set()  # Keep track of processed information
        
        # Process both tem1 and tem2 formats
        def process_header_row(row, is_tem2=False):
            nonlocal current_row
            cells = row.find_all(['th'])
            if not cells:
                return
                
            # Get text content
            row_text = " ".join(cell.get_text(strip=True) for cell in cells)
            if row_text in processed_info:
                return
                
            # Skip empty rows
            if not row_text.strip():
                return

            # Skip if this is a timestamp row we already processed
            if 'Printed' in row_text or 'พิมพ์' in row_text:
                return
                
            processed_info.add(row_text)
            
            # Determine format based on content
            align = 'center'  # default alignment
            if 'text-left' in row.get('class', []) or any('text-left' in cell.get('class', []) for cell in cells):
                align = 'left'
            elif 'text-right' in row.get('class', []) or any('text-right' in cell.get('class', []) for cell in cells):
                align = 'right'
            
            # Special handling for date range row
            if 'ระหว่างวันที่' in row_text or 'Time sheet' in row_text:
                fmt = self._get_format(workbook, 'date_range', {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'align': 'center',
                    'valign': 'vcenter'
                })
                worksheet.merge_range(current_row, 0, current_row, 12, row_text, fmt)
                current_row += 1
                return
            
            # For tem2 format with specific column spans
            if is_tem2 and len(cells) >= 2 and ('Driver Name' in row_text or 'Boss Name' in row_text or 'Position' in row_text):
                # First part (0-7)
                first_text = cells[0].get_text(strip=True)
                worksheet.merge_range(current_row, 0, current_row, 7, first_text,
                    self._get_format(workbook, 'driver_info', {
                        'font_name': 'TH Sarabun New',
                        'font_size': 10,
                        'align': 'left',
                        'valign': 'vcenter'
                    }))
                # Second part (8-12)
                second_text = cells[1].get_text(strip=True)
                worksheet.merge_range(current_row, 8, current_row, 12, second_text,
                    self._get_format(workbook, 'driver_info', {
                        'font_name': 'TH Sarabun New',
                        'font_size': 10,
                        'align': 'left',
                        'valign': 'vcenter'
                    }))
            # For tem1 format with 4 columns driver info
            elif len(cells) >= 4 and ('ชื่อคนขับ' in row_text or 'Driver Name' in row_text):
                # First pair (0-5)
                first_pair = cells[0].get_text(strip=True) + " " + cells[1].get_text(strip=True)
                worksheet.merge_range(current_row, 0, current_row, 5, first_pair,
                    self._get_format(workbook, 'driver_info', {
                        'font_name': 'TH Sarabun New',
                        'font_size': 10,
                        'align': 'left',
                        'valign': 'vcenter'
                    }))
                # Second pair (6-12)
                second_pair = cells[2].get_text(strip=True) + " " + cells[3].get_text(strip=True)
                worksheet.merge_range(current_row, 6, current_row, 12, second_pair,
                    self._get_format(workbook, 'driver_info', {
                        'font_name': 'TH Sarabun New',
                        'font_size': 10,
                        'align': 'left',
                        'valign': 'vcenter'
                    }))
            else:
                # Normal row
                fmt_key = 'company' if 'บริษัท' in row_text else 'customer'
                # Use customer format for customer name rows
                if 'Customer Name' in row_text or 'ชื่อลูกค้า' in row_text:
                    worksheet.merge_range(current_row, 0, current_row, 12, row_text, formats['customer'])
                else:
                    fmt = self._get_format(workbook, fmt_key, {
                        'font_name': 'TH Sarabun New',
                        'font_size': 10,
                        'bold': 'บริษัท' in row_text,
                        'align': align,
                        'valign': 'vcenter'
                    })
                    worksheet.merge_range(current_row, 0, current_row, 12, row_text, fmt)
            
            current_row += 1

        # First try tem1 format (nested table)
        tem1_processed = False
        for table in soup.find_all('table'):
            if table.get('style') and 'margin-bottom: 5px' in table.get('style'):
                thead = table.find('thead')
                if thead:
                    for row in thead.find_all('tr'):
                        process_header_row(row)
                    tem1_processed = True
                    break

        # If tem1 wasn't processed, try tem2 format
        if not tem1_processed:
            for table in soup.find_all('table'):
                thead = table.find('thead')
                if thead and any('head-paper' in th.get('class', []) for th in thead.find_all('th')):
                    for row in thead.find_all('tr'):
                        process_header_row(row, is_tem2=True)
                    break

        current_row += 1  # Add space after header
        return current_row

    def _process_table_headers(self, table, worksheet, current_row, formats, workbook):
        """Process table headers with support for both formats"""
        header_rows = []
        
        # Find the header row with column titles
        for tr in table.find_all('tr'):
            # Skip nested table headers
            if tr.find_parent('table') != table:
                continue
                
            # Check if this is the header row with column titles
            cells = tr.find_all(['th'])
            if cells and any(cell.get_text(strip=True) in ['วันที่', 'วัน', 'เริ่ม', 'สิ้นสุด'] for cell in cells):
                header_rows.append(tr)
                break

        for tr in header_rows:
            headers = tr.find_all(['th'])
            for col, th in enumerate(headers):
                text = th.get_text(strip=True)
                
                # Get width from style or use default
                style = th.get('style', '')
                width = None
                if 'width:' in style:
                    try:
                        width_str = style.split('width:')[1].split(';')[0].strip()
                        if 'px' in width_str:
                            width = float(width_str.replace('px', '')) / 8
                        elif '%' in width_str:
                            width = float(width_str.replace('%', '')) / 8
                    except:
                        width = None
                
                if width is None:
                    width = len(text) * 1.2
                
                worksheet.set_column(col, col, max(width, 8))
                
                # Get background color
                bg_color = None
                if 'background-color:' in style:
                    bg_color = style.split('background-color:')[1].split(';')[0].strip()
                
                # Create format with background color if specified
                if bg_color:
                    header_format = self._get_format(workbook, f'header_{bg_color}', {
                        'font_name': 'TH Sarabun New',
                        'font_size': 10,
                        'bold': True,
                        'align': 'center',
                        'valign': 'vcenter',
                        'border': 1,
                        'bg_color': bg_color
                    })
                else:
                    header_format = formats['header']
                
                worksheet.write(current_row, col, text, header_format)

            worksheet.set_row(current_row, 30)
            current_row += 1

        return current_row

    def _process_table_body(self, table, worksheet, current_row, formats, workbook):
        """Process table body with support for both formats"""
        rows = []
        
        # Find all rows that are not in thead and not summary rows
        all_rows = table.find_all('tr')
        body_rows = []
        
        for row in all_rows:
            # Skip if row is in thead
            if row.find_parent('thead'):
                continue
                
            # Skip if row contains th (header)
            if row.find('th'):
                continue
                
            # Skip if row is a summary/total row
            cells = row.find_all(['td'])
            if cells and any('total' in cell.get_text().lower() for cell in cells):
                continue
                
            body_rows.append(row)

        for tr in body_rows:
            cells = tr.find_all(['td'])
            if not cells:
                continue

            row_data = []
            for td in cells:
                value = td.get_text(strip=True)
                align = self._get_alignment(td)
                
                # Get background color if any
                style = td.get('style', '')
                bg_color = None
                if 'background-color' in style:
                    bg_color = style.split('background-color:')[1].split(';')[0].strip()
                
                fmt_props = {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'align': align,
                    'valign': 'vcenter',
                    'border': 1
                }
                
                if bg_color:
                    fmt_props['bg_color'] = bg_color
                
                fmt = self._get_format(workbook, f'cell_{align}_{bg_color if bg_color else "default"}', fmt_props)
                
                row_data.append((value, fmt))
            
            rows.append(row_data)

        # Write rows in chunks
        for i in range(0, len(rows), self.chunk_size):
            chunk = rows[i:i + self.chunk_size]
            for row_data in chunk:
                for col, (value, fmt) in enumerate(row_data):
                    worksheet.write(current_row, col, value, fmt)
                current_row += 1
            gc.collect()

        return current_row

    def _process_footer_section(self, soup, worksheet, current_row, formats):
        """Process footer with support for both formats"""
        # Try new format (div with footer class)
        footer = soup.find('div', class_='footer')
        if footer:
            col = 0
            for div in footer.find_all('div'):
                text = div.get_text(strip=True)
                if text:
                    worksheet.write(current_row, col, text, formats['footer'])
                    col += 4
            current_row += 1
        else:
            # Try old format (tfoot)
            footer = soup.find('tfoot')
            if footer:
                for tr in footer.find_all('tr'):
                    col = 0
                    for td in tr.find_all(['td']):
                        text = td.get_text(strip=True)
                        if text:
                            worksheet.write(current_row, col, text, formats['footer'])
                            col += 4
                    current_row += 1

        return current_row

    def _process_table_section(self, soup, worksheet, current_row, formats, workbook):
        """Process table section with support for nested tables"""
        # Find the main data table using multiple criteria
        main_table = None
        
        # Try to find table with border-collapse style (tem1 format)
        for table in soup.find_all('table'):
            if table.get('style') and 'border-collapse: collapse' in table.get('style'):
                main_table = table
                break
        
        # If not found, try to find table with table-report class (tem2 format)
        if not main_table:
            for table in soup.find_all('table'):
                tbody = table.find('tbody', class_='table-report')
                if tbody:
                    main_table = table
                    break

        # If still not found, try to find the largest table with headers
        if not main_table:
            max_cells = 0
            for table in soup.find_all('table'):
                headers = table.find_all(['th'])
                if headers and len(headers) > max_cells:
                    if any(header.get_text(strip=True) in ['วันที่', 'วัน', 'เริ่ม', 'สิ้นสุด'] for header in headers):
                        main_table = table
                        max_cells = len(headers)

        if main_table:
            # Set column widths
            column_widths = [8, 5, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 15]  # Adjust these values as needed
            for col, width in enumerate(column_widths):
                worksheet.set_column(col, col, width)

            # Process headers - Pass workbook parameter
            current_row = self._process_table_headers(main_table, worksheet, current_row, formats, workbook)
            
            # Find tbody or process all rows if no tbody
            tbody = main_table.find('tbody', class_='table-report')
            if not tbody:
                tbody = main_table.find('tbody')
            if tbody:
                current_row = self._process_table_body(tbody, worksheet, current_row, formats, workbook)
            else:
                # If no tbody, process all non-header rows
                current_row = self._process_table_body(main_table, worksheet, current_row, formats, workbook)

        return current_row

    def convert(self, html_content, output):
        """Convert HTML to Excel with support for all formats"""
        try:
            start_time = time.time()
            logger.info("Converting HTML to Excel...")
            
            workbook = xlsxwriter.Workbook(output, {'constant_memory': True})
            worksheet = workbook.add_worksheet('Sheet1')
            
            # Pre-define formats
            formats = {
                'default': self._get_format(workbook, 'default', {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1
                }),
                'header': self._get_format(workbook, 'header', {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1,
                    'bg_color': '#a9a9a9'
                }),
                'customer': self._get_format(workbook, 'customer', {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'align': 'left',
                    'valign': 'vcenter'
                }),
                'footer': self._get_format(workbook, 'footer', {
                    'font_name': 'TH Sarabun New',
                    'font_size': 10,
                    'align': 'left'
                })
            }
            
            current_row = 0
            
            # Split HTML into documents if multiple exist
            documents = [doc.strip() for doc in html_content.split('</html>') if doc.strip()]
            if not documents:
                documents = [html_content]

            for doc in documents:
                if not doc.endswith('</html>'):
                    doc += '</html>'
                
                soup = BeautifulSoup(doc, 'html.parser')
                
                # Process each section
                current_row = self._process_header_section(soup, worksheet, current_row, formats, workbook)
                current_row = self._process_table_section(soup, worksheet, current_row, formats, workbook)
                current_row = self._process_footer_section(soup, worksheet, current_row, formats)
                current_row += 2  # Add space between documents
            
            workbook.close()
            logger.info(f"Conversion completed in {time.time() - start_time:.2f} seconds")
            print(json.dumps({"success": True}))
            
        except Exception as e:
            logger.error(f"Error: {str(e)}")
            print(json.dumps({"success": False, "error": str(e)}), file=sys.stderr)
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
            logger.error(f"Error converting file: {str(e)}")
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