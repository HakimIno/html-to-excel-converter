#!/usr/bin/env python3
import sys
import json
import logging
import os
from dataclasses import dataclass
from typing import List, Tuple, Dict, Optional
from selectolax.parser import HTMLParser
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import time
from reportlab.lib.units import cm, inch
from reportlab.lib import pagesizes

@dataclass
class PDFOptions:
    page_size: str = 'A4'
    margin_left: float = 1.0
    margin_right: float = 1.0
    margin_top: float = 1.0
    margin_bottom: float = 1.0
    font_name: str = 'THSarabunNew'
    font_size: int = 9
    optimize_images: bool = True
    dpi: int = 300
    optimize_fonts: bool = True

class FontManager:
    FONT_PATHS = [
        os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'assets/TH')
    ]
    
    REQUIRED_FONTS = {
        'THSarabunNew': {
            'normal': 'THSarabunNew.ttf',
            'bold': 'THSarabunNew Bold.ttf',
            'italic': 'THSarabunNew Italic.ttf',
            'boldItalic': 'THSarabunNew BoldItalic.ttf'
        }
    }

    @classmethod
    def find_font_file(cls, font_name: str, style: str) -> Optional[str]:
        font_filename = cls.REQUIRED_FONTS[font_name][style]
        for font_path in cls.FONT_PATHS:
            full_path = os.path.join(font_path, font_filename)
            if os.path.exists(full_path):
                return full_path
        return None

    @classmethod
    def register_fonts(cls) -> None:
        for font_name, styles in cls.REQUIRED_FONTS.items():
            for style, _ in styles.items():
                font_path = cls.find_font_file(font_name, style)
                if font_path:
                    font_key = f"{font_name}-{style}"
                    try:
                        pdfmetrics.registerFont(TTFont(font_key, font_path))
                        logging.info(f"Registered font: {font_key}")
                    except Exception as e:
                        logging.error(f"Error registering font {font_key}: {str(e)}")
            logging.info(f"Registered {font_name} font family")

class PerformanceTimer:
    def __init__(self):
        self.start_time = None
        self.timings = {}

    def start(self, operation: str):
        self.start_time = time.time()
        self.timings[operation] = []

    def stop(self, operation: str):
        if self.start_time is not None:
            duration = time.time() - self.start_time
            self.timings[operation].append(duration)
            self.start_time = None

    def get_average(self, operation: str) -> float:
        times = self.timings.get(operation, [])
        return sum(times) / len(times) if times else 0

class HTMLToPDFConverter:
    def __init__(self, options: PDFOptions):
        self.options = options
        self.timer = PerformanceTimer()
        FontManager.register_fonts()
        self.font_name = f"{self.options.font_name}-normal"
        logging.info(f"Using font: {self.options.font_name}")

    def parse_html(self, html_content: str):
        """Parse HTML content using selectolax."""
        try:
            self.timer.start('html_parsing')
            document = HTMLParser(html_content)
            self.timer.stop('html_parsing')
            return document
        except Exception as e:
            logging.error(f"Error parsing HTML: {str(e)}")
            raise

    def process_table(self, table_node) -> Tuple[List[List[str]], List[List[Dict]]]:
        """Process a table node and return data and style."""
        rows = []
        styles = []
        max_cols = 0
        
        # First pass: calculate max columns
        for row in table_node.css('tr'):
            col_count = 0
            for cell in row.css('td, th'):
                colspan = int(cell.attributes.get('colspan', 1))
                col_count += colspan
            max_cols = max(max_cols, col_count)
        
        # Process each row
        for row in table_node.css('tr'):
            row_data = []
            row_styles = []
            current_col = 0
            
            # Process each cell
            for cell in row.css('td, th'):
                # Get cell content
                text = cell.text().strip() if cell.text() else ''
                
                # Get cell style
                style = {
                    'align': 'left',  # Default left alignment
                    'padding': (3, 3, 3, 3),  # Smaller padding
                    'background': None,
                    'textColor': None,
                    'colspan': 1,  # Default colspan
                    'fontName': self.font_name,
                    'fontSize': self.options.font_size
                }
                
                # Check for colspan
                if 'colspan' in cell.attributes:
                    try:
                        style['colspan'] = int(cell.attributes['colspan'])
                    except (ValueError, TypeError):
                        style['colspan'] = 1
                
                # Check for alignment
                if 'text-align' in cell.attributes.get('style', ''):
                    align = cell.attributes['style'].split('text-align:')[1].split(';')[0].strip()
                    style['align'] = align
                
                # Add cell content
                row_data.append(text)
                row_styles.append(style)
                current_col += style['colspan']
                
                # Add empty cells for colspan
                for _ in range(style['colspan'] - 1):
                    row_data.append('')
                    empty_style = style.copy()
                    empty_style['colspan'] = 1
                    row_styles.append(empty_style)
            
            # Fill remaining columns with empty cells
            while current_col < max_cols:
                row_data.append('')
                row_styles.append({
                    'align': 'left',
                    'padding': (3, 3, 3, 3),
                    'background': None,
                    'textColor': None,
                    'colspan': 1,
                    'fontName': self.font_name,
                    'fontSize': self.options.font_size
                })
                current_col += 1
            
            if row_data:  # Only add non-empty rows
                rows.append(row_data)
                styles.append(row_styles)
        
        return rows, styles

    def create_pdf_elements(self, document):
        """Create PDF elements from HTML document."""
        elements = []
        total_width = 0
        
        # Find all tables in the document
        tables = document.css('table')
        
        for table in tables:
            # Check if table has border
            table_border = int(table.attributes.get('border', 0))
            
            # Get all rows including header and body
            rows = table.css('tr')
            if not rows:
                continue
            
            # First pass: calculate max columns and track spans
            max_cols = 0
            row_spans = {}  # Track active rowspans
            
            for row_idx, row in enumerate(rows):
                col_idx = 0
                while col_idx in row_spans.get(row_idx, {}):
                    col_idx += 1
                    
                for cell in row.css('td, th'):
                    colspan = int(cell.attributes.get('colspan', 1))
                    rowspan = int(cell.attributes.get('rowspan', 1))
                    
                    # Track rowspans
                    if rowspan > 1:
                        for r in range(row_idx + 1, min(row_idx + rowspan, len(rows))):
                            if r not in row_spans:
                                row_spans[r] = {}
                            for c in range(col_idx, col_idx + colspan):
                                row_spans[r][c] = True
                    
                    col_idx += colspan
                max_cols = max(max_cols, col_idx)
            
            # Initialize table data and styles
            table_data = []
            style_commands = [
                ('FONT', (0, 0), (-1, -1), self.font_name, self.options.font_size),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 2),
                ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            ]

            # Add border only if specified in table attributes
            if table_border > 0:
                style_commands.append(('GRID', (0, 0), (-1, -1), 0.25, colors.black))
            
            # Second pass: fill table data
            row_spans = {}  # Reset rowspans tracking
            for row_idx, row in enumerate(rows):
                row_data = [''] * max_cols
                col_idx = 0
                
                # Skip columns that are part of active rowspans
                while col_idx in row_spans.get(row_idx, {}):
                    col_idx += 1
                
                for cell in row.css('td, th'):
                    # Skip if column is part of a rowspan
                    while col_idx in row_spans.get(row_idx, {}):
                        col_idx += 1
                    
                    if col_idx >= max_cols:
                        break
                    
                    # Get cell content and spans
                    text = cell.text().strip() if cell.text() else ''
                    colspan = int(cell.attributes.get('colspan', 1))
                    rowspan = int(cell.attributes.get('rowspan', 1))
                    
                    # Add cell content
                    row_data[col_idx] = text
                    
                    # Handle spans
                    if colspan > 1 or rowspan > 1:
                        end_col = min(col_idx + colspan - 1, max_cols - 1)
                        end_row = row_idx + rowspan - 1
                        style_commands.append(('SPAN', (col_idx, row_idx), (end_col, end_row)))
                        
                        # Track rowspans
                        if rowspan > 1:
                            for r in range(row_idx + 1, min(row_idx + rowspan, len(rows))):
                                if r not in row_spans:
                                    row_spans[r] = {}
                                for c in range(col_idx, col_idx + colspan):
                                    row_spans[r][c] = True
                    
                    # Handle cell styles
                    if 'style' in cell.attributes:
                        cell_style = cell.attributes['style']
                        if 'text-align' in cell_style:
                            align = cell_style.split('text-align:')[1].split(';')[0].strip()
                            style_commands.append(('ALIGN', (col_idx, row_idx), 
                                                (col_idx + colspan - 1, row_idx + rowspan - 1), 
                                                align.upper()))
                        if 'background-color' in cell_style:
                            bg_color = cell_style.split('background-color:')[1].split(';')[0].strip()
                            if bg_color.startswith('#'):
                                bg_color = bg_color.lstrip('#')
                                bg_color = tuple(int(bg_color[i:i+2], 16)/255 for i in (0, 2, 4))
                                style_commands.append(('BACKGROUND', (col_idx, row_idx), 
                                                    (col_idx + colspan - 1, row_idx + rowspan - 1), 
                                                    bg_color))
                        if 'border' in cell_style:
                            style_commands.append(('BOX', (col_idx, row_idx),
                                                (col_idx + colspan - 1, row_idx + rowspan - 1),
                                                0.25, colors.black))
                    
                    col_idx += colspan
                
                table_data.append(row_data)
            
            # Calculate column widths
            col_widths = []
            table_width = 0
            
            # First pass: check for width attributes in cells
            width_constraints = [None] * max_cols
            for row_idx, row in enumerate(rows):
                col_idx = 0
                for cell in row.css('td, th'):
                    if col_idx >= max_cols:
                        break
                    
                    # Get width from style attribute or width attribute
                    width_str = None
                    if 'style' in cell.attributes:
                        style = cell.attributes['style']
                        if 'width' in style:
                            width_str = style.split('width:')[1].split(';')[0].strip()
                    elif 'width' in cell.attributes:
                        width_str = cell.attributes['width']
                    
                    if width_str:
                        # Convert percentage to actual width
                        if '%' in width_str:
                            percentage = float(width_str.replace('%', '')) / 100.0
                            width_constraints[col_idx] = percentage * (A4[0] - 2*cm)  # A4 width minus margins
                        # Convert pixel values (approximate conversion)
                        elif 'px' in width_str:
                            pixels = float(width_str.replace('px', ''))
                            width_constraints[col_idx] = pixels * 0.0352778 * cm  # Convert px to cm
                    
                    colspan = int(cell.attributes.get('colspan', 1))
                    col_idx += colspan
            
            # Second pass: calculate final widths
            for col in range(max_cols):
                max_width = 0
                for row in table_data:
                    content = str(row[col])
                    content_width = len(content) * self.options.font_size * 0.5
                    max_width = max(max_width, content_width)
                
                # Use width constraint if available, otherwise use content-based width
                if width_constraints[col] is not None:
                    col_width = width_constraints[col]
                else:
                    col_width = max(max_width + 4, 0.8 * cm)
                
                col_widths.append(col_width)
                table_width += col_width
            
            total_width = max(total_width, table_width)
            
            # Create table
            # Calculate row heights based on content
            row_heights = []
            for row in table_data:
                max_height = self.options.font_size  # minimum height
                for cell_content in row:
                    # Calculate wrapped text height
                    lines = len(str(cell_content).split('\n'))
                    cell_height = lines * (self.options.font_size + 2)  # font size + padding
                    max_height = max(max_height, cell_height)
                row_heights.append(max_height)

            # Check for page break attributes
            for row_idx, row in enumerate(rows):
                if 'style' in row.attributes:
                    style = row.attributes['style']
                    if 'page-break-before' in style or 'page-break-after' in style:
                        elements.append(PageBreak())
                        break

            # Check if table has a page-break-after style
            if 'style' in table.attributes and 'page-break-after' in table.attributes['style']:
                add_page_break_after = True
            else:
                add_page_break_after = False

            table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)
            table.setStyle(TableStyle(style_commands))
            elements.append(table)

            # Add page break if needed
            if add_page_break_after:
                elements.append(PageBreak())
        
        self.max_content_width = total_width
        return elements

    def convert(self, html_content: str, output_path: str) -> None:
        """Convert HTML content to PDF."""
        try:
            # Parse HTML
            document = self.parse_html(html_content)
            
            # Create PDF elements
            elements = self.create_pdf_elements(document)
            
            # Determine page size based on content width
            if self.max_content_width > A4[0] - 2*cm:  # If wider than A4 width minus margins
                # Use landscape with custom width
                page_width = self.max_content_width + 2*cm  # Add margins
                page_height = A4[1]  # Keep A4 height
                pagesize = (page_width, page_height)
            else:
                pagesize = A4
            
            # Create PDF
            self.timer.start('pdf_rendering')
            doc = SimpleDocTemplate(
                output_path,
                pagesize=pagesize,
                rightMargin=self.options.margin_right * cm,
                leftMargin=self.options.margin_left * cm,
                topMargin=self.options.margin_top * cm,
                bottomMargin=self.options.margin_bottom * cm
            )
            
            # Build PDF
            doc.build(
                elements,
                onFirstPage=self._header_footer,
                onLaterPages=self._header_footer
            )
            self.timer.stop('pdf_rendering')
            
            # Log performance metrics
            logging.info(f"HTML parsing took {self.timer.get_average('html_parsing'):.3f} seconds")
            logging.info(f"PDF rendering took {self.timer.get_average('pdf_rendering'):.3f} seconds")
            
        except Exception as e:
            logging.error(f"Error converting HTML to PDF: {str(e)}\nTraceback: {sys.exc_info()[2]}")
            raise

    def _header_footer(self, canvas, doc):
        """Add header and footer to each page."""
        canvas.saveState()
        
        # Add page number
        canvas.setFont(self.font_name, 8)
        page_num = f"Page {doc.page} of {doc.page}"
        canvas.drawRightString(doc.pagesize[0] - doc.rightMargin, doc.bottomMargin - 20, page_num)
        
        canvas.restoreState()

if __name__ == "__main__":
    # Configure logging
    logging.basicConfig(level=logging.INFO)
    
    # Read input from stdin
    input_data = json.loads(sys.stdin.read())
    html_content = input_data.get("html")
    output_path = input_data.get("output_path")
    options = PDFOptions(**input_data.get("options", {}))
    
    # Convert HTML to PDF
    converter = HTMLToPDFConverter(options)
    converter.convert(html_content, output_path)
    
    # Return success response
    print(json.dumps({"success": True})) 