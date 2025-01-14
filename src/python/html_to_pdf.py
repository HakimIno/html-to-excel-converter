import sys
import json
import logging
from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration
from pathlib import Path
import time
from dataclasses import dataclass
from typing import Optional, Dict, Union, List
from io import BytesIO
import base64
import gc

# Configure logging
logging.basicConfig(level=logging.INFO)
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
        cls._timings.clear()

@dataclass
class PDFOptions:
    """Configuration options for PDF conversion"""
    page_size: str = 'A4'  # A4, Letter, etc.
    margin_top: str = '1cm'
    margin_right: str = '1cm'
    margin_bottom: str = '1cm'
    margin_left: str = '1cm'
    zoom: float = 1.0
    optimize_images: bool = True
    optimize_fonts: bool = True
    enable_javascript: bool = False
    cache_fonts: bool = True
    cache_images: bool = True
    dpi: int = 96
    base_url: Optional[str] = None
    custom_css: Optional[str] = None

class HTMLToPDFConverter:
    """High-performance HTML to PDF converter using WeasyPrint"""
    
    def __init__(self, options: Dict = None):
        self.options = PDFOptions()
        if options:
            for key, value in options.items():
                if hasattr(self.options, key):
                    setattr(self.options, key, value)
        
        # Initialize font configuration with caching
        self.font_config = FontConfiguration() if self.options.cache_fonts else None
        
        # Default CSS for better PDF rendering
        self.default_css = CSS(string='''
            @page {
                size: ''' + self.options.page_size + ''';
                margin: ''' + self.options.margin_top + ''' ''' + self.options.margin_right + ''' ''' + 
                self.options.margin_bottom + ''' ''' + self.options.margin_left + ''';
            }
            body { font-family: system-ui, -apple-system, sans-serif; }
            table { border-collapse: collapse; width: 100%; }
            td, th { padding: 8px; border: 1px solid #ddd; }
            th { background-color: #f5f5f5; }
        ''')
    
    def _cleanup(self):
        """Clean up resources"""
        gc.collect()
    
    def convert(self, html_content: str, output_path: Optional[str] = None) -> Dict[str, Union[bool, str, bytes]]:
        """Convert HTML to PDF with high performance optimizations"""
        try:
            with PerformanceTimer('Total conversion'):
                # Create WeasyPrint HTML object
                with PerformanceTimer('HTML parsing'):
                    html = HTML(
                        string=html_content,
                        base_url=self.options.base_url,
                        encoding='utf-8'
                    )
                
                # Prepare CSS
                stylesheets = [self.default_css]
                if self.options.custom_css:
                    stylesheets.append(CSS(string=self.options.custom_css))
                
                # Render PDF
                with PerformanceTimer('PDF rendering'):
                    pdf = html.write_pdf(
                        stylesheets=stylesheets,
                        font_config=self.font_config,
                        zoom=self.options.zoom,
                        dpi=self.options.dpi
                    )
                
                if output_path:
                    # Write to file
                    with PerformanceTimer('File writing'):
                        with open(output_path, 'wb') as f:
                            f.write(pdf)
                    return {"success": True, "data": ""}
                else:
                    # Return PDF data
                    return {
                        "success": True,
                        "data": base64.b64encode(pdf).decode('utf-8')
                    }
                
        except Exception as e:
            logger.error(f"Conversion failed: {str(e)}")
            return {"success": False, "error": str(e)}
            
        finally:
            self._cleanup()
            PerformanceTimer.print_summary()

if __name__ == "__main__":
    try:
        # Read input from stdin
        input_data = sys.stdin.read().strip()
        converter = HTMLToPDFConverter()
        
        try:
            data = json.loads(input_data)
            html_content = data.get('html', '')
            output_to_buffer = data.get('buffer', True) 
            output_file = data.get('output', None) if not output_to_buffer else None
            options = data.get('options', {})
            
            if options:
                converter = HTMLToPDFConverter(options)
                
        except json.JSONDecodeError:
            html_content = input_data
            output_file = None
            
        if not html_content:
            print(json.dumps({"error": "No HTML content provided"}), file=sys.stderr)
            sys.exit(1)
            
        result = converter.convert(html_content, output_file)
        print(json.dumps(result))
        sys.exit(0 if result["success"] else 1)
        
    except Exception as e:
        print(json.dumps({"error": str(e)}), file=sys.stderr)
        sys.exit(1) 