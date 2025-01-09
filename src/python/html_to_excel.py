import logging
import sys
import json
import base64
from io import BytesIO
from converter import HTMLToExcelConverter

# Set up logging to stderr only
logging.basicConfig(level=logging.DEBUG, 
                   format='%(asctime)s - %(levelname)s - %(message)s',
                   stream=sys.stderr)

def convert_html_to_excel_buffer(html_file_path):
    try:
        logging.info(f"Received file path: {html_file_path}")
        logging.info(f"Opening file: {html_file_path}")
        
        with open(html_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            logging.info(f"Successfully read HTML file, size: {len(html_content)} bytes")
        
        logging.info("Created BytesIO buffer")
        excel_buffer = BytesIO()
        
        logging.info("Creating converter instance")
        converter = HTMLToExcelConverter()
        logging.info("Initialized HTMLToExcelConverter")
        
        logging.info("Starting HTML to Excel conversion")
        converter.convert(html_content, excel_buffer)
        
        logging.info("Conversion completed successfully")
        excel_data = excel_buffer.getvalue()
        excel_base64 = base64.b64encode(excel_data).decode('utf-8')
        logging.info("Successfully encoded Excel data to base64")
        
        # Write result to stdout
        result = {
            "success": True,
            "data": excel_base64
        }
        sys.stdout.write(json.dumps(result))
        sys.stdout.flush()
        return 0
        
    except Exception as e:
        logging.error(f"Error during conversion: {str(e)}", exc_info=True)
        # Write error to stdout
        result = {
            "success": False,
            "error": str(e)
        }
        sys.stdout.write(json.dumps(result))
        sys.stdout.flush()
        return 1

if __name__ == "__main__":
    if len(sys.argv) != 2:
        result = {
            "success": False,
            "error": "Missing HTML file path argument"
        }
        sys.stdout.write(json.dumps(result))
        sys.stdout.flush()
        sys.exit(1)
    
    html_file_path = sys.argv[1]
    sys.exit(convert_html_to_excel_buffer(html_file_path)) 