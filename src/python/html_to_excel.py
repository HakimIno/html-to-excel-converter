from converter import HTMLToExcelConverter
import sys
import json
from io import BytesIO
import base64

def convert_html_to_excel_buffer(html_content):
    try:
        # Create a BytesIO buffer
        buffer = BytesIO()
        
        # Convert HTML to Excel and write to buffer
        converter = HTMLToExcelConverter()
        converter.convert(html_content, buffer)
        
        # Get the buffer value and encode to base64
        excel_data = base64.b64encode(buffer.getvalue()).decode('utf-8')
        buffer.close()
        
        return json.dumps({
            'success': True,
            'data': excel_data
        })
    except Exception as e:
        return json.dumps({
            'success': False,
            'error': str(e)
        })

if __name__ == "__main__":
    # Read input from Node.js
    html_content = sys.argv[1]
    result = convert_html_to_excel_buffer(html_content)
    print(result) 