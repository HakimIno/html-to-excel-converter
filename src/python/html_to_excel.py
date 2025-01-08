from converter import HTMLToExcelConverter
import sys
import json
from io import BytesIO
import base64

def convert_html_to_excel_buffer(html_file_path):
    try:
        # อ่าน HTML จาก file
        with open(html_file_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            
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
    if len(sys.argv) < 2:
        print(json.dumps({
            'success': False,
            'error': 'HTML file path is required'
        }))
        sys.exit(1)
        
    html_file_path = sys.argv[1]
    result = convert_html_to_excel_buffer(html_file_path)
    print(result) 