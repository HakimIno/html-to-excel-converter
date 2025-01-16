import sys
import json
from python.converter import HTMLTableConverter

def main():
    try:
        if len(sys.argv) < 3:
            print(json.dumps({"success": False, "error": "Missing input or output file path"}))
            sys.exit(1)

        input_file = sys.argv[1]
        output_file = sys.argv[2]

        # Read HTML content
        with open(input_file, 'r', encoding='utf-8') as f:
            html_content = f.read()

        # Convert to Excel
        converter = HTMLTableConverter()
        result = converter.convert(html_content, output_file)

        # Return result
        print(json.dumps(result))
        sys.exit(0 if result["success"] else 1)

    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}))
        sys.exit(1)

if __name__ == "__main__":
    main() 