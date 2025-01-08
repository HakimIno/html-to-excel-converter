# main.py
from fastapi import FastAPI, HTTPException, UploadFile, File, Response
from fastapi.responses import FileResponse
from pydantic import BaseModel
import uvicorn
from typing import Optional
import tempfile
import os
from datetime import datetime
from io import BytesIO

# Import converter class
from converter import HTMLToExcelConverter

app = FastAPI(title="HTML to Excel Converter API")

class HTMLInput(BaseModel):
    html_content: str
    filename: Optional[str] = None

@app.post("/convert")
async def convert_html_to_excel(input_data: HTMLInput):
    try:
        # Create temp directory if it doesn't exist
        temp_dir = "temp"
        os.makedirs(temp_dir, exist_ok=True)

        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = input_data.filename or f"converted_{timestamp}.xlsx"
        output_path = os.path.join(temp_dir, output_filename)

        # Convert HTML to Excel
        converter = HTMLToExcelConverter()
        converter.convert(input_data.html_content, output_path)

        # Return the file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/html-to-excel")
async def convert_html_file_to_excel(
    file: UploadFile = File(...),
    use_custom_config: bool = False
):
    try:
        # Create temp directory if it doesn't exist
        temp_dir = "temp"
        os.makedirs(temp_dir, exist_ok=True)

        # Read the uploaded HTML file
        html_content = await file.read()
        html_content = html_content.decode()

        # Generate output filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"converted_{timestamp}.xlsx"
        output_path = os.path.join(temp_dir, output_filename)

        # Convert HTML to Excel
        converter = HTMLToExcelConverter()
        converter.convert(html_content, output_path)

        # Return the file
        return FileResponse(
            path=output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/convert/buffer")
async def convert_html_to_excel_buffer(input_data: HTMLInput):
    try:
        # Create a BytesIO buffer
        buffer = BytesIO()
        
        # Convert HTML to Excel and write to buffer
        converter = HTMLToExcelConverter()
        converter.convert(input_data.html_content, buffer)
        
        # Get the buffer value
        excel_data = buffer.getvalue()
        buffer.close()
        
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = input_data.filename or f"converted_{timestamp}.xlsx"
        
        # Return the buffer as response
        return Response(
            content=excel_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                'Content-Disposition': f'attachment; filename="{filename}"'
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)