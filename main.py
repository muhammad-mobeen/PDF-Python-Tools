from fastapi import FastAPI, File, UploadFile, HTTPException, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
import uvicorn
from pdf2docx import Converter
import os
import tempfile
from fastapi.templating import Jinja2Templates

app = FastAPI()
templates = Jinja2Templates(directory="templates")

# Define input and output paths
path_input = 'input/'
path_output = 'output/'
project_dir = os.path.dirname(os.path.abspath(__file__))  # Get the absolute path of the project directory


def convert_pdf_to_docx(pdf_path, docx_path):
    try:
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error converting PDF to DOCX: {e}")


@app.get("/", response_class=HTMLResponse)
async def show_upload_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/convert")
async def convert_and_download(file: UploadFile = File(...)):
    # Check if the file is a PDF
    if not file.filename.lower().endswith('.pdf'):
        print("Uploaded file in wrong format:",file.filename)
        raise HTTPException(status_code=400, detail="Uploaded file must be a PDF.")

    # Ensure the input and output directories exist
    if not os.path.exists(path_input):
        os.makedirs(path_input)
    if not os.path.exists(path_output):
        os.makedirs(path_output)

    # Construct input and output file paths
    pdf_path = os.path.join(path_input, file.filename)
    docx_path = os.path.join(path_output, os.path.splitext(file.filename)[0] + '.docx')

    # Save the uploaded PDF to the input directory
    with open(pdf_path, 'wb') as pdf_file:
        pdf_file.write(file.file.read())

    # Convert the PDF to DOCX
    convert_pdf_to_docx(pdf_path, docx_path)

    # Provide the converted DOCX file for download
    response = FileResponse(docx_path, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    response.headers["Content-Disposition"] = f"attachment; filename={file.filename.replace('.pdf', '.docx')}"

    return response


if __name__ == '__main__':
    uvicorn.run(app)