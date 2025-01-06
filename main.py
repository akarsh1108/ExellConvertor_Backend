import io
from turtle import pd
import chardet
from fastapi import FastAPI, File, HTTPException, UploadFile, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from tempfile import NamedTemporaryFile
import os
from fpdf import FPDF
import pandas as pd

from accident import accident_process_excel
from advance import advance_process_excel
from bonusFromC import bonus_process_excel
from damage import damage_process_excel
from esicpf import esicpf_process_excel
from fine import fine_process_excel
from formD import formD_process_excel
from muster import muster_process_excel
from overtime import overtime_process_excel
from pdfHighlighter import extract_identifiers_from_excel, highlight_identifiers_in_pdf
from wagesRegister import wages_process_excel
from workmen import workmen_process_excel
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from io import BytesIO

 # Assuming you have added damage_process_excel

# FastAPI instance
app = FastAPI()

# Configure CORS settings
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],  # List of allowed origins (e.g., React frontend)
    allow_credentials=True,                  # Allow cookies and authentication headers
    allow_methods=["*"],                     # Allow all HTTP methods
    allow_headers=["*"],                     # Allow all headers
)

# File paths
INPUT_FOLDER = "input"
OUTPUT_FOLDER = "output"
ACCIDENT_FILE_PATH = os.path.join(INPUT_FOLDER, "accident.xlsx")
ADVANCE_FILE_PATH = os.path.join(INPUT_FOLDER, "advance.xlsx")
BONUSFORMC_FILE_PATH = os.path.join(INPUT_FOLDER, "bonusformc.xlsx")
DAMAGE_FILE_PATH = os.path.join(INPUT_FOLDER, "damage.xlsx")  # Add path for damage.xlsx
ESP_FILE_PATH = os.path.join(INPUT_FOLDER, "esicpf.xlsx")
FINE_FILE_PATH = os.path.join(INPUT_FOLDER, "fine.xlsx")
FORMD_FILE_PATH = os.path.join(INPUT_FOLDER, "formD.xlsx")
MUSTER_FILE_PATH = os.path.join(INPUT_FOLDER, "muster.xlsx")
OVERTIME_FILE_PATH = os.path.join(INPUT_FOLDER, "overtime.xlsx")
WAGES_FILE_PATH = os.path.join(INPUT_FOLDER, "wages.xlsx")
WORKMEN_FILE_PATH = os.path.join(INPUT_FOLDER, "workmen.xlsx")

from fastapi import FastAPI, Form, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from tempfile import NamedTemporaryFile
import os
import zipfile
from typing import List

app = FastAPI()

# Configure CORS settings
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],  
)

@app.post("/process-all-excel/")
async def process_excel_endpoint(
    process_name: str = Form(...),  # Accept process name as form data
    master_file: UploadFile = File(...)  # Accept master file as a file upload
):
    """API to process Excel files with a specified process."""
    file_path_mapping = {
        "accident": ACCIDENT_FILE_PATH,
        "advance": ADVANCE_FILE_PATH,
        "bonusformc": BONUSFORMC_FILE_PATH,
        "damage": DAMAGE_FILE_PATH,
        "esicpf": ESP_FILE_PATH,
        "fine": FINE_FILE_PATH,
        "formd": FORMD_FILE_PATH,
        "muster": MUSTER_FILE_PATH,
        "overtime": OVERTIME_FILE_PATH,
        "wages": WAGES_FILE_PATH,
        "workmen": WORKMEN_FILE_PATH,
        "all": None,  # Special key to process all
    }

    # Validate process name
    process_name = process_name.lower()
    if process_name not in file_path_mapping:
        return JSONResponse({"error": f"Unsupported process name: {process_name}"}, status_code=400)

    processed_files = []

    try:
        # Save the uploaded master file to a temporary location
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_master_file:
            temp_master_file.write(await master_file.read())
            temp_master_file_path = temp_master_file.name

            # Determine processes to run based on process_name
            if process_name == "all":
                processes_to_run = file_path_mapping.keys() - {"all"}  # Exclude "all" itself
            else:
                processes_to_run = [process_name]

            # Loop through the selected processes
            for key in processes_to_run:
                base_file_path = file_path_mapping[key]

                # Validate base file existence for each process
                if not os.path.exists(base_file_path):
                    return JSONResponse({"error": f"{base_file_path} not found in the input folder."}, status_code=400)

                with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_output_file:
                    temp_output_file_path = temp_output_file.name

                    # Call the appropriate processing function
                    if key == "accident":
                        output_file_path = accident_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "advance":
                        output_file_path = advance_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "bonusformc":
                        output_file_path = bonus_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "damage":
                        output_file_path = damage_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "esicpf":
                        output_file_path = esicpf_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "fine":
                        output_file_path = fine_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "formd":
                        output_file_path = formD_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "muster":
                        output_file_path = muster_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "overtime":
                        output_file_path = overtime_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "wages":
                        output_file_path = wages_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                    elif key == "workmen":
                        output_file_path = workmen_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)

                    # Add the output file to the list
                    processed_files.append((key, output_file_path))

        # Create a ZIP file containing all processed files
        with NamedTemporaryFile(delete=False, suffix=".zip") as temp_zip_file:
            with zipfile.ZipFile(temp_zip_file.name, 'w') as zipf:
                for key, file_path in processed_files:
                    zipf.write(file_path, arcname=f"{key}_updated.xlsx")

            return FileResponse(temp_zip_file.name, filename="processed_files.zip")

    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=500)


@app.post("/process-excel/")
async def process_excel_endpoint(
    process_name: str = Form(...),  # Accept process name as form data
    master_file: UploadFile = File(...)  # Accept master file as a file upload
):
    """API to process Excel files with a specified process."""
    file_path_mapping = {
        "accident": ACCIDENT_FILE_PATH,
        "advance": ADVANCE_FILE_PATH,
        "bonusformc": BONUSFORMC_FILE_PATH,
        "damage": DAMAGE_FILE_PATH,  # Add damage to the mapping
        "esicpf": ESP_FILE_PATH,
        "fine": FINE_FILE_PATH,
        "formd": FORMD_FILE_PATH,
        "muster": MUSTER_FILE_PATH,
        "overtime": OVERTIME_FILE_PATH,
        "wages": WAGES_FILE_PATH,
        "workmen": WORKMEN_FILE_PATH,

    }

    # Validate process name
    process_name = process_name.lower()
    if process_name not in file_path_mapping:
        return {"error": f"Unsupported process name: {process_name}"}

    # Validate base file existence
    base_file_path = file_path_mapping[process_name]
    if not os.path.exists(base_file_path):
        return {"error": f"{base_file_path} not found in the input folder."}

    try:
        # Save the uploaded master file to a temporary file
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_master_file, NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_output_file:
            temp_master_file.write(await master_file.read())
            temp_master_file_path = temp_master_file.name
            temp_output_file_path = temp_output_file.name

            # Process the files based on process_name
            if process_name == "accident":
                output_file_path = accident_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "advance":
                output_file_path = advance_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "bonusformc":
                output_file_path = bonus_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "damage":
                output_file_path = damage_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "esicpf":
                output_file_path = esicpf_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "fine":
                 output_file_path = fine_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "formd":
                output_file_path = formD_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "muster":
                output_file_path = muster_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "overtime":
                output_file_path = overtime_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "wages":
                output_file_path = wages_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
            elif process_name == "workmen":
                output_file_path = workmen_process_excel(base_file_path, temp_master_file_path, temp_output_file_path)
                
            # Return the processed file
            return FileResponse(output_file_path, filename=f"{process_name}_updated.xlsx")

    except Exception as e:
        return {"error": str(e)}

@app.post("/convert-to-pdf/")
async def convert_xlsx_to_pdf(xlsx_file: UploadFile = File(...), process_name: str = Form(...)):
    output_filename = None  # Ensure variable is defined before try block
    try:
        # Validate file extension
        if not xlsx_file.filename.endswith(('.xlsx', '.xlsm')):
            raise HTTPException(status_code=400, detail="Uploaded file is not a valid Excel file.")

        # Read the uploaded Excel file
        contents = await xlsx_file.read()
        excel_data = BytesIO(contents)

        # Attempt to load the workbook
        try:
            workbook = load_workbook(excel_data)
        except Exception as e:
            raise HTTPException(status_code=400, detail="Failed to read Excel file. Ensure the file is valid and not corrupted.")

        # Extract data from the workbook
        sheet = workbook.active
        data = sheet.values
        columns = next(data)  # Extract header row
        df = pd.DataFrame(data, columns=columns)

        # Define output PDF filename
        output_filename = f"{process_name}.pdf"

        # Create a PDF with matplotlib
        with PdfPages(output_filename) as pdf:
            fig, ax = plt.subplots(figsize=(8.5, 11))  # Letter size dimensions
            ax.axis('tight')
            ax.axis('off')

            # Create a table plot
            table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.auto_set_column_width(col=list(range(len(df.columns))))

            # Add the table to the PDF
            pdf.savefig(fig, bbox_inches='tight')
            plt.close(fig)

        return FileResponse(output_filename, media_type="application/pdf", filename=output_filename)

    except HTTPException as http_ex:
        raise http_ex
    except Exception as e:
        return {"error": str(e)}
    finally:
        # Clean up generated file
        if output_filename and os.path.exists(output_filename):
            os.remove(output_filename)


@app.post("/highlight-pdf/")
async def highlight_pdf_endpoint(xlsx_file: UploadFile = File(...), pdf_file: UploadFile = File(...)):
    """
    API endpoint to process an Excel file and PDF file, highlight identifiers in the PDF,
    and return the highlighted PDF as a response.

    Args:
        xlsx_file (UploadFile): The Excel file containing identifiers.
        pdf_file (UploadFile): The PDF file to be highlighted.

    Returns:
        FileResponse: The highlighted PDF file.
    """
    try:
        # Save the uploaded files temporarily
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_xlsx, NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_xlsx.write(await xlsx_file.read())
            temp_pdf.write(await pdf_file.read())
            temp_xlsx_path = temp_xlsx.name
            temp_pdf_path = temp_pdf.name

        # Extract identifiers from the Excel file
        identifiers = extract_identifiers_from_excel(temp_xlsx_path)

        # Define output file path for the highlighted PDF
        with NamedTemporaryFile(delete=False, suffix=".pdf") as temp_output_pdf:
            output_pdf_path = temp_output_pdf.name

        # Highlight the identifiers in the PDF
        result_pdf_path = highlight_identifiers_in_pdf(
            input_pdf_path=temp_pdf_path,
            output_pdf_path=output_pdf_path,
            identifiers=identifiers,
        )

        # Return the highlighted PDF
        return FileResponse(result_pdf_path, filename="highlighted_result.pdf", media_type="application/pdf")

    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        return {"error": str(e)}
