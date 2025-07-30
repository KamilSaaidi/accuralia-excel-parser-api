from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import base64
import io
import json
import csv
from typing import Dict, Any, List

app = FastAPI(title="Excel/CSV Parser API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "Excel/CSV Parser API", "version": "1.0.0", "status": "running"}

@app.post("/parse-excel")
async def parse_excel(fileData: str = Form(...), filename: str = Form(...)):
    """
    Parse un fichier Excel (XLSX, XLS) et extrait le contenu des cellules
    """
    try:
        # Décoder le fichier base64
        file_bytes = base64.b64decode(fileData)
        
        # Détecter le type de fichier
        file_extension = filename.lower().split('.')[-1]
        
        if file_extension in ['xlsx', 'xls']:
            return await process_excel_file(file_bytes, filename, file_extension)
        elif file_extension == 'csv':
            return await process_csv_file(file_bytes, filename)
        else:
            return {"success": False, "error": f"Format de fichier non supporté: {file_extension}"}
            
    except Exception as e:
        return {"success": False, "error": str(e), "method": "EXCEL_PARSER_ERROR"}

async def process_excel_file(file_bytes: bytes, filename: str, file_type: str) -> Dict[str, Any]:
    """
    Traite un fichier Excel et extrait le contenu
    """
    try:
        # Créer un buffer en mémoire
        file_buffer = io.BytesIO(file_bytes)
        
        # Lire le fichier Excel selon son type
        if file_type == 'xlsx':
            # Utiliser openpyxl pour XLSX
            excel_file = pd.ExcelFile(file_buffer, engine='openpyxl')
        else:
            # Utiliser xlrd pour XLS
            excel_file = pd.ExcelFile(file_buffer, engine='xlrd')
        
        # Extraire le contenu de toutes les feuilles
        all_sheets_data = {}
        extracted_text = ""
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Lire la feuille
                df = pd.read_excel(file_buffer, sheet_name=sheet_name, engine='openpyxl' if file_type == 'xlsx' else 'xlrd')
                
                # Convertir en texte
                sheet_text = df.to_string(index=False, header=True)
                extracted_text += f"\n=== Feuille: {sheet_name} ===\n"
                extracted_text += sheet_text + "\n"
                
                # Stocker les données structurées
                all_sheets_data[sheet_name] = {
                    "headers": df.columns.tolist() if not df.empty else [],
                    "rows": df.values.tolist() if not df.empty else [],
                    "shape": df.shape
                }
                
            except Exception as sheet_error:
                print(f"Erreur lecture feuille {sheet_name}: {sheet_error}")
                continue
        
        # Nettoyer le texte extrait
        extracted_text = extracted_text.strip()
        
        return {
            "success": True,
            "text": extracted_text,
            "method": f"EXCEL_PARSING_{file_type.upper()}",
            "confidence": 0.9,
            "filename": filename,
            "file_type": file_type,
            "sheets_count": len(excel_file.sheet_names),
            "sheets_data": all_sheets_data,
            "textLength": len(extracted_text)
        }
        
    except Exception as e:
        return {"success": False, "error": f"Erreur traitement Excel: {str(e)}", "method": "EXCEL_PARSER_ERROR"}

async def process_csv_file(file_bytes: bytes, filename: str) -> Dict[str, Any]:
    """
    Traite un fichier CSV et extrait le contenu
    """
    try:
        # Décoder le contenu CSV
        csv_content = file_bytes.decode('utf-8')
        
        # Lire avec pandas
        df = pd.read_csv(io.StringIO(csv_content))
        
        # Convertir en texte
        extracted_text = df.to_string(index=False, header=True)
        
        return {
            "success": True,
            "text": extracted_text,
            "method": "CSV_PARSING",
            "confidence": 0.9,
            "filename": filename,
            "file_type": "csv",
            "headers": df.columns.tolist() if not df.empty else [],
            "rows_count": len(df),
            "textLength": len(extracted_text)
        }
        
    except Exception as e:
        return {"success": False, "error": f"Erreur traitement CSV: {str(e)}", "method": "CSV_PARSER_ERROR"}

@app.post("/parse-base64")
async def parse_base64(fileData: str = Form(...), filename: str = Form(...)):
    """
    Endpoint alternatif pour la compatibilité
    """
    return await parse_excel(fileData, filename)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000) 