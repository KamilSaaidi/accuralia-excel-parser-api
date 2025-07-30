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
        
        # Détecter le type de fichier basé sur l'extension
        file_extension = filename.lower().split('.')[-1]
        
        print(f"Processing file: {filename}, extension: {file_extension}")
        
        if file_extension in ['xlsx', 'xls']:
            return await process_excel_file(file_bytes, filename, file_extension)
        elif file_extension == 'csv':
            return await process_csv_file(file_bytes, filename)
        else:
            return {"success": False, "error": f"Format de fichier non supporté: {file_extension}"}
            
    except Exception as e:
        print(f"Error in parse_excel: {str(e)}")
        return {"success": False, "error": str(e), "method": "EXCEL_PARSER_ERROR"}

async def process_excel_file(file_bytes: bytes, filename: str, file_type: str) -> Dict[str, Any]:
    """
    Traite un fichier Excel et extrait le contenu
    """
    try:
        print(f"Processing Excel file: {filename}, declared type: {file_type}")
        
        # Détecter le vrai type de fichier basé sur la signature binaire
        file_signature = file_bytes[:8]
        signature_hex = ''.join(f'{b:02x}' for b in file_signature)
        print(f"File signature: {signature_hex}")
        
        # Déterminer l'engine basé sur la signature réelle
        if signature_hex.startswith('504b0304') or signature_hex.startswith('504b0506'):
            real_type = 'xlsx'
            engine = 'openpyxl'
            print(f"Detected real type: XLSX (Office 2007+)")
        elif signature_hex.startswith('d0cf11e0'):
            real_type = 'xls'
            engine = 'xlrd'
            print(f"Detected real type: XLS (Office 97-2003)")
        else:
            return {"success": False, "error": f"Signature de fichier non reconnue: {signature_hex}"}
        
        print(f"Using engine: {engine} for real type: {real_type}")
        
        # Créer un buffer en mémoire
        file_buffer = io.BytesIO(file_bytes)
        
        # Lire le fichier Excel avec le bon engine
        excel_file = pd.ExcelFile(file_buffer, engine=engine)
        
        # Extraire le contenu de toutes les feuilles
        all_sheets_data = {}
        extracted_text = ""
        
        for sheet_name in excel_file.sheet_names:
            try:
                print(f"Processing sheet: {sheet_name}")
                
                # Lire la feuille avec le bon engine
                df = pd.read_excel(file_buffer, sheet_name=sheet_name, engine=engine)
                
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
                
                print(f"Successfully processed sheet: {sheet_name}")
                
            except Exception as sheet_error:
                print(f"Erreur lecture feuille {sheet_name}: {sheet_error}")
                continue
        
        # Nettoyer le texte extrait
        extracted_text = extracted_text.strip()
        
        print(f"Extracted text length: {len(extracted_text)}")
        
        return {
            "success": True,
            "text": extracted_text,
            "method": f"EXCEL_PARSING_{real_type.upper()}",
            "confidence": 0.9,
            "filename": filename,
            "file_type": real_type,
            "sheets_count": len(excel_file.sheet_names),
            "sheets_data": all_sheets_data,
            "textLength": len(extracted_text)
        }
        
    except Exception as e:
        print(f"Error in process_excel_file: {str(e)}")
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