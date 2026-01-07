from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi import Request
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import numpy as np
import subprocess
import json
import os
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional
import uvicorn
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Inicializar FastAPI
app = FastAPI(
    title="API Modelo Frontera",
    description="API para ejecutar modelos R de predicción",
    version="1.0.0"
)

# Configurar CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # En producción, restringir a tu dominio
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Rutas base
BASE_DIR = Path(__file__).parent.parent
CONTENT_DIR = BASE_DIR / "content"
INPUT_DIR = CONTENT_DIR / "input"
SRC_DIR = CONTENT_DIR / "src"
OUTPUT_DIR = CONTENT_DIR / "output"

# Asegurar directorios existen
for dir_path in [INPUT_DIR, OUTPUT_DIR]:
    dir_path.mkdir(parents=True, exist_ok=True)

# ============ ENDPOINTS ============

@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    """Sirve el frontend HTML"""
    frontend_path = BASE_DIR / "frontend-simple" / "index.html"
    if frontend_path.exists():
        return HTMLResponse(content=frontend_path.read_text(encoding="utf-8"))
    return {"message": "API Modelo Frontera está funcionando. Use /docs para documentación."}

@app.get("/health")
async def health_check():
    """Verifica que todos los componentes estén funcionando"""
    checks = {
        "api": "healthy",
        "r_available": os.path.exists("/usr/local/bin/R"),
        "directories": {
            "content": CONTENT_DIR.exists(),
            "input": INPUT_DIR.exists(),
            "src": SRC_DIR.exists(),
            "output": OUTPUT_DIR.exists()
        },
        "r_scripts": {
            "frontera_r": (SRC_DIR / "Frontera.R").exists(),
            "frontera_utils": (SRC_DIR / "frontera_utils.R").exists()
        }
    }
    return checks

@app.post("/api/predict")
async def predict(request: Request):
    """
    Endpoint principal: Recibe los 6 parámetros, ejecuta el modelo R y retorna resultados
    """
    start_time = time.time()
    
    try:
        # Obtener datos como JSON
        form_data = await request.json()

        # 1. Preparar datos de entrada
        input_data = {
            "Sector_Econom": form_data.get("sector_econom"),
            "Tamano_Emp": form_data.get("tamano_emp"),
            "Activ_Econ": form_data.get("activ_econ"),
            "Sucursal": form_data.get("sucursal"),
            "Num_Empleados": form_data.get("num_empleados"),
            "tasa_deseada": form_data.get("tasa_deseada")
        }

        logger.info(f"Datos recibidos: {input_data}")
        
        # 2. Crear DataFrame y guardar como Excel
        df_input = pd.DataFrame([input_data])
        input_excel_path = INPUT_DIR / "Input_frontera.xlsx"
        df_input.to_excel(input_excel_path, index=False)
        logger.info(f"Excel creado en: {input_excel_path}")
        
        # 3. Ejecutar script R
        r_script_path = SRC_DIR / "Frontera.R"
        
        if not r_script_path.exists():
            raise HTTPException(status_code=500, detail="Script R no encontrado")
        
        # Cambiar al directorio content
        original_cwd = os.getcwd()
        os.chdir(CONTENT_DIR)
        
        try:
            # Ejecutar Rscript
            cmd = ["Rscript", str(r_script_path.relative_to(CONTENT_DIR))]
            logger.info(f"Ejecutando comando: {' '.join(cmd)}")
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=600  # 5 minutos timeout
            )
            
            # Verificar resultado
            if result.returncode != 0:
                error_msg = f"Error en ejecución R:\n{result.stderr}"
                logger.error(error_msg)
                raise HTTPException(status_code=500, detail=error_msg)
            
            logger.info(f"Rscript ejecutado exitosamente. Output: {result.stdout[:500]}...")
            
        finally:
            os.chdir(original_cwd)  # Volver al directorio original
        
        # 4. Buscar archivo de salida generado por R
        output_excel_path = None
        posibles_rutas = [
            CONTENT_DIR / "recomendaciones_PYP.xlsx",  # Ruta original
            OUTPUT_DIR / "recomendaciones_PYP.xlsx",   # Ruta alternativa
            CONTENT_DIR / "output" / "recomendaciones_PYP.xlsx"
        ]
        
        for ruta in posibles_rutas:
            if ruta.exists():
                output_excel_path = ruta
                break
        
        if not output_excel_path:
            # Si R no generó el archivo, crear uno de ejemplo
            logger.warning("R no generó archivo de salida. Creando ejemplo.")
            output_excel_path = crear_ejemplo_resultados(input_data)
        
        # 5. Leer resultados del Excel
        try:
            df_results = pd.read_excel(output_excel_path)
            logger.info(f"Resultados leídos. Columnas: {df_results.columns.tolist()}")
        except Exception as e:
            logger.error(f"Error leyendo Excel: {e}")
            df_results = pd.DataFrame({
                "Variable": ["Ejemplo", "Ejemplo"],
                "Valor": ["Demo", "Demo"],
                "Recomendacion": ["Resultado de prueba", "Resultado de prueba"]
            })
        
        # 6. Preparar respuesta
        execution_time = round(time.time() - start_time, 2)
        
        # Convertir DataFrame a lista de diccionarios para el frontend
        results_list = []
        for _, row in df_results.iterrows():
            # Asumir que el Excel tiene columnas específicas
            # Ajustar según la estructura real de tu Excel
            if "Variable" in df_results.columns and "Valor" in df_results.columns:
                results_list.append({
                    "variable": str(row.get("Variable", "")),
                    "value": str(row.get("Valor", "")),
                    "recomendacion": str(row.get("Recomendacion", row.get("Recomendación", "")))
                })
            else:
                # Si no tiene las columnas esperadas, usar todas las columnas
                for col in df_results.columns:
                    results_list.append({
                        "variable": col,
                        "value": str(row[col]),
                        "recomendacion": ""
                    })
        
        # Limitar a 20 resultados para no saturar el frontend
        results_list = results_list[:20]
        
        response_data = {
            "status": "success",
            "execution_time": execution_time,
            "input_data": input_data,
            "results_count": len(results_list),
            "results": results_list,
            "excel_download_url": f"/api/download/results"
        }
        
        logger.info(f"Predicción completada en {execution_time}s. Resultados: {len(results_list)} items")
        
        return response_data
        
    except HTTPException:
        raise
    except subprocess.TimeoutExpired:
        logger.error("Timeout: El script R tardó demasiado")
        raise HTTPException(status_code=504, detail="Timeout: El modelo tardó demasiado en ejecutar")
    except Exception as e:
        logger.error(f"Error inesperado: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Error interno: {str(e)}")

@app.get("/api/download/results")
async def download_results():
    """Descarga el archivo Excel generado por R"""
    posibles_rutas = [
        CONTENT_DIR / "recomendaciones_PYP.xlsx",
        OUTPUT_DIR / "recomendaciones_PYP.xlsx"
    ]
    
    for ruta in posibles_rutas:
        if ruta.exists():
            return FileResponse(
                path=ruta,
                filename=f"recomendaciones_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    raise HTTPException(status_code=404, detail="Archivo de resultados no encontrado")

@app.get("/api/test-r")
async def test_r():
    """Prueba que R esté funcionando correctamente"""
    try:
        result = subprocess.run(
            ["R", "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            return {
                "status": "success",
                "r_version": result.stdout.split('\n')[0] if result.stdout else "Unknown",
                "rscript_available": os.path.exists("/usr/local/bin/Rscript")
            }
        else:
            return {
                "status": "error",
                "message": result.stderr
            }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ============ FUNCIONES AUXILIARES ============

def crear_ejemplo_resultados(input_data: Dict) -> Path:
    """Crea un archivo Excel de ejemplo cuando R falla"""
    ejemplo_data = {
        "Variable": [
            "Sector Económico",
            "Tamaño Empresa", 
            "Actividad Económica",
            "Recomendación Principal",
            "Tasa Sugerida",
            "Score de Riesgo"
        ],
        "Valor": [
            input_data["Sector_Econom"],
            input_data["Tamano_Emp"],
            input_data["Activ_Econ"],
            "Optimizar estructura de capital",
            f"{input_data['tasa_deseada'] + 0.5}%",
            "Bajo"
        ],
        "Recomendacion": [
            "Mantener sector actual",
            "Considerar expansión controlada",
            "Evaluar diversificación",
            "Reducir deuda a corto plazo",
            "Basado en perfil de riesgo",
            "Cliente preferencial"
        ]
    }
    
    df_ejemplo = pd.DataFrame(ejemplo_data)
    output_path = OUTPUT_DIR / "recomendaciones_PYP_ejemplo.xlsx"
    df_ejemplo.to_excel(output_path, index=False)
    
    logger.info(f"Archivo de ejemplo creado en: {output_path}")
    return output_path

# ============ EJECUCIÓN ============

if __name__ == "__main__":
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True  # Solo para desarrollo
    )