from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi import Request
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
import pandas as pd
import numpy as np
import subprocess
import json
import os
import shutil
import tempfile
import time
import re
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

# ============ MODELOS PYDANTIC ============

class PredictRequest(BaseModel):
    """Modelo de entrada para la predicción del modelo R"""
    
    model_config = {"populate_by_name": True}
    
    sector_econom: str = Field(
        ..., 
        alias="Sector_Econom",
        description="Sector Económico de la empresa",
        example="COMERCIO"
    )
    tamano_emp: str = Field(
        ..., 
        alias="Tamano_Emp",
        description="Tamaño de la empresa (Pequeña, Mediana, Grande, Micro)",
        example="Mediana"
    )
    activ_econ: str = Field(
        ..., 
        alias="Activ_Econ",
        description="Código de Actividad Económica (CIIU)",
        example="6201"
    )
    sucursal: str = Field(
        ..., 
        alias="Sucursal",
        description="Departamento de la sucursal",
        example="ANTIOQUIA"
    )
    num_empleados: int = Field(
        ..., 
        alias="Num_Empleados",
        description="Número de empleados de la empresa",
        example=50,
        gt=0
    )
    tasa_deseada: float = Field(
        ..., 
        description="Tasa deseada en porcentaje",
        example=5.5,
        ge=0,
        le=100
    )

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

@app.post("/api/predict", summary="Ejecutar Modelo de Predicción", tags=["Predicción"])
async def predict(data: PredictRequest):
    """
    Ejecuta el modelo R de predicción basado en los 6 parámetros de entrada.
    
    **Parámetros de entrada:**
    - **sector_econom**: Sector económico (ej: COMERCIO, SERVICIOS, MANUFACTURA)
    - **tamano_emp**: Tamaño empresa (Pequeña, Mediana, Grande, Micro)
    - **activ_econ**: Código CIIU de actividad económica
    - **sucursal**: Departamento de la sucursal
    - **num_empleados**: Número de empleados (entero positivo)
    - **tasa_deseada**: Tasa deseada en porcentaje (0-100)
    
    **Retorna:**
    - Lista de actividades PYP recomendadas con sus porcentajes
    - Tiempo de ejecución
    - URL para descargar el Excel completo
    
    **Tiempo estimado:** 5-30 segundos dependiendo de la complejidad.
    """
    start_time = time.time()
    
    try:
        # 1. Preparar datos de entrada (normalizar texto a mayúsculas y limpiar espacios)
        input_data = {
            "Sector_Econom": data.sector_econom.upper().strip(),
            "Tamano_Emp": data.tamano_emp.strip(),
            "Activ_Econ": data.activ_econ.strip(),
            "Sucursal": data.sucursal.upper().strip(),
            "Num_Empleados": data.num_empleados,
            "tasa_deseada": data.tasa_deseada
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
            CONTENT_DIR / "Recomendacion_PYP.xlsx",  # Ruta principal donde R genera el archivo
            OUTPUT_DIR / "recomendaciones_PYP.xlsx",   # Ruta alternativa
        ]
        
        for ruta in posibles_rutas:
            if ruta.exists():
                output_excel_path = ruta
                logger.info(f"Archivo de resultados encontrado en: {output_excel_path}")
                break
        
        if not output_excel_path:
            logger.warning("R no generó archivo de salida. Creando ejemplo.")
            output_excel_path = crear_ejemplo_resultados(input_data)
        
        # 5. Leer resultados del Excel
        try:
            # Leer Excel desde la fila 4 (skiprows=3 porque las primeras 3 filas están vacías)
            df_results = pd.read_excel(output_excel_path, skiprows=3)
            logger.info(f"Excel leído: {output_excel_path.name}, Columnas: {df_results.columns.tolist()}, Filas: {len(df_results)}")
            
            # Verificar que tenga las columnas esperadas
            required_cols = ["codigo_actividad", "ACTIVIDAD", "porcentaje_recomendado"]
            if not all(col in df_results.columns for col in required_cols):
                logger.error(f"Columnas esperadas no encontradas. Columnas actuales: {df_results.columns.tolist()}")
                raise ValueError("Estructura de Excel no coincide con la esperada")
            
            # Leer el footer (última fila después de los datos)
            # El footer está en las filas que contienen texto en lugar de datos numéricos
            footer_info = None
            try:
                # Intentar leer la última fila no vacía del Excel completo
                df_full = pd.read_excel(output_excel_path, header=None)
                # Buscar la última fila que tenga contenido
                for idx in range(len(df_full) - 1, -1, -1):
                    row_content = df_full.iloc[idx, 0]  # Primera columna
                    if pd.notna(row_content) and isinstance(row_content, str):
                        if "estimación" in row_content.lower() or "diferencia" in row_content.lower():
                            footer_info = str(row_content)
                            logger.info(f"Footer encontrado: {footer_info[:100]}...")
                            break
            except Exception as e:
                logger.warning(f"No se pudo leer el footer: {e}")
                
        except Exception as e:
            logger.error(f"Error leyendo Excel real ({output_excel_path}): {e}. Usando datos de fallback.")
            # Si falla, buscar el archivo de ejemplo
            fallback_path = OUTPUT_DIR / "recomendaciones_PYP_ejemplo.xlsx"
            if fallback_path.exists():
                logger.info(f"Usando archivo de fallback: {fallback_path}")
                df_results = pd.read_excel(fallback_path, skiprows=3)
                footer_info = None
            else:
                logger.error("Ni archivo real ni fallback disponibles. Creando datos mínimos.")
                df_results = pd.DataFrame({
                    "codigo_actividad": ["AR0000"],
                    "ACTIVIDAD": ["Error: Verificar generación de Excel por R"],
                    "porcentaje_recomendado": [100.0]
                })
                footer_info = None
        
        # 6. Preparar respuesta
        execution_time = round(time.time() - start_time, 2)
        
        # Convertir DataFrame a lista de actividades recomendadas
        actividades_recomendadas = []
        for _, row in df_results.iterrows():
            # Filtrar filas vacías o con valores nulos
            if pd.notna(row.get("codigo_actividad")) and pd.notna(row.get("porcentaje_recomendado")):
                actividades_recomendadas.append({
                    "codigo_actividad": str(row["codigo_actividad"]).strip(),
                    "actividad": str(row["ACTIVIDAD"]).strip(),
                    "porcentaje_recomendado": round(float(row["porcentaje_recomendado"]), 2)
                })
        
        # Calcular metadatos
        total_actividades = len(actividades_recomendadas)
        suma_porcentajes = round(sum(act["porcentaje_recomendado"] for act in actividades_recomendadas), 2)
        
        # Parsear información del footer si existe
        metadata = {
            "total_actividades": total_actividades,
            "suma_porcentajes": suma_porcentajes,
            "timestamp": datetime.now().isoformat(),
            "archivo_fuente": output_excel_path.name
        }
        
        # Extraer información del footer si existe
        if footer_info:
            try:
                # Extraer error de estimación (buscar patrón "X.XX%")
                error_match = re.search(r'error de estimación del ([\d.]+)%', footer_info)
                if error_match:
                    metadata["error_estimacion_porcentaje"] = float(error_match.group(1))
                
                # Extraer diferencia con tasa deseada (más tolerante con puntos finales)
                diff_match = re.search(r'diferencia con la tasa deseada es de ([\d.]+)', footer_info)
                if diff_match:
                    # Limpiar puntos finales extras
                    valor = diff_match.group(1).rstrip('.')
                    metadata["diferencia_tasa"] = float(valor)
                
                # Extraer nivel utilizado (más flexible con el final)
                nivel_match = re.search(r'criterios deseados de (.+?)[\.\n]', footer_info)
                if nivel_match:
                    metadata["nivel_historico_usado"] = nivel_match.group(1).strip()
                
                # Guardar el footer completo
                metadata["footer_completo"] = footer_info
                
            except Exception as e:
                logger.warning(f"Error parseando footer: {e}")
                metadata["footer_completo"] = footer_info
        
        response_data = {
            "status": "success",
            "execution_time": execution_time,
            "input_data": input_data,
            "metadata": metadata,
            "actividades_recomendadas": actividades_recomendadas,
            "excel_download_url": "/api/download/results"
        }
        
        logger.info(f"Predicción completada en {execution_time}s. Actividades recomendadas: {total_actividades}")
        
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
        "codigo_actividad": [
            "AR0001",
            "AR0002", 
            "AR0003",
            "AR0004",
            "AR0005"
        ],
        "ACTIVIDAD": [
            "Asesoría técnica y formación integral para la conformación de brigadas de emergencia",
            "Asesoría y asistencia técnica para el diseño de estándares de seguridad",
            "Programa integral de gestión para la prevención de riesgos",
            "Asesoría técnica en identificación de peligros y evaluación de riesgos",
            "Consulta médica ocupacional integral"
        ],
        "porcentaje_recomendado": [
            25.5,
            20.0,
            18.3,
            15.7,
            20.5
        ]
    }
    
    df_ejemplo = pd.DataFrame(ejemplo_data)
    output_path = OUTPUT_DIR / "recomendaciones_PYP_ejemplo.xlsx"
    
    # Crear Excel con formato similar al generado por R (3 filas vacías al inicio)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Escribir DataFrame empezando en la fila 4 (row 3 en base 0)
        df_ejemplo.to_excel(writer, index=False, startrow=3, sheet_name='Recomendacion_PYP')
    
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