import pandas as pd
from pathlib import Path
import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)

def crear_input_excel(data: Dict[str, Any], output_path: Path):
    """
    Crea el archivo Excel de entrada para el modelo R
    
    Args:
        data: Diccionario con los 6 campos
        output_path: Ruta donde guardar el Excel
    """
    # Orden específico que espera R
    columnas_ordenadas = [
        "Sector_Econom",
        "Tamano_Emp", 
        "Activ_Econ",
        "Sucursal",
        "Num_Empleados",
        "tasa_deseada"
    ]
    
    # Crear DataFrame en el orden correcto
    df = pd.DataFrame({col: [data.get(col, "")] for col in columnas_ordenadas})
    
    # Guardar como Excel
    df.to_excel(output_path, index=False)
    logger.info(f"Excel creado en: {output_path}")
    
    return df

def leer_output_excel(excel_path: Path) -> pd.DataFrame:
    """
    Lee el Excel generado por R
    
    Args:
        excel_path: Ruta al Excel de salida
    
    Returns:
        DataFrame con los resultados
    """
    if not excel_path.exists():
        raise FileNotFoundError(f"Archivo de resultados no encontrado: {excel_path}")
    
    try:
        df = pd.read_excel(excel_path)
        logger.info(f"Excel leído. Filas: {len(df)}, Columnas: {df.columns.tolist()}")
        return df
    except Exception as e:
        logger.error(f"Error leyendo Excel: {e}")
        raise