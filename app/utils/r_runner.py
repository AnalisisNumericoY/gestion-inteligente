import subprocess
import os
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

def ejecutar_script_r(r_script_path: Path, args: list = None, timeout: int = 300):
    """
    Ejecuta un script R y retorna el resultado
    
    Args:
        r_script_path: Ruta al script .R
        args: Argumentos para pasar al script
        timeout: Tiempo máximo en segundos
    
    Returns:
        dict: {"success": bool, "stdout": str, "stderr": str, "returncode": int}
    """
    if not r_script_path.exists():
        raise FileNotFoundError(f"Script R no encontrado: {r_script_path}")
    
    cmd = ["Rscript", str(r_script_path)]
    if args:
        cmd.extend([str(arg) for arg in args])
    
    logger.info(f"Ejecutando R: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            env={**os.environ, "LC_ALL": "C.UTF-8", "LANG": "C.UTF-8"}
        )
        
        return {
            "success": result.returncode == 0,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "returncode": result.returncode
        }
        
    except subprocess.TimeoutExpired:
        logger.error(f"Timeout al ejecutar script R: {r_script_path}")
        raise
    except Exception as e:
        logger.error(f"Error ejecutando R: {e}")
        raise