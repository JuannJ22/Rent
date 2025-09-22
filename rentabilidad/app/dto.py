from pydantic import BaseModel
from typing import Optional


class GenerarInformeRequest(BaseModel):
    ruta_plantilla: str
    fecha: Optional[str] = None  # None = día anterior


class GenerarInformeResponse(BaseModel):
    ok: bool
    mensaje: str
    ruta_salida: Optional[str] = None
