"""Auditoría de CFDIs ante el SAT y cálculo de contraprestación variable."""

from .auditoria import Configuracion, Resultado, auditar, auditar_archivos
from .lector import CFDI, expandir_archivos, leer_carpeta, leer_cfdi
from .reporte import generar_excel

__all__ = ["CFDI", "Configuracion", "Resultado", "auditar", "auditar_archivos",
           "expandir_archivos", "generar_excel", "leer_carpeta", "leer_cfdi"]
