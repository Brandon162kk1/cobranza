from Codigo.Reniec.reniec_scraper import open_reniec_page,buscar_dni
from Codigo.Reniec.reniec_parser import parse_reniec
import os
from fastapi import HTTPException

#----- Variables de Entorno -------
url_reniec = os.getenv("url_reniec")

def consultar_dni_service(page, dni: str):

    open_reniec_page(page,url_reniec)
    buscar_dni(page, dni)
    resultado = parse_reniec(page, dni)

    if resultado is None:
        raise HTTPException(status_code=404,detail="DNI no encontrado")

    return resultado