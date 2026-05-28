
from Codigo.Sunat.sunat_scraper import open_sunat_page
from Codigo.Sunat.sunat_parser import parse_sunat

def consultar_ruc_service(page, ruc: str):
    open_sunat_page(page, ruc)
    return parse_sunat(page, ruc)