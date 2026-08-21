import os
from fastapi import Header, HTTPException

def validar_api_key(api_key_env: str):
    async def auth(x_api_key: str = Header(...)):
        api_key = os.getenv(api_key_env)
        if not api_key:
            raise HTTPException(
                status_code=500,
                detail="API Key no configurada"
            )
        if x_api_key != api_key:
            raise HTTPException(
                status_code=401,
                detail="API Key invalida"
            )
    return auth
