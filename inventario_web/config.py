class Config:
    SECRET_KEY = "clave_super_segura_123"
    DB_HOST = "127.0.0.1"
    DB_PORT = 5433
    DB_NAME = "inventario"
    DB_USER = "n8n"
    DB_PASSWORD = "n8n_pass"

    WF2_URL = "http://127.0.0.1:5678/webhook/procesar-venta"
    WF3_URL = "http://127.0.0.1:5678/webhook/generar-reporte"