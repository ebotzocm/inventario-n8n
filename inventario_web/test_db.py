import psycopg2

try:
    conn = psycopg2.connect(
        host="127.0.0.1",
        port=5433,
        dbname="inventario",
        user="n8n",
        password="n8n_pass",
    )
    print("Conexion exitosa")
    conn.close()
except Exception as e:
    print("Error:", repr(e))