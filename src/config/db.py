import pyodbc

def get_db_connection() -> pyodbc.Connection:
    """Crea y retorna una conexión a SQL Server para el proyecto
    Escalas de Crecimiento Soriana."""
    
    conn_str = (
        "DRIVER={ODBC Driver 18 for SQL Server};"
        "SERVER=ATL20AF2222SQ19;"
        "DATABASE=SORIANA_MX_2024_PROD_F;"
        "Trusted_Connection=yes;"
    )
    return pyodbc.connect(conn_str)