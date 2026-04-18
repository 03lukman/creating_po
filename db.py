import firebirdsql
import os
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv()

HOST = os.getenv('HOST')
DATABASE = os.getenv('DATABASE')
USER = os.getenv('USER')
PASSWORD = os.getenv('PASSWORD')
PORT = int(os.getenv('PORT', 3050))


def fetch_data():
    conn = firebirdsql.connect(
        host=HOST,
        database=DATABASE,
        user=USER,
        password=PASSWORD,
        port=PORT
    )

    cursor = conn.cursor()

    today = datetime.now().date()
    start_date = today - timedelta(days=10)
    end_date = today + timedelta(days=10)

    sql = f"""
    SELECT 
        c.SODATE,
        c.SONO || ASCII_CHAR(65 + b.SOSEQ),
        b.ITEMNO,
        b.ITEMOVDESC,
        d.QUANTITY,
        b.ITEMUNIT,
        b.ITEMOVDESC
    FROM ARINV a
    JOIN ARINVDET b ON a.ARINVOICEID = b.ARINVOICEID
    JOIN SO c ON b.SOID = c.SOID
    JOIN SODET d ON b.SOID = d.SOID AND b.SOSEQ = d.SEQ
    WHERE c.SODATE BETWEEN '{start_date}' AND '{end_date}'
    AND c.SONO LIKE 'SON%'
    ORDER BY c.SODATE, c.SONO, b.SOSEQ
    """

    cursor.execute(sql)

    data = cursor.fetchall()
    columns = [col[0] for col in cursor.description]

    conn.close()

    return data, columns