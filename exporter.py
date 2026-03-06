import pandas as pd
from io import BytesIO
from database import get_conn


def export_excel():

    conn = get_conn()

    df = pd.read_sql_query(
        "SELECT * FROM orders",
        conn
    )

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df.to_excel(writer,index=False)

    output.seek(0)

    return output