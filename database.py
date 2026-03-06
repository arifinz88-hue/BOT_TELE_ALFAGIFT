import sqlite3
from contextlib import closing

DB = "database.db"

def get_conn():
    return sqlite3.connect(DB)


def init_db():

    with closing(get_conn()) as conn:

        conn.execute("""
        CREATE TABLE IF NOT EXISTS orders(
            oid TEXT,
            tanggal TEXT,
            toko TEXT,
            nama TEXT,
            produk TEXT,
            qty INTEGER
        )
        """)

        conn.commit()


def insert_orders(rows):

    with closing(get_conn()) as conn:

        conn.executemany("""
        INSERT INTO orders
        VALUES (?,?,?,?,?,?)
        """,rows)

        conn.commit()


def search(keyword):

    keyword = f"%{keyword}%"

    with closing(get_conn()) as conn:

        rows = conn.execute("""
        SELECT nama, oid, produk, qty, toko
        FROM orders
        WHERE nama LIKE ?
        OR oid LIKE ?
        LIMIT 30
        """,(keyword,keyword)).fetchall()

    return rows


def produk_summary():

    with closing(get_conn()) as conn:

        rows = conn.execute("""
        SELECT produk, SUM(qty)
        FROM orders
        GROUP BY produk
        ORDER BY SUM(qty) DESC
        LIMIT 30
        """).fetchall()

    return rows


def toko_summary():

    with closing(get_conn()) as conn:

        rows = conn.execute("""
        SELECT toko, SUM(qty)
        FROM orders
        GROUP BY toko
        ORDER BY SUM(qty) DESC
        LIMIT 30
        """).fetchall()

    return rows


def status():

    with closing(get_conn()) as conn:

        total = conn.execute(
            "SELECT COUNT(*) FROM orders"
        ).fetchone()[0]

    return total