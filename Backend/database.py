import psycopg2
from datetime import date, datetime

def get_connection():
    return psycopg2.connect(
        host="10.0.54.52",
        port=5432,
        database="central_licensing",
        user="postgres",
        password="Admin123"
    )

def get_user_license(username):
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            u.username,
            l.license_key,
            l.expiry_date,
            l.status
        FROM users u
        JOIN licenses l
            ON u.customer_id = l.customer_id
        WHERE u.username = %s
    """, (username,))

    row = cur.fetchone()

    cur.close()
    conn.close()

    if not row:
        return None

    return {
        "username": row[0],
        "license_key": row[1],
        "expiry_date": row[2],
        "status": row[3]
    }

def validate_user_license(username):
    license_info = get_user_license(username)

    if not license_info:
        return False, "User is not licensed to use this application."

    if license_info["status"] != "ACTIVE":
        return False, "User license is inactive."

    if license_info["expiry_date"] < date.today():
        return False, "User license has expired."

    return True, "License valid."