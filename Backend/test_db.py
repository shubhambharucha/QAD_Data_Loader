from database import get_connection

conn = get_connection()

print("Connected through database.py!")

conn.close()