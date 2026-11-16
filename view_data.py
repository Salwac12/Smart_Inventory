import sqlite3

from database import DATABASE

conn = sqlite3.connect(DATABASE)

cursor = conn.cursor()

cursor.execute("SELECT * FROM components")

components = cursor.fetchall()

print("\nINVENTORY COMPONENTS")
print("-" * 80)

for component in components:

    print(
        f"ID: {component[0]} | "
        f"Part: {component[1]} | "
        f"Name: {component[2]} | "
        f"Quantity: {component[4]}"
    )

conn.close()