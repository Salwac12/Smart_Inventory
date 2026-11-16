import sqlite3

from database import DATABASE, create_database

# Ensure the database table exists
create_database()

# Connect to database
conn = sqlite3.connect(DATABASE)

cursor = conn.cursor()

# Sample electronic components
components = [

    ("RES-10K", "10K Resistor", "Resistor", 10000, 0.50, "Supplier A", 2000),

    ("CAP-100UF", "100uF Capacitor", "Capacitor", 2000, 2.50, "Supplier B", 1500),

    ("IC-NE555", "NE555 Timer IC", "IC", 500, 15.00, "Supplier C", 1000),

    ("CON-2PIN", "2 Pin Connector", "Connector", 5000, 3.00, "Supplier D", 1000),

    ("LED-RED", "Red LED", "LED", 8000, 0.75, "Supplier A", 2000)

]

# Insert sample data
cursor.executemany("""

    INSERT OR IGNORE INTO components

    (part_number, component_name, category,
     quantity, unit_price, supplier, minimum_stock)

    VALUES (?, ?, ?, ?, ?, ?, ?)

""", components)

# Save changes
conn.commit()

# Close database
conn.close()

print("Sample components added successfully!")