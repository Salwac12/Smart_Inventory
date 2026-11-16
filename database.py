import sqlite3
from pathlib import Path

# Database location
BASE_DIR = Path(__file__).resolve().parent
DATABASE = BASE_DIR / "inventory.db"


def create_database():

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    # Components Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS components (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            part_number TEXT NOT NULL UNIQUE,

            component_name TEXT NOT NULL,

            category TEXT,

            quantity INTEGER NOT NULL DEFAULT 0
                CHECK (quantity >= 0),

            unit_price REAL NOT NULL DEFAULT 0
                CHECK (unit_price >= 0),

            supplier TEXT,

            minimum_stock INTEGER NOT NULL DEFAULT 0
                CHECK (minimum_stock >= 0),

            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Stock Transactions Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stock_transactions (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            component_id INTEGER NOT NULL,

            transaction_type TEXT NOT NULL
                CHECK (transaction_type IN ('IN', 'OUT')),

            quantity INTEGER NOT NULL
                CHECK (quantity > 0),

            remarks TEXT,

            transaction_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (component_id)
                REFERENCES components(id)
        )
    """)

    conn.commit()
    conn.close()

    print("Inventory database updated successfully!")


if __name__ == "__main__":
    create_database()