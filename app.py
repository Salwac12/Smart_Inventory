from flask import Flask, render_template, request, redirect, url_for
from werkzeug.exceptions import RequestEntityTooLarge

import sqlite3
import math
import json
import uuid

from datetime import date, timedelta

from database import DATABASE, create_database
from bom_reader import read_bom

# NEW: Purchase Requirement Excel Export
from purchase_report import purchase_bp


app = Flask(__name__)

app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024

# NEW: Register Purchase Excel Download Routes
app.register_blueprint(purchase_bp)

create_database()


# -------------------------------------
# DATABASE CONNECTION
# -------------------------------------

def get_db_connection():

    conn = sqlite3.connect(DATABASE, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")

    return conn


# -------------------------------------
# CREATE BOM TABLES
# -------------------------------------

def create_bom_tables():

    conn = get_db_connection()

    try:

        with conn:

            conn.execute("""
                CREATE TABLE IF NOT EXISTS customer_boms (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    product_name TEXT NOT NULL,
                    revision TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS customer_bom_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    bom_id INTEGER NOT NULL,
                    sr_no INTEGER NOT NULL,
                    description TEXT NOT NULL,
                    quantity INTEGER NOT NULL,
                    unit TEXT,
                    type_pack TEXT,
                    make TEXT,

                    FOREIGN KEY (bom_id)
                    REFERENCES customer_boms(id)
                    ON DELETE CASCADE
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS bom_upload_previews (
                    token TEXT PRIMARY KEY,
                    items_json TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            conn.execute("""
                CREATE TABLE IF NOT EXISTS customer_bom_matches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,

                    bom_item_id INTEGER NOT NULL UNIQUE,
                    component_id INTEGER NOT NULL,

                    matched_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

                    FOREIGN KEY (bom_item_id)
                    REFERENCES customer_bom_items(id)
                    ON DELETE CASCADE,

                    FOREIGN KEY (component_id)
                    REFERENCES components(id)
                )
            """)

    finally:

        conn.close()


create_bom_tables()


# -------------------------------------
# DASHBOARD
# -------------------------------------

@app.route("/")
def home():

    conn = get_db_connection()

    total_components = conn.execute("""
        SELECT COUNT(*) FROM components
    """).fetchone()[0]

    total_stock = conn.execute("""
        SELECT COALESCE(SUM(quantity), 0)
        FROM components
    """).fetchone()[0]

    low_stock = conn.execute("""
        SELECT COUNT(*)
        FROM components
        WHERE quantity > 0
        AND quantity <= minimum_stock
    """).fetchone()[0]

    out_of_stock = conn.execute("""
        SELECT COUNT(*)
        FROM components
        WHERE quantity = 0
    """).fetchone()[0]

    inventory_value = conn.execute("""
        SELECT COALESCE(
            SUM(quantity * unit_price), 0
        )
        FROM components
    """).fetchone()[0]

    total_transactions = conn.execute("""
        SELECT COUNT(*)
        FROM stock_transactions
    """).fetchone()[0]

    recent_transactions = conn.execute("""
        SELECT
            st.id,
            c.part_number,
            st.transaction_type,
            st.quantity,
            st.transaction_date

        FROM stock_transactions st

        JOIN components c
        ON st.component_id = c.id

        ORDER BY
            st.transaction_date DESC,
            st.id DESC

        LIMIT 5
    """).fetchall()

    today = date.today()
    start_date = today - timedelta(days=6)

    movement_rows = conn.execute("""
        SELECT
            DATE(transaction_date) AS movement_date,
            transaction_type,
            SUM(quantity) AS total_quantity

        FROM stock_transactions

        WHERE DATE(transaction_date) >= ?
          AND DATE(transaction_date) <= ?

        GROUP BY
            DATE(transaction_date),
            transaction_type

        ORDER BY movement_date
    """, (
        start_date.isoformat(),
        today.isoformat()
    )).fetchall()

    conn.close()

    days = [
        start_date + timedelta(days=i)
        for i in range(7)
    ]

    chart_labels = [
        day.strftime("%d %b")
        for day in days
    ]

    movement_data = {}

    for row in movement_rows:

        key = (
            row["movement_date"],
            row["transaction_type"]
        )

        movement_data[key] = row["total_quantity"]

    stock_in_data = []
    stock_out_data = []

    for day in days:

        day_string = day.isoformat()

        stock_in_data.append(
            movement_data.get(
                (day_string, "IN"),
                0
            )
        )

        stock_out_data.append(
            movement_data.get(
                (day_string, "OUT"),
                0
            )
        )

    return render_template(
        "dashboard.html",
        total_components=total_components,
        total_stock=total_stock,
        low_stock=low_stock,
        out_of_stock=out_of_stock,
        inventory_value=inventory_value,
        total_transactions=total_transactions,
        recent_transactions=recent_transactions,
        chart_labels=chart_labels,
        stock_in_data=stock_in_data,
        stock_out_data=stock_out_data
    )


# -------------------------------------
# INVENTORY PAGE
# -------------------------------------

@app.route("/inventory")
def inventory():

    conn = get_db_connection()

    components = conn.execute("""
        SELECT *
        FROM components
        ORDER BY id
    """).fetchall()

    conn.close()

    return render_template(
        "inventory.html",
        components=components
    )


# -------------------------------------
# ADD COMPONENT
# -------------------------------------

@app.route("/add-component", methods=["GET", "POST"])
def add_component():

    if request.method == "POST":

        part_number = request.form.get(
            "part_number", ""
        ).strip()

        component_name = request.form.get(
            "component_name", ""
        ).strip()

        category = request.form.get(
            "category", ""
        ).strip()

        supplier = request.form.get(
            "supplier", ""
        ).strip()

        try:

            quantity = int(
                request.form.get("quantity", "0")
            )

            unit_price = float(
                request.form.get("unit_price", "0")
            )

            minimum_stock = int(
                request.form.get("minimum_stock", "0")
            )

        except (ValueError, TypeError):

            return render_template(
                "add_component.html",
                error="Please enter valid numeric values."
            ), 400

        if not part_number or not component_name:

            return render_template(
                "add_component.html",
                error="Part number and component name are required."
            ), 400

        if (
            quantity < 0
            or minimum_stock < 0
            or not math.isfinite(unit_price)
            or unit_price < 0
        ):

            return render_template(
                "add_component.html",
                error="Enter valid non-negative quantities and prices."
            ), 400

        conn = get_db_connection()

        try:

            with conn:

                cursor = conn.execute("""
                    INSERT INTO components
                    (
                        part_number,
                        component_name,
                        category,
                        quantity,
                        unit_price,
                        supplier,
                        minimum_stock
                    )

                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    part_number,
                    component_name,
                    category,
                    quantity,
                    unit_price,
                    supplier,
                    minimum_stock
                ))

                if quantity > 0:

                    conn.execute("""
                        INSERT INTO stock_transactions
                        (
                            component_id,
                            transaction_type,
                            quantity,
                            remarks
                        )

                        VALUES (?, 'IN', ?, ?)
                    """, (
                        cursor.lastrowid,
                        quantity,
                        "Opening stock"
                    ))

        except sqlite3.IntegrityError:

            return render_template(
                "add_component.html",
                error="Part number already exists or the data is invalid."
            ), 400

        finally:

            conn.close()

        return redirect(
            url_for("inventory")
        )

    return render_template(
        "add_component.html"
    )


# -------------------------------------
# STOCK IN / STOCK OUT
# -------------------------------------

@app.route("/stock", methods=["GET", "POST"])
def stock():

    conn = get_db_connection()

    error = None

    if request.method == "POST":

        try:

            component_id = int(
                request.form.get("component_id", "")
            )

            transaction_type = request.form.get(
                "transaction_type", ""
            )

            quantity = int(
                request.form.get("quantity", "")
            )

            remarks = request.form.get(
                "remarks", ""
            ).strip()

            if transaction_type not in ("IN", "OUT"):

                raise ValueError(
                    "Invalid transaction type."
                )

            if quantity <= 0:

                raise ValueError(
                    "Quantity must be greater than zero."
                )

            conn.execute("BEGIN IMMEDIATE")

            component = conn.execute("""
                SELECT *
                FROM components
                WHERE id = ?
            """, (component_id,)).fetchone()

            if component is None:

                raise ValueError(
                    "Component not found."
                )

            if transaction_type == "OUT":

                if component["quantity"] < quantity:

                    raise ValueError(
                        "Insufficient stock! Available quantity: "
                        + str(component["quantity"])
                    )

                new_quantity = (
                    component["quantity"] - quantity
                )

            else:

                new_quantity = (
                    component["quantity"] + quantity
                )

            conn.execute("""
                UPDATE components
                SET quantity = ?
                WHERE id = ?
            """, (
                new_quantity,
                component_id
            ))

            conn.execute("""
                INSERT INTO stock_transactions
                (
                    component_id,
                    transaction_type,
                    quantity,
                    remarks
                )

                VALUES (?, ?, ?, ?)
            """, (
                component_id,
                transaction_type,
                quantity,
                remarks
            ))

            conn.commit()

            return redirect(
                url_for("transactions")
            )

        except (ValueError, TypeError) as e:

            conn.rollback()
            error = str(e)

        except sqlite3.Error:

            conn.rollback()
            error = "Database error. Please try again."

    components = conn.execute("""
        SELECT *
        FROM components
        ORDER BY component_name
    """).fetchall()

    conn.close()

    return render_template(
        "stock.html",
        components=components,
        error=error
    )


# -------------------------------------
# TRANSACTION HISTORY
# -------------------------------------

@app.route("/transactions")
def transactions():

    conn = get_db_connection()

    history = conn.execute("""
        SELECT
            st.id,
            c.part_number,
            c.component_name,
            st.transaction_type,
            st.quantity,
            st.remarks,
            st.transaction_date

        FROM stock_transactions st

        JOIN components c
        ON st.component_id = c.id

        ORDER BY
            st.transaction_date DESC,
            st.id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "transactions.html",
        history=history
    )


# -------------------------------------
# EXCEL BOM UPLOAD AND PREVIEW
# -------------------------------------

@app.route("/upload-bom", methods=["GET", "POST"])
def upload_bom():

    bom_data = None
    error = None
    preview_token = None

    if request.method == "POST":

        file = request.files.get("bom_file")

        if not file or not file.filename:

            error = "Please select an Excel BOM file."

        elif not file.filename.lower().endswith(".xlsx"):

            error = "Only .xlsx Excel files are supported."

        else:

            try:

                bom_data = read_bom(file.stream)

                preview_token = uuid.uuid4().hex

                conn = get_db_connection()

                try:

                    with conn:

                        conn.execute("""
                            DELETE FROM bom_upload_previews
                            WHERE created_at < datetime('now', '-1 day')
                        """)

                        conn.execute("""
                            INSERT INTO bom_upload_previews
                            (
                                token,
                                items_json
                            )
                            VALUES (?, ?)
                        """, (
                            preview_token,
                            json.dumps(bom_data)
                        ))

                finally:

                    conn.close()

            except (
                ValueError,
                OSError,
                EOFError,
                KeyError,
                TypeError,
                sqlite3.Error
            ) as e:

                error = str(e)
                bom_data = None
                preview_token = None

            except Exception:

                error = (
                    "Unable to process this Excel file. "
                    "Please check that it is a valid .xlsx BOM."
                )

                bom_data = None
                preview_token = None

    return render_template(
        "upload_bom.html",
        bom_data=bom_data,
        error=error,
        preview_token=preview_token,
        saved=False,
        bom=None,
        saved_boms=None
    )


# -------------------------------------
# CONFIRM AND SAVE BOM
# -------------------------------------

@app.route("/save-bom", methods=["POST"])
def save_bom():

    product_name = request.form.get(
        "product_name", ""
    ).strip()

    revision = request.form.get(
        "revision", ""
    ).strip()

    preview_token = request.form.get(
        "preview_token", ""
    ).strip()

    if not product_name:
        return "Product name is required.", 400

    if not revision:
        return "BOM revision is required.", 400

    if len(product_name) > 150 or len(revision) > 50:
        return "Product name or revision is too long.", 400

    if not preview_token:
        return "Missing BOM preview. Upload again.", 400

    conn = get_db_connection()

    try:

        conn.execute("BEGIN IMMEDIATE")

        preview = conn.execute("""
            SELECT items_json
            FROM bom_upload_previews
            WHERE token = ?
              AND created_at >= datetime('now', '-1 day')
        """, (preview_token,)).fetchone()

        if preview is None:

            conn.rollback()

            return (
                "BOM preview has expired or was already saved. "
                "Please upload again.",
                400
            )

        bom_data = json.loads(preview["items_json"])

        if not bom_data:

            conn.rollback()
            return "BOM is empty.", 400

        cursor = conn.execute("""
            INSERT INTO customer_boms
            (
                product_name,
                revision
            )
            VALUES (?, ?)
        """, (
            product_name,
            revision
        ))

        bom_id = cursor.lastrowid

        for item in bom_data:

            conn.execute("""
                INSERT INTO customer_bom_items
                (
                    bom_id,
                    sr_no,
                    description,
                    quantity,
                    unit,
                    type_pack,
                    make
                )

                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                bom_id,
                item["sr_no"],
                item["description"],
                item["quantity"],
                item["unit"],
                item["type_pack"],
                item["make"]
            ))

        conn.execute("""
            DELETE FROM bom_upload_previews
            WHERE token = ?
        """, (preview_token,))

        conn.commit()

        return redirect(
            url_for("view_saved_bom", bom_id=bom_id)
        )

    except (sqlite3.Error, ValueError, KeyError, TypeError):

        conn.rollback()

        return (
            "Unable to save BOM. Please try again.",
            500
        )

    finally:

        conn.close()


# -------------------------------------
# VIEW ALL SAVED BOMS
# -------------------------------------

@app.route("/saved-boms")
def saved_boms():

    conn = get_db_connection()

    records = conn.execute("""
        SELECT
            b.id,
            b.product_name,
            b.revision,
            b.created_at,
            COUNT(i.id) AS total_components

        FROM customer_boms b

        LEFT JOIN customer_bom_items i
        ON i.bom_id = b.id

        GROUP BY
            b.id,
            b.product_name,
            b.revision,
            b.created_at

        ORDER BY b.id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "upload_bom.html",
        bom_data=None,
        error=None,
        preview_token=None,
        saved=False,
        bom=None,
        saved_boms=records
    )


# -------------------------------------
# VIEW INDIVIDUAL SAVED BOM
# -------------------------------------

@app.route("/saved-boms/<int:bom_id>")
def view_saved_bom(bom_id):

    conn = get_db_connection()

    bom = conn.execute("""
        SELECT *
        FROM customer_boms
        WHERE id = ?
    """, (bom_id,)).fetchone()

    if bom is None:

        conn.close()
        return "BOM not found.", 404

    bom_data = conn.execute("""
        SELECT
            id,
            sr_no,
            description,
            quantity,
            unit,
            type_pack,
            make

        FROM customer_bom_items

        WHERE bom_id = ?

        ORDER BY sr_no, id
    """, (bom_id,)).fetchall()

    conn.close()

    return render_template(
        "upload_bom.html",
        bom_data=bom_data,
        error=None,
        preview_token=None,
        saved=True,
        bom=bom,
        saved_boms=None
    )


# -------------------------------------
# INVENTORY MATCHING PAGE
# -------------------------------------

@app.route(
    "/saved-boms/<int:bom_id>/match",
    methods=["GET", "POST"]
)
def match_bom(bom_id):

    conn = get_db_connection()

    error = None
    success = request.args.get("saved") == "1"

    try:

        bom = conn.execute("""
            SELECT *
            FROM customer_boms
            WHERE id = ?
        """, (bom_id,)).fetchone()

        if bom is None:
            return "BOM not found.", 404

        if request.method == "POST":

            try:

                conn.execute("BEGIN IMMEDIATE")

                bom_items = conn.execute("""
                    SELECT id
                    FROM customer_bom_items
                    WHERE bom_id = ?
                """, (bom_id,)).fetchall()

                valid_item_ids = {
                    row["id"]
                    for row in bom_items
                }

                inventory_rows = conn.execute("""
                    SELECT id
                    FROM components
                """).fetchall()

                valid_component_ids = {
                    row["id"]
                    for row in inventory_rows
                }

                selections = {}

                for item_id in valid_item_ids:

                    field_name = f"component_{item_id}"

                    selected = request.form.get(
                        field_name
                    )

                    if selected is None:

                        raise ValueError(
                            "Missing component selection. "
                            "Please submit the complete form."
                        )

                    if selected == "":

                        selections[item_id] = None

                    else:

                        component_id = int(selected)

                        if component_id not in valid_component_ids:

                            raise ValueError(
                                "Invalid inventory component selected."
                            )

                        selections[item_id] = component_id

                for item_id, component_id in selections.items():

                    if component_id is None:

                        conn.execute("""
                            DELETE FROM customer_bom_matches
                            WHERE bom_item_id = ?
                        """, (item_id,))

                    else:

                        conn.execute("""
                            INSERT INTO customer_bom_matches
                            (
                                bom_item_id,
                                component_id
                            )
                            VALUES (?, ?)

                            ON CONFLICT(bom_item_id)
                            DO UPDATE SET
                                component_id = excluded.component_id,
                                matched_at = CURRENT_TIMESTAMP
                        """, (
                            item_id,
                            component_id
                        ))

                conn.commit()

                return redirect(
                    url_for(
                        "match_bom",
                        bom_id=bom_id,
                        saved=1
                    )
                )

            except (ValueError, TypeError) as e:

                conn.rollback()
                error = str(e)

            except sqlite3.Error:

                conn.rollback()
                error = "Database error while saving matches."

        bom_items = conn.execute("""
            SELECT
                i.id,
                i.sr_no,
                i.description,
                i.quantity,
                i.unit,
                i.type_pack,
                i.make,

                m.component_id,

                c.part_number AS matched_part_number,
                c.component_name AS matched_name,
                c.quantity AS available_stock

            FROM customer_bom_items i

            LEFT JOIN customer_bom_matches m
                ON m.bom_item_id = i.id

            LEFT JOIN components c
                ON c.id = m.component_id

            WHERE i.bom_id = ?

            ORDER BY i.sr_no, i.id
        """, (bom_id,)).fetchall()

        inventory_components = conn.execute("""
            SELECT
                id,
                part_number,
                component_name,
                quantity

            FROM components

            ORDER BY component_name, part_number
        """).fetchall()

        matched_count = sum(
            1 for item in bom_items
            if item["component_id"] is not None
        )

        total_count = len(bom_items)

        return render_template(
            "match_bom.html",
            bom=bom,
            bom_items=bom_items,
            inventory_components=inventory_components,
            matched_count=matched_count,
            unmatched_count=total_count - matched_count,
            total_count=total_count,
            error=error,
            success=success
        )

    finally:

        conn.close()


# -------------------------------------
# MATERIAL REQUIREMENT CALCULATION
# -------------------------------------

@app.route(
    "/saved-boms/<int:bom_id>/calculate",
    methods=["GET", "POST"]
)
def calculate_materials(bom_id):

    conn = get_db_connection()

    try:

        bom = conn.execute("""
            SELECT *
            FROM customer_boms
            WHERE id = ?
        """, (bom_id,)).fetchone()

        if bom is None:
            return "BOM not found.", 404

        error = None
        results = None
        production_quantity = None

        if request.method == "POST":

            try:

                production_quantity = int(
                    request.form.get(
                        "production_quantity", ""
                    )
                )

                if not 1 <= production_quantity <= 1000000:

                    raise ValueError(
                        "Production quantity must be between "
                        "1 and 1,000,000."
                    )

            except (ValueError, TypeError):

                error = (
                    "Please enter a valid whole-number "
                    "production quantity between 1 and 1,000,000."
                )

            if error is None:

                rows = conn.execute("""
                    SELECT
                        i.id,
                        i.sr_no,
                        i.description,
                        i.quantity,
                        i.unit,
                        i.type_pack,
                        i.make,

                        m.component_id,

                        c.part_number,
                        c.component_name,
                        c.quantity AS available_stock,
                        c.unit_price

                    FROM customer_bom_items i

                    LEFT JOIN customer_bom_matches m
                        ON m.bom_item_id = i.id

                    LEFT JOIN components c
                        ON c.id = m.component_id

                    WHERE i.bom_id = ?

                    ORDER BY i.sr_no, i.id
                """, (bom_id,)).fetchall()

                grouped_components = {}
                unmatched_items = []

                for row in rows:

                    required = (
                        row["quantity"] * production_quantity
                    )

                    component_id = row["component_id"]

                    if component_id is None:

                        unmatched_items.append({
                            "sr_no": row["sr_no"],
                            "description": row["description"],
                            "type_pack": row["type_pack"],
                            "make": row["make"],
                            "quantity": row["quantity"],
                            "required": required
                        })

                        continue

                    if component_id not in grouped_components:

                        grouped_components[component_id] = {
                            "part_number": row["part_number"],
                            "component_name": row["component_name"],
                            "quantity_per_pcb": 0,
                            "required": 0,
                            "available": row["available_stock"],
                            "unit_price": row["unit_price"]
                        }

                    grouped_components[component_id][
                        "quantity_per_pcb"
                    ] += row["quantity"]

                    grouped_components[component_id][
                        "required"
                    ] += required

                matched_items = []

                shortage_count = 0
                shortage_cost = 0.0

                for component in grouped_components.values():

                    available = component["available"]
                    required = component["required"]

                    shortage = max(
                        required - available,
                        0
                    )

                    cost = (
                        shortage * component["unit_price"]
                    )

                    if shortage > 0:
                        shortage_count += 1

                    shortage_cost += cost

                    component["shortage"] = shortage
                    component["shortage_cost"] = cost

                    matched_items.append(component)

                matched_items.sort(
                    key=lambda item: item["part_number"].lower()
                )

                results = {
                    "matched_items": matched_items,
                    "unmatched_items": unmatched_items,
                    "matched_count": len(matched_items),
                    "unmatched_count": len(unmatched_items),
                    "shortage_count": shortage_count,
                    "shortage_cost": shortage_cost
                }

        return render_template(
            "calculate_materials.html",
            bom=bom,
            production_quantity=production_quantity,
            results=results,
            error=error
        )

    finally:

        conn.close()


# -------------------------------------
# EXCEL UPLOAD SIZE ERROR
# -------------------------------------

@app.errorhandler(RequestEntityTooLarge)
def file_too_large(error):

    return render_template(
        "upload_bom.html",
        bom_data=None,
        error="File is too large. Maximum allowed size is 5 MB.",
        preview_token=None,
        saved=False,
        bom=None,
        saved_boms=None
    ), 413


# -------------------------------------
# RUN APPLICATION
# -------------------------------------

if __name__ == "__main__":

    app.run(debug=True)