from flask import Blueprint, request, send_file, abort

from io import BytesIO
from datetime import datetime
from decimal import Decimal
from collections import defaultdict

from openpyxl import Workbook

from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)

from openpyxl.utils import get_column_letter

import sqlite3

from database import (
    DATABASE,
    get_postgres_connection,
    using_postgresql
)


# =====================================================
# PURCHASE REQUIREMENT BLUEPRINT
# =====================================================

purchase_bp = Blueprint(
    "purchase",
    __name__
)


# =====================================================
# DATABASE CONNECTION
# =====================================================

def get_report_connection():

    """
    Local computer:
        SQLite database.

    Vercel:
        Neon PostgreSQL database.
    """

    if using_postgresql():

        return get_postgres_connection()

    conn = sqlite3.connect(
        DATABASE,
        timeout=10
    )

    conn.row_factory = sqlite3.Row

    conn.execute(
        "PRAGMA foreign_keys = ON"
    )

    return conn


# =====================================================
# DATABASE QUERY HELPER
# =====================================================

def execute_query(
    conn,
    query,
    parameters=()
):

    """
    SQLite uses ? placeholders.

    PostgreSQL uses %s placeholders.
    """

    if using_postgresql():

        query = query.replace(
            "?",
            "%s"
        )

    return conn.execute(
        query,
        parameters
    )


# =====================================================
# SAFE TEXT FOR EXCEL
# =====================================================

def safe_text(value):

    if value is None:

        return ""

    text = str(value)

    # Protect against Excel formula injection.

    if text.startswith(
        ("=", "+", "-", "@", "\t", "\r")
    ):

        return "'" + text

    return text


# =====================================================
# CONVERT DATABASE NUMBERS
# =====================================================

def to_number(value):

    """
    SQLite may return float.

    PostgreSQL NUMERIC may return Decimal.

    Convert numbers for consistent Excel output.
    """

    if value is None:

        return 0.0

    if isinstance(value, Decimal):

        return float(value)

    return float(value)


# =====================================================
# STYLE EXCEL WORKSHEET
# =====================================================

def style_sheet(
    ws,
    column_widths
):

    # Header background

    header_fill = PatternFill(
        start_color="1E293B",
        end_color="1E293B",
        fill_type="solid"
    )

    # Header font

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    # Table borders

    thin_border = Border(
        bottom=Side(
            style="thin",
            color="E2E8F0"
        )
    )

    # Apply header formatting

    for cell in ws[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[1].height = 30

    # Apply row formatting

    for row in ws.iter_rows(
        min_row=2
    ):

        for cell in row:

            cell.border = thin_border

            cell.alignment = Alignment(
                vertical="center"
            )

    # Set column widths

    for index, width in enumerate(
        column_widths,
        start=1
    ):

        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    # Freeze table header

    ws.freeze_panes = "A2"

    # Enable Excel filters

    ws.auto_filter.ref = ws.dimensions


# =====================================================
# PURCHASE REQUIREMENT EXCEL ROUTE
# =====================================================

@purchase_bp.route(
    "/saved-boms/<int:bom_id>/purchase-excel"
)
def download_purchase_excel(bom_id):

    # -----------------------------------------
    # VALIDATE PRODUCTION QUANTITY
    # -----------------------------------------

    try:

        production_quantity = int(
            request.args.get(
                "production_quantity",
                ""
            )
        )

        if not 1 <= production_quantity <= 1000000:

            raise ValueError

    except (ValueError, TypeError):

        return (
            "Invalid production quantity. "
            "Enter a value between 1 and 1,000,000.",
            400
        )

    # -----------------------------------------
    # CONNECT TO DATABASE
    # -----------------------------------------

    conn = get_report_connection()

    try:

        # -------------------------------------
        # GET PRODUCT DETAILS
        # -------------------------------------

        bom = execute_query(
            conn,
            """
            SELECT *
            FROM customer_boms
            WHERE id = ?
            """,
            (bom_id,)
        ).fetchone()

        if bom is None:

            abort(404)

        # -------------------------------------
        # GET BOM AND INVENTORY DETAILS
        # -------------------------------------

        rows = execute_query(
            conn,
            """
            SELECT

                i.sr_no,
                i.description,
                i.quantity AS bom_quantity,
                i.unit,
                i.type_pack,
                i.make,

                m.component_id,

                c.part_number,
                c.component_name,
                c.quantity AS available_stock,
                c.unit_price,
                c.supplier

            FROM customer_bom_items i

            LEFT JOIN customer_bom_matches m
                ON m.bom_item_id = i.id

            LEFT JOIN components c
                ON c.id = m.component_id

            WHERE i.bom_id = ?

            ORDER BY i.sr_no, i.id
            """,
            (bom_id,)
        ).fetchall()

    finally:

        conn.close()

    # =================================================
    # GROUP MATCHED MATERIALS
    # =================================================

    grouped = {}

    unmatched = []

    for row in rows:

        required = (
            row["bom_quantity"]
            * production_quantity
        )

        component_id = row["component_id"]

        # -------------------------------------
        # UNMATCHED BOM COMPONENTS
        # -------------------------------------

        if component_id is None:

            unmatched.append({

                "sr_no": row["sr_no"],

                "description": row["description"],

                "type_pack": row["type_pack"],

                "make": row["make"],

                "qty_per_pcb": row["bom_quantity"],

                "required": required

            })

            continue

        # -------------------------------------
        # GROUP MATCHED INVENTORY ITEMS
        # -------------------------------------

        if component_id not in grouped:

            grouped[component_id] = {

                "part_number": row["part_number"],

                "component_name": row["component_name"],

                "qty_per_pcb": 0,

                "required": 0,

                "available": row["available_stock"],

                "unit_price": to_number(
                    row["unit_price"]
                ),

                "supplier": row["supplier"]

            }

        grouped[component_id][
            "qty_per_pcb"
        ] += row["bom_quantity"]

        grouped[component_id][
            "required"
        ] += required

    # =================================================
    # CALCULATE MATERIAL SHORTAGES
    # =================================================

    purchase_items = []

    for item in grouped.values():

        shortage = max(
            item["required"] - item["available"],
            0
        )

        if shortage <= 0:

            continue

        purchase_items.append({

            **item,

            "order_quantity": shortage,

            "estimated_cost": (
                shortage * item["unit_price"]
            )

        })

    # Sort by part number

    purchase_items.sort(
        key=lambda item: item["part_number"].lower()
    )

    # =================================================
    # CREATE EXCEL WORKBOOK
    # =================================================

    wb = Workbook()

    ws = wb.active

    ws.title = "To Order"

    # =================================================
    # SHEET 1: TO ORDER
    # =================================================

    ws.append([

        "Sr. No.",

        "Part Number",

        "Component Description",

        "Qty/PCB",

        "Total Required",

        "Available Stock",

        "Order Quantity",

        "Unit Price (INR)",

        "Estimated Cost (INR)",

        "Supplier",

        "Purchase Status"

    ])

    total_cost = 0.0

    for index, item in enumerate(
        purchase_items,
        start=1
    ):

        total_cost += item["estimated_cost"]

        ws.append([

            index,

            safe_text(
                item["part_number"]
            ),

            safe_text(
                item["component_name"]
            ),

            item["qty_per_pcb"],

            item["required"],

            item["available"],

            item["order_quantity"],

            item["unit_price"],

            item["estimated_cost"],

            safe_text(
                item["supplier"]
            ),

            "Pending Order"

        ])

    # No shortage condition

    if not purchase_items:

        ws.append([
            "",
            "No confirmed shortages",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
        ])

    style_sheet(
        ws,
        [
            12,
            22,
            38,
            14,
            18,
            18,
            18,
            20,
            23,
            25,
            20
        ]
    )

    # Format currency columns

    for row in ws.iter_rows(
        min_row=2
    ):

        for index in (8, 9):

            row[index - 1].number_format = (
                '#,##0.00'
            )

    # =================================================
    # SHEET 2: NEEDS REVIEW
    # =================================================

    review_ws = wb.create_sheet(
        "Needs Review"
    )

    review_ws.append([

        "Sr. No.",

        "Component Description",

        "Type/Pack",

        "Make",

        "Qty/PCB",

        "Total Required",

        "Available Stock",

        "Suggested Order Qty",

        "Status"

    ])

    for item in unmatched:

        review_ws.append([

            item["sr_no"],

            safe_text(
                item["description"]
            ),

            safe_text(
                item["type_pack"]
            ),

            safe_text(
                item["make"]
            ),

            item["qty_per_pcb"],

            item["required"],

            "Unknown",

            "",

            "Verify Inventory"

        ])

    if not unmatched:

        review_ws.append([
            "",
            "All BOM components matched",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            ""
        ])

    style_sheet(
        review_ws,
        [
            12,
            45,
            20,
            22,
            15,
            20,
            22,
            24,
            24
        ]
    )

    # =================================================
    # SHEET 3: REPORT SUMMARY
    # =================================================

    summary_ws = wb.create_sheet(
        "Report Summary"
    )

    summary_ws.append([

        "Field",

        "Value"

    ])

    summary_ws.append([

        "Product",

        safe_text(
            bom["product_name"]
        )

    ])

    summary_ws.append([

        "BOM Revision",

        safe_text(
            bom["revision"]
        )

    ])

    summary_ws.append([

        "Production Quantity",

        production_quantity

    ])

    summary_ws.append([

        "Confirmed Shortage Materials",

        len(purchase_items)

    ])

    summary_ws.append([

        "Unmatched BOM Rows",

        len(unmatched)

    ])

    summary_ws.append([

        "Estimated Purchase Cost (INR)",

        total_cost

    ])

    summary_ws.append([

        "Generated On",

        datetime.now().strftime(
            "%d-%m-%Y %H:%M"
        )

    ])

    summary_ws.append([

        "Stock Basis",

        "Current inventory at export time"

    ])

    summary_ws.append([

        "Important",

        "Unmatched items require stock verification"

    ])

    style_sheet(
        summary_ws,
        [
            38,
            65
        ]
    )

    # Format total cost

    summary_ws["B7"].number_format = (
        '#,##0.00'
    )

    # =================================================
    # GENERATE EXCEL FILE
    # =================================================

    output = BytesIO()

    wb.save(
        output
    )

    output.seek(0)

    filename = (
        f"BOM_{bom_id}_"
        f"Purchase_Requirement_"
        f"{production_quantity}_PCBs.xlsx"
    )

    # =================================================
    # DOWNLOAD EXCEL FILE
    # =================================================

    return send_file(

        output,

        as_attachment=True,

        download_name=filename,

        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )

    )