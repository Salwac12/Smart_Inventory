from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import math


# -------------------------------------
# EXCEL BOM READER
# -------------------------------------

def read_bom(file):

    bom_data = []
    workbook = None

    try:

        # Open uploaded Excel file
        workbook = load_workbook(
            file,
            read_only=True,
            data_only=True
        )

        # -------------------------------------
        # SELECT BOM WORKSHEET
        # -------------------------------------

        if "BOM" not in workbook.sheetnames:

            raise ValueError(
                "Worksheet named 'BOM' was not found."
            )

        sheet = workbook["BOM"]

        # -------------------------------------
        # READ COMPONENT ROWS
        # -------------------------------------

        # Ignitor BOM format:
        #
        # Row 7 = Headings
        # Row 8 onward = Components
        #
        # B = Sr. No.
        # E = Description
        # H = Quantity
        # I = Unit
        # J = Type/Pack
        # K = Make

        for row in sheet.iter_rows(
            min_row=8,
            max_col=11,
            values_only=True
        ):

            sr_no = row[1]
            description = row[4]
            quantity = row[7]
            unit = row[8]
            type_pack = row[9]
            make = row[10]

            # -------------------------------------
            # SKIP ROWS WITHOUT DESCRIPTION
            # -------------------------------------

            if description is None:

                continue

            description = str(description).strip()

            if not description:

                continue

            # -------------------------------------
            # SKIP NON-COMPONENT ROWS
            # -------------------------------------

            try:

                if isinstance(sr_no, bool):
                    continue

                serial_number = float(sr_no)

                if (
                    not math.isfinite(serial_number)
                    or not serial_number.is_integer()
                    or serial_number <= 0
                ):

                    continue

                serial_number = int(serial_number)

            except (ValueError, TypeError, OverflowError):

                continue

            # -------------------------------------
            # SKIP ROWS WITH INVALID QUANTITY
            # -------------------------------------

            try:

                if isinstance(quantity, bool):
                    continue

                qty = float(quantity)

                if (
                    not math.isfinite(qty)
                    or not qty.is_integer()
                    or qty <= 0
                ):

                    continue

                qty = int(qty)

            except (ValueError, TypeError, OverflowError):

                continue

            # -------------------------------------
            # SAVE VALID COMPONENT
            # -------------------------------------

            bom_data.append({

                "sr_no": serial_number,

                "description": description,

                "quantity": qty,

                "unit": (
                    str(unit).strip()
                    if unit is not None
                    else ""
                ),

                "type_pack": (
                    str(type_pack).strip()
                    if type_pack is not None
                    else ""
                ),

                "make": (
                    str(make).strip()
                    if make is not None
                    else ""
                )

            })

        # -------------------------------------
        # CHECK RESULT
        # -------------------------------------

        if not bom_data:

            raise ValueError(
                "No valid components found in the Excel BOM."
            )

        return bom_data

    except InvalidFileException as error:

        raise ValueError(
            "Invalid Excel file. Please upload a valid .xlsx BOM."
        ) from error

    finally:

        if workbook is not None:
            workbook.close()