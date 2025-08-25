from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import pandas as pd
import io

def style_and_export_excel(df: pd.DataFrame, metadata: dict) -> io.BytesIO:
    """
    Exports a styled Excel file:
    - Inserts image at D2
    - Metadata rows: Brand, Campaign, Version, Start, End (A2:B6)
    - DataFrame starting from row 10
    - Hyperlinks automatically styled
    """
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        # Write DataFrame starting at row 10
        df.to_excel(writer, index=False, sheet_name="Processed Data", startrow=9)
        workbook = writer.book
        worksheet = writer.sheets["Processed Data"]

        # --- Insert Image ---
        image_path = "static/images/for_excel_output.png"
        try:
            img = XLImage(image_path)
            img.anchor = "D2"
            img.height = img.height / 2
            img.width = img.width / 2
            worksheet.add_image(img)
        except Exception as e:
            print(f"Could not insert image: {e}")

        # --- Styles ---
        title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        header_font_white = Font(name="Calibri", size=9, color="FFFFFF", bold=True)
        header_font_red = Font(name="Calibri", size=9, color="FF0000", bold=True)
        body_font = Font(name="Calibri", size=9)
        hyperlink_font = Font(name="Calibri", size=9, color="0000FF", underline="single")

        title_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        special_columns = {
            "Rent/ month", "Total rent", "Production", "Posting",
            "Ag Comm %", "Agency commission",
            "Advertising taxe %", "Advertising taxe",
            "Total Cost"
        }

        max_col_width = 40

        # --- Title Row ---
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
        cell = worksheet.cell(row=1, column=1)
        cell.value = "MP OOH CAMPAIGN"
        cell.font = title_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

        # --- Metadata Rows (A2:B6) ---
        meta_fields = ["Brand", "Campaign", "Version", "Start", "End"]
        for i, field in enumerate(meta_fields, start=2):
            worksheet.cell(row=i, column=1, value=field).font = Font(bold=True, name="Calibri", size=9)
            worksheet.cell(row=i, column=2, value=metadata.get(field, "")).font = body_font

        # --- Determine hyperlink columns dynamically ---
        hyperlink_columns = [col for col in df.columns if any(k in col.lower() for k in ["photo", "link", "address"])]

        # --- Format DataFrame area ---
        for col_num, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_num)
            for row_idx, cell in enumerate(worksheet[col_letter], start=1):
                if row_idx == 10:  # Header
                    if col_name in special_columns:
                        cell.fill = yellow_fill
                        cell.font = header_font_red
                    else:
                        cell.fill = red_fill
                        cell.font = header_font_white
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif row_idx > 10:  # Body
                    if col_name in hyperlink_columns and cell.value and str(cell.value).startswith("http"):
                        link = str(cell.value)
                        cell.value = "link" if "address" in col_name.lower() else "photo"
                        cell.hyperlink = link
                        cell.font = hyperlink_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.font = body_font
                        cell.alignment = Alignment(wrap_text=True, vertical="top")
                    # Zebra striping
                    if row_idx % 2 == 0:
                        cell.fill = gray_fill

                cell.border = thin_border

            # Auto column width
            max_length = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 2
            worksheet.column_dimensions[col_letter].width = min(max_length, max_col_width)

    output_buffer.seek(0)
    return output_buffer
