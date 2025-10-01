from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import pandas as pd
import io

def style_and_export_excel(df: pd.DataFrame, metadata: dict) -> io.BytesIO:
    # --- Reorder columns according to desired order ---
    columns_ordered = [
        "County", "City", "Address", "ID", "IDF", "Panel type", 
        "Format", "Base", "Height", "Size", "Faces", "Start", "End", 
        "No. of months", "Rent/month", "Total rent", "Production",
        "Posting", "Ag Comm %", "Agency commission", "Advertising taxe %", 
        "Advertising taxe", "Total Cost", "Photo Link", "GPS", "Sketch name", 
        "Tech Details", "idx", "NUME FURNIZOR", "CHIRIE FURNIZOR", "POSTARE FURNIZOR",
        "COST PRODUCTIE", "TIP MATERIAL", "__source_file"
    ]
    df = df[[col for col in columns_ordered if col in df.columns]]

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Processed Data", startrow=9)
        workbook = writer.book
        worksheet = writer.sheets["Processed Data"]

        # --- Hide default gridlines ---
        worksheet.sheet_view.showGridLines = False

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

        # --- Apply formulas ---
        start_row = 11
        for i, row in enumerate(df.itertuples(index=False), start=start_row):
            worksheet[f"Q{i}"].value = f"=J{i}*5"                       # Production
            worksheet[f"R{i}"].value = f"=AE{i}*1.2"                     # Posting
            worksheet[f"P{i}"].value = f"=O{i}*N{i}"                     # Total rent
            worksheet[f"T{i}"].value = f"=(R{i}+Q{i}+P{i})*S{i}"         # Agency commission
            worksheet[f"V{i}"].value = f"=((P{i}+R{i})*S{i}+P{i}+R{i})*0.03"  # Advertising taxe
            worksheet[f"W{i}"].value = f"=U{i}+T{i}+R{i}+Q{i}+P{i}"      # Total Cost

        # --- Styles ---
        title_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        header_font_white = Font(name="Calibri", size=9, color="FFFFFF", bold=True)
        header_font_red = Font(name="Calibri", size=9, color="FF0000", bold=True)
        body_font = Font(name="Calibri", size=9)
        hyperlink_font = Font(name="Calibri", size=9, color="0000FF", underline="single")

        title_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
        red_fill = PatternFill(start_color="F05055", end_color="F05055", fill_type="solid")
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
        last_col_letter = get_column_letter(len(df.columns))

        # --- Title Row ---
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
        cell = worksheet.cell(row=1, column=1)
        cell.value = "MP OOH CAMPAIGN"
        cell.font = title_font
        cell.fill = title_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border  # border for A1:B1

        # --- Metadata Rows (A2:B6) ---
        meta_fields = ["Brand", "Campaign", "Version", "Start", "End"]
        for i, field in enumerate(meta_fields, start=2):
            c1 = worksheet.cell(row=i, column=1, value=field)
            c1.font = Font(bold=True, name="Calibri", size=9)
            c1.border = thin_border
            c2 = worksheet.cell(row=i, column=2, value=metadata.get(field, ""))
            c2.font = body_font
            c2.border = thin_border

        # --- Determine hyperlink columns dynamically ---
        hyperlink_columns = [col for col in df.columns if any(k in col.lower() for k in ["photo", "link", "address", "tech details"])]

        # --- Format DataFrame area ---
        for col_num, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_num)
            for row_idx, cell in enumerate(worksheet[col_letter], start=1):

                # Rows 1–9, columns C→last column: untouched (no font, no fill, no border)
                if 1 <= row_idx <= 9 and col_num >= 3:
                    continue

                # Header row
                if row_idx == 10:
                    if col_name in special_columns:
                        cell.fill = yellow_fill
                        cell.font = header_font_red
                    else:
                        cell.fill = red_fill
                        cell.font = header_font_white
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border

                # Body rows
                elif row_idx > 10:
                    if col_name in hyperlink_columns and cell.value:
                        link = str(cell.value).strip()
                        if link.lower().startswith(("http://", "https://", "www.")):
                            if "tech details" in col_name.lower():
                                display_text = "sketch"
                            elif "address" in col_name.lower():
                                display_text = "link"
                            else:
                                display_text = "photo"
                            if not link.lower().startswith(("http://", "https://")):
                                link = "http://" + link
                            cell.value = display_text
                            cell.hyperlink = link
                            cell.font = hyperlink_font
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.font = body_font
                            cell.alignment = Alignment(wrap_text=True, vertical="top")
                    else:
                        cell.font = body_font
                        cell.alignment = Alignment(wrap_text=True, vertical="top")

                    if row_idx % 2 == 0:
                        cell.fill = gray_fill

                    cell.border = thin_border  # body borders

            # Auto column width
            max_length = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 2
            worksheet.column_dimensions[col_letter].width = min(max_length, max_col_width)

    output_buffer.seek(0)
    return output_buffer
