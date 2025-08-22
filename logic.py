import pandas as pd
import glob
import os
import re
import unicodedata
import numpy as np
import datetime
from openpyxl import load_workbook

class TextUtils:
    @staticmethod
    def normalize_text(text: str) -> str:
        if not isinstance(text, str):
            text = str(text)
        text = text.lower().strip()
        text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
        text = text.replace("\n", " ").replace("/", " ")
        text = re.sub(r"\s+", " ", text)
        return text

    @staticmethod
    def contains_whole_word(text: str, phrase: str) -> bool:
        text = TextUtils.normalize_text(text)
        phrase = TextUtils.normalize_text(phrase)
        pattern = r'(?<!\w)' + re.escape(phrase) + r'(?!\w)'
        return re.search(pattern, text) is not None


class ColumnMatcher:
    @staticmethod
    def score_match(col: str, keywords, priority=None, avoid=None) -> int:
        if priority is None: priority = []
        if avoid is None: avoid = []

        txt = TextUtils.normalize_text(col)
        score = 0

        for p in priority:
            p_n = TextUtils.normalize_text(p)
            if txt == p_n:
                score = max(score, 100)
            elif TextUtils.contains_whole_word(txt, p_n):
                score = max(score, 90)
            elif p_n in txt:
                score = max(score, 80)

        for k in keywords:
            k_n = TextUtils.normalize_text(k)
            if txt == k_n:
                score = max(score, 70)
            elif TextUtils.contains_whole_word(txt, k_n):
                score = max(score, 60)
            elif k_n in txt:
                score = max(score, 40)

        for a in avoid:
            a_n = TextUtils.normalize_text(a)
            if TextUtils.contains_whole_word(txt, a_n) or a_n in txt:
                score -= 50

        return score

    @staticmethod
    def find_best_match(columns, keywords, priority=None, avoid=None):
        if priority is None: priority = []
        if avoid is None: avoid = []

        scored = [(col, ColumnMatcher.score_match(col, keywords, priority, avoid)) 
                  for col in columns if ColumnMatcher.score_match(col, keywords, priority, avoid) > 0]

        if not scored:
            return None, []

        scored.sort(key=lambda x: (-x[1], len(str(x[0]))))
        return scored[0][0], scored


class DataProcessor:
    def __init__(self, groups, directory=None):
        self.groups = groups
        self.directory = directory


    @staticmethod
    def extract_hyperlinks(file_input):
        """
        Returns a dict mapping (row_idx, col_idx) -> hyperlink URL
        row_idx, col_idx are 0-based to match pandas indexing.
        """
        if hasattr(file_input, "seek"):  # BytesIO case
            file_input.seek(0)
        wb = load_workbook(file_input, data_only=True)
        ws = wb.active

        links = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.hyperlink:
                    # openpyxl row/column are 1-based, convert to 0-based
                    links[(cell.row - 1, cell.column - 1)] = cell.hyperlink.target
        return links
    

    def extract_standardized_dataframe(self, file_input, file_name=None):
        """
        file_input can be:
        - a string path
        - a file-like object (BytesIO)
        file_name is optional (used for labelling source when file_input is a buffer)
        """
        if isinstance(file_input, (str, bytes, os.PathLike)):
            file_name = os.path.basename(file_input)
            read_target = file_input
        else:
            read_target = file_input  # assume file-like
            if file_name is None:
                file_name = "uploaded_file.xlsx"

        # read once without headers to find header row
        try:
            if hasattr(read_target, "seek"):
                read_target.seek(0)
            raw_data = pd.read_excel(read_target, sheet_name=0, header=None, engine="openpyxl")
        except Exception as e:
            print(f"Failed to read {file_name}: {e}")
            return pd.DataFrame()

        header_row_index = None
        for i, row in raw_data.iterrows():
            if row.count() >= 9:  # your heuristic for header
                header_row_index = i
                break

        if header_row_index is None:
            return pd.DataFrame()

        # read again with correct header row
        if hasattr(read_target, "seek"):
            read_target.seek(0)
        df = pd.read_excel(read_target, sheet_name=0, header=header_row_index, engine="openpyxl")
        columns = df.columns

        # also load hyperlinks with openpyxl
        if hasattr(read_target, "seek"):
            read_target.seek(0)
        hyperlinks = self.extract_hyperlinks(read_target)

        extracted = pd.DataFrame()
        for name, cfg in self.groups.items():
            best, _ = ColumnMatcher.find_best_match(
                columns,
                keywords=cfg["keywords"],
                priority=cfg.get("priority", []),
                avoid=cfg.get("avoid", [])
            )

            if best:
                extracted[name] = df[best]

                # special handling: if it's a photo link column, replace text with hyperlink
                if name.lower() in ["photo link", "photo", "link foto", "poza", "picture", "schita", "link",  
                                    "foto","imagine", "poza",  "picture",  "sketch", "pagina prezentare",  "schita productie","google map","photo", "google maps"]:
                                
                    col_idx = list(columns).index(best)
                    for row_idx in range(len(df)):
                        excel_row_idx = row_idx + header_row_index + 1  # adjust for skipped rows
                        if (excel_row_idx, col_idx) in hyperlinks:
                            extracted.at[row_idx, name] = hyperlinks[(excel_row_idx, col_idx)]
            else:
                extracted[name] = ""

        extracted["__source_file"] = file_name
        return extracted

    @staticmethod
    def remove_empty_columns(df, excepted_columns):
        cols_to_check = [col for col in df.columns if col != excepted_columns]
        df = df.replace(r'^\s*$', np.nan, regex=True)
        return df.dropna(how="all", subset=cols_to_check)

    @staticmethod
    def normalize_dimension(value: str):
        if pd.isna(value) or str(value).strip() == "":
            return None
        value = str(value).strip().lower()
        value = value.replace(",", ".")
        value = re.sub(r"\s*m", "", value)
        return value

    @staticmethod
    def process_size_base_height(row):
        base = DataProcessor.normalize_dimension(row.get("Base"))
        height = DataProcessor.normalize_dimension(row.get("Height"))
        size = str(row.get("Size")) if not pd.isna(row.get("Size")) else ""

        if size.strip():
            parts = re.split(r"[xX]", size)
            if len(parts) == 2:
                base, height = DataProcessor.normalize_dimension(parts[0]), DataProcessor.normalize_dimension(parts[1])
        elif base and height:
            size = f"{base}m x {height}m"

        return pd.Series({
            "Base": f"{base}m" if base else None,
            "Height": f"{height}m" if height else None,
            "Size": size if size else None
        })

    def process_dates(self, df):
        df["Start"] = df["Start"].apply(self._safe_to_date)
        df["End"]   = df["End"].apply(self._safe_to_date)
        return df

    @staticmethod
    def check_if_literal_date(value):
        pattern = r'^\d{1,2}\s+\w+\s*-\s*\d{1,2}\s+\w+$'
        return bool(re.match(pattern, str(value).strip(), re.IGNORECASE))

    @staticmethod
    def split_literal_date(value):
        dates_splitted = [" ".join((d.strip(), str(datetime.datetime.now().year)))
                          for d in str(value).split('-')]
        if len(dates_splitted) == 2:
            return dates_splitted
        return [None, None]

    def deal_with_literal_dates(self, df):
        def process(x, current_end):
            if self.check_if_literal_date(x):
                start, end = self.split_literal_date(x)
                return pd.Series([start, end])
            else:
                return pd.Series([x, current_end])

        df[["Start", "End"]] = df.apply(
            lambda row: process(row["Start"], row["End"]),
            axis=1
        )
        return df

    @staticmethod
    def _safe_to_date(value):
        try:
            dt = pd.to_datetime(value, dayfirst=True, errors="raise")
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return value  

    @staticmethod
    def calculate_no_of_months(row):
        try:
            start = pd.to_datetime(row["Start"], errors="coerce")
            end = pd.to_datetime(row["End"], errors="coerce")
            if pd.isna(start) or pd.isna(end):
                return None
            delta_days = (end - start).days
            return round(delta_days / 30, 1)
        except Exception:
            return None

    @staticmethod
    def build_gps_from_lat_long(row):
        latitude = row.get("Latitude")
        longitude = row.get("Longitude")
        gps_coordinates = row.get("GPS")

        if pd.notna(gps_coordinates) and str(gps_coordinates).strip():
            return gps_coordinates
        if pd.notna(latitude) and pd.notna(longitude):
            return f"{latitude}, {longitude}"
        return None

    def process_files(self, file_objs):
        all_data = []
        for file_input, file_name in file_objs:
            extracted = self.extract_standardized_dataframe(file_input, file_name=file_name)
            if not extracted.empty:
                all_data.append(extracted)

        if not all_data:
            return pd.DataFrame()

        final_df = pd.concat(all_data, ignore_index=True)
        final_df.dropna(how="all", inplace=True)
        final_df = self.remove_empty_columns(final_df, excepted_columns="__source_file")
        final_df[["Base", "Height", "Size"]] = final_df.apply(self.process_size_base_height, axis=1)
        final_df["GPS"] = final_df.apply(self.build_gps_from_lat_long, axis=1)
        final_df = self.process_dates(final_df)
        final_df = self.deal_with_literal_dates(final_df)
        final_df["No. of months"] = final_df.apply(self.calculate_no_of_months, axis=1)

        return final_df
