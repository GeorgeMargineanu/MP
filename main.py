import pandas as pd
import glob
import os
import re
import unicodedata
import json
import numpy as np

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
    def __init__(self, groups, directory):
        self.groups = groups
        self.directory = directory

    def extract_standardized_dataframe(self, file_path):
        file_name = os.path.basename(file_path)
        print(f'Processing file: {file_name}')

        try:
            raw_data = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl')
        except Exception as e:
            print(f"Failed to read {file_name}: {e}")
            return pd.DataFrame()

        header_row_index = None
        for i, row in raw_data.iterrows():
            if row.count() >= 9:
                header_row_index = i
                break

        if header_row_index is None:
            print(f"No suitable header found in {file_name}")
            return pd.DataFrame()

        df = pd.read_excel(file_path, sheet_name=0, header=header_row_index, engine='openpyxl')
        columns = df.columns
        print(f'Cols are {columns}')

        extracted = pd.DataFrame()
        for name, cfg in self.groups.items():
            best, _ = ColumnMatcher.find_best_match(
                columns,
                keywords=cfg["keywords"],
                priority=cfg.get("priority", []),
                avoid=cfg.get("avoid", [])
            )
            extracted[name] = df[best] if best else ""

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

    def process_all_files(self):
        correct_files = [
            file for file in glob.glob(os.path.join(self.directory, '*'))
            if os.path.isfile(file) and (file.endswith('.xlsx') or file.endswith('.csv'))
        ]

        all_data = []
        for file in correct_files:
            extracted = self.extract_standardized_dataframe(file)
            if not extracted.empty:
                all_data.append(extracted)

        if not all_data:
            return pd.DataFrame()

        final_df = pd.concat(all_data, ignore_index=True)
        final_df.dropna(how="all", inplace=True)
        final_df = self.remove_empty_columns(final_df, excepted_columns="__source_file")
        final_df[["Base", "Height", "Size"]] = final_df.apply(self.process_size_base_height, axis=1)
        final_df["GPS"] = final_df.apply(self.build_gps_from_lat_long, axis=1)
        return final_df


if __name__ == "__main__":
    directory = r'\\unm-srv-nor\WorkFolders\Others\_NON TV\CORE\OFERTE FURNIZORI OOH\Facute'
    with open('groups.json', 'r') as file:
        groups = json.load(file)

    processor = DataProcessor(groups, directory)
    final_df = processor.process_all_files()
    final_df.to_excel("test.xlsx", index=False)
