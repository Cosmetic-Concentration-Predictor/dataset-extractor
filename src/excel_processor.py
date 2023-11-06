import os
import pandas as pd
import xlrd
import re

ROW_AVERAGE = 16


class ExcelProcessor:
    def __init__(self, input_folder, output_csv, output_txt, processing_data=False):
        self.input_folder = input_folder
        self.output_csv = output_csv
        self.output_txt = open(output_txt, "w")
        self.processing_data = processing_data
        self.total_sheets = 0
        self.total_unread_sheets = 0

    def process_excel_files(self):
        result_df = pd.DataFrame(columns=["acronym", "code", "concentration"])

        for filename in os.listdir(self.input_folder):
            # print(filename)
            if filename.lower().endswith(".xls"):
                file_path = os.path.join(self.input_folder, filename)
                result_df = pd.concat(
                    [result_df, self.process_xls(file_path)], ignore_index=True
                )
            # else:
            #     print(filename)
        print(f"Total Unread Sheets: {self.total_unread_sheets}")

        print(self.total_sheets)
        result_df.to_csv(self.output_csv, index=False)

    def process_xls(self, file_path):
        result_df = None
        wb = xlrd.open_workbook(file_path)

        for sheet in wb.sheets():
            start_row, start_col, end_row, _ = self.find_table_range_xls(sheet)

            if start_row is not None and start_col is not None:
                table_data = []
                for row_num in range(start_row, end_row):
                    if (
                        "PHASE" not in str(sheet.cell_value(row_num, start_col))
                        or " " not in str(sheet.cell_value(row_num, start_col))
                        or "Code" not in str(sheet.cell_value(row_num, start_col))
                        or "Pour" not in str(sheet.cell_value(row_num, start_col))
                    ):
                        row_data = [
                            sheet.cell_value(row_num, start_col),
                            sheet.cell_value(row_num, start_col + 2),
                        ]
                        table_data.append(row_data)
                    else:
                        continue

                table_df = pd.DataFrame(table_data, columns=["Code", "Pour 1"])
                table_df.dropna(inplace=True)

                table_df["Pour 1"] = (
                    table_df["Pour 1"]
                    .replace("%", "", regex=True)
                    .apply(pd.to_numeric, errors="coerce")
                )

                table_df["acronym"] = table_df["Code"].str.extract(
                    r"([A-Za-z]+\s?(?=\d)(?:\d+[A-Za-z]+)?)"
                )
                table_df["code"] = table_df["Code"].str.extract(
                    r"(\s?\d*(?!.*[A-Za-z]))"
                )

                table_df["concentration"] = table_df["Pour 1"].astype(float)

                table_df.drop("Code", axis=1, inplace=True)
                table_df.drop("Pour 1", axis=1, inplace=True)
                table_df["acronym"] = table_df[table_df["acronym"].str.len() > 1][
                    "acronym"
                ]
                table_df["acronym"] = table_df["acronym"].str.rstrip()
                table_df.dropna(inplace=True)

                df_sum = table_df["concentration"].sum()
                if df_sum > 1.8:
                    table_df["concentration"] = table_df["concentration"].div(100)

                table_df["acronym"] = table_df["acronym"].apply(
                    lambda x: re.sub(
                        "^(§|MY|ME|MT|AMT|LABOMT|LABOME|LABOAMT|LABO|LAB)",
                        "",
                        x.upper(),
                    )
                )

                try:
                    table_df["code"] = table_df["code"].astype(int)
                except ValueError:
                    print(f"{file_path} - {sheet.name}")
                    continue

                self.output_txt.write(f"{file_path}: {sheet.name}\n{table_df}\n\n")

                if not table_df.empty:
                    if self.processing_data:
                        table_df = self.add_missing_rows(table_df)
                        table_df = self.rotate_dataframe(table_df)
                    result_df = pd.concat([result_df, table_df], ignore_index=True)
                    result_df.dropna(inplace=True)
                    self.total_sheets += 1
            else:
                self.total_unread_sheets += 1
                print(f"{file_path}: {sheet.name}")

        return result_df

    def find_table_range_xls(self, sheet):
        for row_num in range(1, sheet.nrows):
            for col_num in range(0, sheet.ncols - 2):
                if (
                    (
                        str(sheet.cell_value(row_num, col_num)).lower() == "code"
                        and str(sheet.cell_value(row_num, col_num)).lower()
                        != "code action"
                    )
                    and (
                        sheet.cell_value(row_num, col_num + 2) == "Pour 1"
                        or sheet.cell_value(row_num, col_num + 2) == "%global"
                        or sheet.cell_value(row_num, col_num + 2) == "Qtté th (g)"
                    )
                ) or (
                    (
                        sheet.cell_value(row_num, col_num) == ""
                        or sheet.cell_value(row_num, col_num) == "N°MP"
                    )
                    and sheet.cell_value(row_num, col_num + 2) == "%"
                ):
                    start_row, start_col = row_num, col_num
                    end_row, end_col = sheet.nrows, col_num + 2
                    return start_row, start_col, end_row, end_col

        return None, None, None, None

    def add_missing_rows(self, df, target_rows=ROW_AVERAGE):
        current_rows = len(df)

        if current_rows >= target_rows:
            return df

        rows_to_add = target_rows - current_rows
        rows_to_add_df = pd.DataFrame([[0, 0, 0]] * rows_to_add, columns=df.columns)

        df = pd.concat([df, rows_to_add_df], ignore_index=True)

        return df

    def rotate_dataframe(self, df):
        num_rows = len(df)
        rotated_dataframes = []

        for i in range(num_rows):
            rotated_df = df.copy()

            rotated_values = df.values.copy()
            rotated_values = rotated_values[-i:].tolist() + rotated_values[:-i].tolist()
            rotated_df.iloc[:, :] = rotated_values

            rotated_dataframes.append(rotated_df)

        return pd.concat(rotated_dataframes)
