import pandas as pd
import datetime, os, openpyxl
from datetime import timedelta
from openpyxl.worksheet.datavalidation import DataValidation


class SupplyDataReader:
    def __init__(self, path_supply: str = None, path_groups: str = None):
        self.path_supply = path_supply
        self.path_groups = path_groups
        self.data = None
        self.groups = None
        self.supply_info = None

    def read_data(self):
        try:
            self.data = pd.read_excel(self.path_supply, sheet_name=None)
            sheet_names = list(self.data.keys())
            self.data = self.data[sheet_names[1]]
        except Exception as e:
            print(f"Error loading data: {e}")

    def read_groups(self):
        try:
            self.groups = pd.read_excel(self.path_groups, sheet_name="groups")
        except Exception as e:
            print(f"Error loading data: {e}")

    def prepare_data(self):
        self.data = self.data.iloc[1:, 4:]
        self.data = self.data.reset_index(drop=True)
        headers_row = 0
        new_headers = self.data.iloc[headers_row].tolist()
        self.data = self.data.rename(columns=dict(zip(self.data.columns, new_headers)))
        self.data = self.data.drop(self.data.index[0])
        self.data = self.data.rename(columns={self.data.columns[0]: "COMPONENT"})

    def select_data(self):
        ready_columns = ["COMPONENT", "factory\n(destination)", "date", "Supplier"]
        date_columns = [
            x for x in self.data.columns if isinstance(x, datetime.datetime)
        ]

        today = datetime.datetime.today()
        today = today.replace(hour=0, minute=0, second=0, microsecond=0)
        start_of_week = today - timedelta(days=today.weekday())

        date_columns = date_columns[date_columns.index(start_of_week) :]
        self.data = self.data[ready_columns + date_columns]

    def filter_data(self):
        self.data = pd.DataFrame(
            self.data[
                (self.data["date"] == "supply: C")
                & (self.data["factory\n(destination)"] == "CZ")
            ]
        )
        self.data.rename(columns={"Supplier": "SUPPLIER"}, inplace=True)
        self.data.drop(columns=["date", "factory\n(destination)"], inplace=True)
        self.data.fillna(0, inplace=True)

    def add_groups(self):
        self.groups = self.groups[["COMPONENT", "GROUP"]].drop_duplicates()
        self.groups.reset_index(inplace=True, drop=True)
        self.supply_info = pd.merge(self.data, self.groups, on="COMPONENT", how="inner")
        self.supply_info.insert(1, "GROUP", self.supply_info.pop("GROUP"))

    def melt_supply_data(self):
        self.supply_info = self.supply_info.melt(
            id_vars=["GROUP", "COMPONENT", "SUPPLIER"],
            var_name="ETD_DATE_WEEK",
            value_name="QTY",
        )
        self.supply_info = self.supply_info.loc[self.supply_info["QTY"] != 0]
        self.supply_info.reset_index(inplace=True, drop=True)
        self.supply_info[["STATUS", "SHIPMENT_ID", "COMMENT"]] = None

    def save_to_excel(self):
        current_datetime = datetime.datetime.now()

        current_year = current_datetime.strftime("%Y")
        current_week = current_datetime.strftime("%V")
        filename = f"new_Components_supply_CW{current_week}_{current_year}.xlsx"
        directory_path = os.path.dirname(self.path_supply)
        report_file_path = os.path.join(directory_path, filename)
        source_file = os.path.basename(self.path_supply)

        writer = pd.ExcelWriter(report_file_path)

        self.supply_info.to_excel(writer, sheet_name=f"current_info", index=False)

        info_df = pd.DataFrame({"Source_file": [source_file]})
        info_df.to_excel(writer, sheet_name="INFO", index=False)

        writer._save()
        add_dropdown_statuses(
            excel_filename=report_file_path, sheet_name="current_info"
        )

    def __call__(self):
        self.read_data()
        self.read_groups()
        self.prepare_data()
        self.select_data()
        self.filter_data()
        self.add_groups()
        self.melt_supply_data()
        self.save_to_excel()


def add_dropdown_statuses(excel_filename: str, sheet_name: str):
    workbook = openpyxl.load_workbook(excel_filename)
    sheet = workbook[sheet_name]

    choices = ["Shipped", "Open_WH", "Delivered", "Other"]
    choices_ready = '"' + ",".join(choices) + '"'

    data_val = DataValidation(type="list", formula1=choices_ready, allow_blank=True)

    sheet.add_data_validation(data_val)
    data_val_range = "F2:F" + str(sheet.max_row)
    data_val.add(data_val_range)

    workbook.save(excel_filename)
