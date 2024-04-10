from datetime import timedelta
import os, datetime
import pandas as pd


class ProductsDataReader:
    def __init__(
        self,
        file_path: str = None,
        folder_path: str = None,
        to_excel: bool = False,
    ):
        self.file_path = file_path
        self.folder_path = folder_path
        self.to_excel = to_excel
        self.data = None
        self.raw_data = None
        self.products = None
        self.orders = None
        self.supply = None
        self.balances = None
        self.non_date_columns = [f"COLUMN_{x+1}" for x in range(5)]

    def read_data(self):
        try:
            self.data = pd.read_excel(self.file_path)
        except Exception as e:
            print(f"Error loading data: {e}")

    def apply_new_headers(self):
        self.data = self.data.iloc[11:, 4:]
        self.data = self.data.reset_index(drop=True)
        rows_for_header = 2
        header_rows = self.data.iloc[:rows_for_header]
        new_header = header_rows.fillna(" ").astype(str).apply(" ".join, axis=0)
        self.data.columns = new_header
        self.data = self.data.iloc[rows_for_header:]
        new_columns = self.non_date_columns + self.data.columns[5:].to_list()
        self.data.columns = new_columns
        self.data = self.data.drop(2)
        self.raw_data = self.data.copy()

    def separate_components_row(self):
        column_len = len(self.data["COLUMN_5"])
        rows_per_comp = 26
        self.data["COLUMN_5"] = [
            "ROW_" + str(((i - 1) % rows_per_comp) + 1)
            for i in range(1, column_len + 1)
        ]

    def select_weeks_data_only(self):
        penultimate_column = self.data.columns[-2]
        non_date_columns = self.data.columns[:5].to_list()

        all_weeks_data = [col for col in self.data.columns if "W Total" in col]
        today = datetime.datetime.today()
        start_of_week = (today - timedelta(days=today.weekday())).strftime("%#m/%#d")
        weeks_data = []
        for i, x in enumerate(all_weeks_data):
            if start_of_week in x:
                weeks_data = weeks_data + all_weeks_data[i:]
                break
        self.data = self.data[non_date_columns + weeks_data + [penultimate_column]]

        new_weeks_dates = [
            x.split(" ")[2][:-1] if "/" in x else x for x in self.data.columns
        ]
        self.data.columns = new_weeks_dates
        self.data = self.data.rename(
            columns={self.data.columns[-1]: f"After {self.data.columns[-2]}"}
        )

    def get_products(self):
        self.products = pd.DataFrame(self.data[self.data["COLUMN_5"] == "ROW_3"])
        self.products = self.products[["COLUMN_1"]]
        self.products.rename(columns={"COLUMN_1": "SOI"}, inplace=True)
        self.products.reset_index(drop=True, inplace=True)

    def _get_data(self, row_value):
        result = pd.DataFrame(self.data[self.data["COLUMN_5"] == row_value])
        result = result.drop(columns=self.non_date_columns)
        result = result.reset_index(drop=True)
        result = result.fillna(0)
        result = pd.concat([self.products, result], axis=1)
        return result

    def get_supply(self):
        self.supply = self._get_data("ROW_5")

    def get_orders(self):
        self.orders = self._get_data("ROW_13")

    def get_balances(self):
        self.balances = self._get_data("ROW_24")

    def read_multiple_files(self):
        supply_to_concat = []
        balance_to_concat = []
        orders_to_concat = []
        for file_path in os.listdir(self.folder_path):
            full_file_path = os.path.join(self.folder_path, file_path)
            self.file_path = full_file_path
            self.read_one_file()

            supply_to_concat.append(self.supply)
            balance_to_concat.append(self.balances)
            orders_to_concat.append(self.orders)

            self.supply = pd.concat(supply_to_concat, ignore_index=True)
            self.balances = pd.concat(balance_to_concat, ignore_index=True)
            self.orders = pd.concat(orders_to_concat, ignore_index=True)

    def read_one_file(self):
        self.read_data()
        self.apply_new_headers()
        self.separate_components_row()
        self.select_weeks_data_only()
        self.get_products()
        self.get_supply()
        self.get_orders()
        self.get_balances()

    def save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Report_products_data_{now.strftime('%d%m%Y_%H%M')}.xlsx"

        directory_path = os.path.dirname(self.file_path)
        report_file_path = os.path.join(directory_path, filename)
        source_file = os.path.basename(self.file_path)

        writer = pd.ExcelWriter(report_file_path)

        # self.raw_data.to_excel(writer, sheet_name=f"raw_Data", index=True)
        # self.data.to_excel(writer, sheet_name=f"all_data", index=True)
        # self.products.to_excel(writer, sheet_name=f"SOI", index=True)

        self.supply.to_excel(writer, sheet_name=f"Supply", index=False)
        self.orders.to_excel(writer, sheet_name=f"Orders", index=False)
        self.balances.to_excel(writer, sheet_name=f"Balances", index=False)

        info_df = pd.DataFrame({"Source_file": [source_file]})
        info_df.to_excel(writer, sheet_name="INFO", index=False)

        writer._save()

    def __call__(self):
        # if os.path.isdir(self.file_path):
        if self.folder_path:
            self.read_multiple_files()
        else:
            self.read_one_file()
        if self.to_excel:
            self.save_to_excel()
