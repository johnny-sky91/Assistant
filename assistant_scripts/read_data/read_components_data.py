import os, datetime
import pandas as pd
from datetime import timedelta


class ComponentsDataReader:
    def __init__(
        self, file_path: str, fix_weeks: bool, add_forecast: bool, to_excel: bool
    ):
        self.file_path = file_path
        self.fix_weeks = fix_weeks
        self.to_excel = to_excel
        self.add_forecast = add_forecast
        self.data = None
        self.raw_data = None
        self.components = None
        self.components_stock = None
        self.orders = None
        self.demand_plan = None
        self.orders_demand_plan = None
        self.final_demand = None

    def read_data(self):
        try:
            self.data = pd.read_excel(self.file_path)
        except Exception as e:
            print(f"Error loading data: {e}")

    def apply_new_headers(self):
        self.data = self.data.iloc[7:, 11:]
        self.data = self.data.reset_index(drop=True)
        rows_for_header = 2
        header_rows = self.data.iloc[:rows_for_header]
        new_header = header_rows.fillna(" ").astype(str).apply(" ".join, axis=0)
        self.data.columns = new_header
        self.data = self.data.iloc[rows_for_header:]
        first_column_names = [f"COLUMN_{x+1}" for x in range(10)]
        new_columns = first_column_names + self.data.columns[10:].to_list()
        self.data.columns = new_columns
        self.data = self.data.iloc[5:, :]

    def separate_components_row(self):
        column_len = len(self.data["COLUMN_8"])
        rows_per_comp = 24
        self.data["COLUMN_8"] = [
            "ROW_" + str(((i - 1) % rows_per_comp) + 1)
            for i in range(1, column_len + 1)
        ]

    def select_weeks_data_only(self):
        all_weeks_data = [col for col in self.data.columns if "W Total" in col]

        today = datetime.datetime.today()
        start_of_week = today - timedelta(days=today.weekday())
        start_of_week_str = start_of_week.strftime("%#m/%#d")

        for i, x in enumerate(all_weeks_data):
            if start_of_week_str in x:
                weeks_data = all_weeks_data[i:]
                break

        non_date_columns = self.data.columns[:10].to_list()

        self.data = self.data[non_date_columns + weeks_data]

        new_weeks_dates = [
            x.split(" ")[2][:-1] if "/" in x else x for x in self.data.columns
        ]

        self.data.columns = new_weeks_dates
        self.raw_data = self.data.copy()

    def get_components(self):
        self.components = pd.DataFrame(self.data[self.data["COLUMN_8"] == "ROW_12"])
        self.components = self.components[["COLUMN_1"]]
        self.components.rename(columns={"COLUMN_1": "COMPONENT"}, inplace=True)
        self.components["COMPONENT"] = self.components["COMPONENT"].str.replace(" ", "")
        self.components.reset_index(drop=True, inplace=True)

    def get_components_stock(self):
        warehouse_stock = pd.DataFrame(self.data[self.data["COLUMN_8"] == "ROW_6"])
        warehouse_stock = warehouse_stock[["COLUMN_2"]].rename(
            columns={"COLUMN_2": "WAREHOUSE_STOCK"}
        )
        warehouse_stock.reset_index(drop=True, inplace=True)
        warehouse_stock = warehouse_stock.astype(int)

        factory_stock = pd.DataFrame(self.data[self.data["COLUMN_8"] == "ROW_8"])
        factory_stock = factory_stock[["COLUMN_2"]].rename(
            columns={"COLUMN_2": "FACTORY_STOCK"}
        )

        factory_stock["FACTORY_STOCK"] = factory_stock["FACTORY_STOCK"].apply(
            lambda x: int(str(x)[str(x).find("s") + 1 :]) if "s" in str(x) else 0
        )
        factory_stock.reset_index(drop=True, inplace=True)

        self.components_stock = pd.concat(
            [self.components, warehouse_stock, factory_stock], axis=1
        )
        self.components_stock["TOTAL_STOCK"] = (
            self.components_stock["WAREHOUSE_STOCK"]
            + self.components_stock["FACTORY_STOCK"]
        )

    def get_orders(self):
        self.orders = pd.DataFrame(self.data[self.data["COLUMN_8"] == "ROW_2"])
        self.orders = self.orders.iloc[:, 10:]
        self.orders.reset_index(drop=True, inplace=True)
        self.orders.fillna(0, inplace=True)
        self.orders = pd.concat([self.components, self.orders], axis=1)

    def get_demand_plan(self):
        self.demand_plan = pd.DataFrame(self.data[self.data["COLUMN_8"] == "ROW_1"])
        self.demand_plan = self.demand_plan.iloc[:, 10:]
        self.demand_plan.reset_index(drop=True, inplace=True)
        self.demand_plan.fillna(0, inplace=True)
        self.demand_plan = pd.concat([self.components, self.demand_plan], axis=1)

    def join_orders_demand_plan(self):
        self.orders["DATA"] = "ORDERS"
        self.orders.insert(1, "DATA", self.orders.pop("DATA"))

        self.demand_plan["DATA"] = "DEMAND_PLAN"
        self.demand_plan.insert(1, "DATA", self.demand_plan.pop("DATA"))

        self.orders_demand_plan = pd.concat([self.orders, self.demand_plan])

    def split_orders_demand_plan(self):
        self.orders = pd.DataFrame(
            self.orders_demand_plan[self.orders_demand_plan["DATA"] == "ORDERS"]
        )
        self.orders.drop("DATA", axis=1, inplace=True)
        self.demand_plan = pd.DataFrame(
            self.orders_demand_plan[self.orders_demand_plan["DATA"] == "DEMAND_PLAN"]
        )
        self.demand_plan.drop("DATA", axis=1, inplace=True)

    def fix_no_valid_weeks(self):
        # get first days of next X weeks
        current_date = datetime.datetime.now()
        cw_first_day = current_date - timedelta(days=current_date.weekday())
        weeks = 20
        next_weeks_dates = [cw_first_day + timedelta(weeks=i) for i in range(weeks)]
        next_weeks = [date.strftime("%#m/%#d") for date in next_weeks_dates]
        # get only weeks from given data
        weeks_data = self.orders_demand_plan.iloc[:, 2:]
        weeks_data_columns = weeks_data.columns.to_list()
        new_weeks_data_columns = weeks_data_columns[
            : weeks_data_columns.index(next_weeks[-1]) + 2
        ]
        # select only choosen weeks
        weeks_data = weeks_data[new_weeks_data_columns]
        # add no valid week to previoous week and then remove them
        columns_to_drop = []
        for i, column_name in enumerate(weeks_data.columns):
            if column_name not in next_weeks:
                previous_column = weeks_data.columns[i - 1]
                weeks_data[previous_column] = (
                    weeks_data[previous_column] + weeks_data[column_name]
                )
                columns_to_drop.append(column_name)
        weeks_data.drop(columns=columns_to_drop, inplace=True)
        self.orders_demand_plan = pd.concat(
            [self.orders_demand_plan.iloc[:, :2], weeks_data], axis=1, join="inner"
        )

    def prepare_forecast(self):
        self.final_demand = self.demand_plan.copy()
        self.final_demand.iloc[:, 1] = self.orders.iloc[:, 1].values
        self.final_demand["stock"] = self.components_stock["TOTAL_STOCK"]
        self.final_demand.insert(1, "stock", self.final_demand.pop("stock"))

    def save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Report_components_data_{now.strftime('%d%m%Y_%H%M')}.xlsx"
        directory_path = os.path.dirname(self.file_path)
        report_file_path = os.path.join(directory_path, filename)
        source_file = os.path.basename(self.file_path)

        writer = pd.ExcelWriter(report_file_path)
        if self.add_forecast:
            self.final_demand.to_excel(writer, sheet_name=f"Sheet1", index=False)
        self.components_stock.to_excel(
            writer, sheet_name=f"Components_stock", index=False
        )
        self.orders.to_excel(writer, sheet_name=f"Orders", index=False)
        self.demand_plan.to_excel(writer, sheet_name=f"Demand_plan", index=False)
        # self.raw_data.to_excel(writer, sheet_name=f"Raw_data", index=True)

        info_df = pd.DataFrame({"Source_file": [source_file]})
        info_df.to_excel(writer, sheet_name="INFO", index=False)

        writer._save()

    def __call__(self):
        self.read_data()
        self.apply_new_headers()
        self.separate_components_row()
        self.select_weeks_data_only()
        self.get_components()
        self.get_components_stock()
        self.get_orders()
        self.get_demand_plan()
        self.join_orders_demand_plan()
        if self.fix_weeks:
            self.fix_no_valid_weeks()
        self.split_orders_demand_plan()
        if self.add_forecast:
            self.prepare_forecast()
        if self.to_excel:
            self.save_to_excel()
        return self
