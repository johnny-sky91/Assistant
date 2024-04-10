import os, datetime
import pandas as pd
from assistant_scripts.read_data.read_companion_data import CompanionDbReader


class FactoryDataReader:
    def __init__(self, factory_data_path, db_path):
        self.factory_data_path = factory_data_path
        self.db_path = db_path
        self.active_components = None
        self.factory_data = None
        self.active_components_data = None
        self.final_data = None

    def get_companion_info(self):
        companion_info = CompanionDbReader(path_db=self.db_path)
        companion_info()
        self.active_components = pd.DataFrame(
            companion_info.components[companion_info.components["STATUS"] == "Active"]
        )
        self.active_components = self.active_components[["COMPONENT"]]

    def get_factroy_data(self):
        extension = self.factory_data_path.split(".")[-1].lower()
        # check what extension file have
        if extension == "txt":
            self.factory_data = pd.read_csv(self.factory_data_path, delimiter="\t")
            self.factory_data["Available Qty"] = self.factory_data[
                "Available Qty"
            ].str.replace(",", "")
        else:
            self.factory_data = pd.read_excel(self.factory_data_path)
        # qty to int type
        self.factory_data["Available Qty"] = self.factory_data["Available Qty"].astype(
            int
        )

    def get_active_components_data(self):
        # get active components
        self.active_components_data = self.factory_data[
            self.factory_data["FJJ P/N"].isin(self.active_components["COMPONENT"])
        ]
        # get desired data columns
        self.active_components_data = self.active_components_data[
            ["FJJ P/N", "SLoc", "Available Qty"]
        ]
        # rename column
        self.active_components_data = self.active_components_data.rename(
            columns={"FJJ P/N": "COMPONENT"}
        )
        # what are desired good parts sloc
        list_good_codes = [
            "W101",
            "W10A",
            "W191",
            "W1LA",
            "W1PA",
            "W104",
            "W2Y3",
            "WIP",
            "W1L0",
        ]
        # get bool from sloc codes
        self.active_components_data["SLoc"] = (
            self.active_components_data["SLoc"].isin(list_good_codes).astype(bool)
        )

    def group_active_components(self):
        # group components on false and true
        grouped_mem = (
            self.active_components_data.groupby(["COMPONENT", "SLoc"])["Available Qty"]
            .sum()
            .reset_index()
        )
        # make pivot
        pivot_mem = grouped_mem.pivot_table(
            index="COMPONENT", columns="SLoc", values="Available Qty", fill_value=0
        ).reset_index()
        # rename pivot columns names
        pivot_mem.columns = ["COMPONENT", "Not_good_qty", "Good_qty"]
        # merge mem comp with pivot
        ready_wst_data = self.active_components.merge(
            pivot_mem, on="COMPONENT", how="left"
        ).fillna(0)
        # reoder columns
        desired_order = ["COMPONENT", "Good_qty", "Not_good_qty"]
        self.final_data = ready_wst_data[desired_order]

    def save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"Report_factory_inventory_{now.strftime('%d%m%Y_%H%M')}.xlsx"
        directory_path = os.path.dirname(self.factory_data_path)
        report_file_path = os.path.join(directory_path, filename)
        writer = pd.ExcelWriter(report_file_path)

        self.final_data.to_excel(writer, sheet_name=f"Ready_data", index=False)
        self.active_components_data.to_excel(
            writer, sheet_name=f"Active_components_data", index=False
        )
        self.factory_data.to_excel(writer, sheet_name=f"Factory_data", index=False)

        writer._save()

    def __call__(self):
        self.get_companion_info()
        self.get_factroy_data()
        self.get_active_components_data()
        self.group_active_components()
        self.save_to_excel()
