import os, datetime
import numpy as np
import pandas as pd


class CreateSupplyInfo:
    def __init__(
        self,
        path_supply: str,
        path_components: str,
        supplier1: str,
        supplier2: str,
        supplier3: str,
        incoterms: str,
        t_mode: str,
    ):
        self.path_supply = path_supply
        self.path_components = path_components
        self.supplier1 = supplier1.split(",")
        self.supplier2 = supplier2.split(",")
        self.supplier3 = supplier3.split(",")
        self.incoterms = incoterms
        self.t_mode = t_mode
        self.supply_info_raw = None
        self.components_info = None
        self.supply_info = None

    def read_supply_data(self):
        try:
            self.supply_info_raw = pd.read_excel(
                self.path_supply, sheet_name="current_info"
            )
            self.components_info = pd.read_excel(self.path_components)
        except Exception as e:
            print(f"Error loading data: {e}")

    def drop_columns(self):
        self.supply_info = self.supply_info_raw.copy()
        columns_to_drop = ["GROUP", "STATUS", "SHIPMENT_ID", "COMMENT"]
        self.supply_info.drop(columns_to_drop, axis=1, inplace=True)
        self.supply_info = self.supply_info.loc[self.supply_info["QTY"] != 0]

    def rename_suppliers(self):
        mapping_conditions = [
            self.supply_info["SUPPLIER"].str.contains(
                self.supplier1[0], case=False, na=False
            ),
            self.supply_info["SUPPLIER"].str.contains(
                self.supplier2[0], case=False, na=False
            ),
            self.supply_info["SUPPLIER"].str.contains(
                self.supplier3[0], case=False, na=False
            ),
        ]
        mapping_values = [self.supplier1[1], self.supplier2[1], self.supplier3[1]]
        self.supply_info["SUPPLIER"] = np.select(
            mapping_conditions,
            mapping_values,
            default=self.supply_info["SUPPLIER"],
        )

    def calucate_dates(self):
        self.supply_info.rename(
            columns={
                "ETD_DATE_WEEK": "Confirmed Delivery Date",
                "QTY": "Confirmed Quantity",
            },
            inplace=True,
        )
        self.supply_info["Confirmed Delivery Date"] = pd.to_datetime(
            self.supply_info["Confirmed Delivery Date"]
        ) + pd.Timedelta(days=3)
        self.supply_info["Confirmed Delivery Week"] = (
            self.supply_info["Confirmed Delivery Date"].dt.isocalendar().week
        )

    def merge_component_info(self):
        self.supply_info = pd.merge(
            self.supply_info, self.components_info, how="left", on="COMPONENT"
        )
        self.supply_info = self.supply_info.drop_duplicates()

        self.supply_info["Codenumber"].fillna("NOT_FOUND", inplace=True)
        self.supply_info["Part number"].fillna("NOT_FOUND", inplace=True)
        self.supply_info["Description"].fillna("NOT_FOUND", inplace=True)

    def add_more_info(self):
        self.supply_info["Incoterms"] = self.incoterms
        self.supply_info["Transportation Mode"] = self.t_mode

    def final_update(self):
        new_column_order = [
            "Codenumber",
            "Description",
            "Confirmed Quantity",
            "Confirmed Delivery Date",
            "Confirmed Delivery Week",
            "SUPPLIER",
            "Incoterms",
            "Transportation Mode",
        ]

        self.supply_info = self.supply_info[new_column_order]
        self.supply_info.rename(columns={"Codenumber": "COMPONENT"}, inplace=True)

    def save_to_excel(self):
        now = datetime.datetime.now()
        filename = f"mem_supply_{now.strftime('%d%m%Y')}.xlsx"
        directory_path = os.path.dirname(self.path_supply)
        report_file_path = os.path.join(directory_path, filename)
        writer = pd.ExcelWriter(report_file_path)
        self.supply_info.to_excel(writer, sheet_name=f"Sheet1", index=False)
        writer._save()

    def __call__(self):
        self.read_supply_data()
        self.drop_columns()
        self.rename_suppliers()
        self.calucate_dates()
        self.merge_component_info()
        self.add_more_info()
        self.final_update()
        self.save_to_excel()
