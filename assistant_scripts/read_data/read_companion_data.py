import os, sqlite3
import pandas as pd


class CompanionDbReader:
    def __init__(self, path_db):
        self.path_db = path_db
        self.db_conn = None
        self.components = None
        self.products = None
        self.groups = None

    def _check_data_type(self):
        _, file_extension = os.path.splitext(self.path_db)
        return file_extension.lower() == ".db"

    def get_excel_data(self):
        data = pd.read_excel(self.path_db, sheet_name=None)
        self.components = data["Component_info"]
        self.products = data["SOI_info"]
        self.groups = data["Group_info"]

    def _connect_db(self):
        self.db_conn = sqlite3.connect(self.path_db)

    def _unconnect_db(self):
        self.db_conn.close()

    def get_groups(self):
        query_group = "SELECT * FROM my_group"
        query_group_comments = "SELECT * FROM my_group_comment"
        query_group_products = "SELECT * FROM my_group_product"

        groups = pd.read_sql(query_group, self.db_conn)
        groups = groups.rename(columns={"id": "my_group_id"})

        groups_comments = pd.read_sql(query_group_comments, self.db_conn)
        groups_comments = groups_comments.rename(columns={"product_id": "my_group_id"})

        group_products = pd.read_sql(query_group_products, self.db_conn)
        group_products["my_group_id"] = group_products["my_group_id"].astype("Int64")

        self.groups = pd.merge(
            groups,
            group_products,
            left_on="my_group_id",
            right_on="my_group_id",
            how="left",
        )
        last_comment_index = groups_comments.groupby("my_group_id")[
            "timestamp"
        ].idxmax()
        last_comments = groups_comments.loc[
            last_comment_index, ["my_group_id", "text", "timestamp"]
        ]

        self.groups = pd.merge(self.groups, last_comments, on="my_group_id", how="left")
        self.groups["timestamp"] = pd.to_datetime(self.groups["timestamp"])
        self.groups["timestamp"] = self.groups["timestamp"].dt.strftime(
            "%d/%m/%y - %H:%M"
        )
        self.groups = self.groups.rename(columns={"name": "Group"})

    def get_products(self):
        query_soi = "SELECT * FROM soi"
        query_soi_comments = "SELECT * FROM soi_comment"

        products = pd.read_sql(query_soi, self.db_conn)
        products = products.rename(columns={"id": "product_id"})

        products_comments = pd.read_sql(query_soi_comments, self.db_conn)
        products = products.rename(columns={"id": "comment_id"})

        merged_data = pd.merge(
            products,
            products_comments,
            left_on="product_id",
            right_on="product_id",
            how="left",
        )

        soi_agg_data = {
            "product_id": "last",
            "status": "last",
            "note": "last",
            "check": "last",
            "dummy": "last",
            "text": "last",
            "timestamp": "last",
        }
        self.products = merged_data.groupby("name").agg(soi_agg_data).reset_index()

        self.products["timestamp"] = pd.to_datetime(self.products["timestamp"])
        self.products["timestamp"] = self.products["timestamp"].dt.strftime(
            "%d/%m/%y - %H:%M"
        )
        self.products = self.products.rename(columns={"name": "SOI"})

        self.products = pd.merge(
            self.products,
            self.groups[["Group", "soi_id"]],
            left_on="product_id",
            right_on="soi_id",
            how="left",
        )
        self.products.drop(["product_id", "soi_id"], axis=1, inplace=True)

    def get_components(self):
        query_comps = "SELECT * FROM component"
        query_comps_comments = "SELECT * FROM component_comment"

        components = pd.read_sql(query_comps, self.db_conn)
        components = components.rename(columns={"id": "product_id"})

        components_comments = pd.read_sql(query_comps_comments, self.db_conn)
        components_comments = components_comments.rename(columns={"id": "comment_id"})

        merged_data = pd.merge(
            components,
            components_comments,
            left_on="product_id",
            right_on="product_id",
            how="left",
        )
        comp_agg_data = {
            "product_id": "last",
            "status": "last",
            "note": "last",
            "check": "last",
            "supplier": "last",
            "text": "last",
            "timestamp": "last",
        }

        self.components = merged_data.groupby("name").agg(comp_agg_data).reset_index()
        self.components["timestamp"] = pd.to_datetime(self.components["timestamp"])
        self.components["timestamp"] = self.components["timestamp"].dt.strftime(
            "%d/%m/%y - %H:%M"
        )
        self.components = self.components.rename(columns={"name": "COMPONENT"})

        self.components = pd.merge(
            self.components,
            self.groups[["Group", "component_id"]],
            left_on="product_id",
            right_on="component_id",
            how="left",
        )
        self.components.drop(["product_id", "component_id"], axis=1, inplace=True)

    def _clear_data(self):
        self.groups.drop(
            ["my_group_id", "id", "soi_id", "component_id"], axis=1, inplace=True
        )
        self.groups = self.groups.drop_duplicates()
        self.components = self.components.drop_duplicates()
        self.products = self.products.drop_duplicates()

        self.groups.columns = self.groups.columns.str.upper()
        self.components.columns = self.components.columns.str.upper()
        self.products.columns = self.products.columns.str.upper()

    def __call__(self):
        if self._check_data_type():
            self._connect_db()
            self.get_groups()
            self.get_components()
            self.get_products()
            self._clear_data()
            self._unconnect_db()
        else:
            self.get_excel_data()
