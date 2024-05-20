import pandas as pd


class DispoviewDataReader:
    def __init__(
        self,
        dispo_file_path: str,
    ):
        self.dispo_file_path = dispo_file_path
        self.raw_dispoview = None
        self.ready_dispoview = None

    def _read_dispoview(self):
        self.raw_dispoview = pd.read_excel(self.dispo_file_path)

    def _select_dispoview(self):
        columns_to_drop = [
            "SAP - Material Description",
            "SAP - Material Number (106xxx)",
            "SAP - Purchasing Group",
            "StratBuyer",
            "MRPtype",
            "SAP - MRP Controller",
            "FrameContract",
            "SAP - Planned Deliv. Time",
            "Trading Partner Code",
            "SAP - Vendor Name",
        ]
        self.ready_dispoview = self.raw_dispoview.drop(columns=columns_to_drop)

        self.ready_dispoview.rename(
            columns={"SAP - Product Number Print": "CODENUMBER", "Figure": "DATA"},
            inplace=True,
        )
        self.ready_dispoview.dropna(how="all", inplace=True)
        self.ready_dispoview.fillna(0, inplace=True)
        self.ready_dispoview = self.ready_dispoview.replace("'", "", regex=True)
        self.ready_dispoview.iloc[:, 2:] = self.ready_dispoview.iloc[:, 2:].astype(int)

    def __call__(self):
        self._read_dispoview()
        self._select_dispoview()
