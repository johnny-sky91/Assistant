import pandas as pd
import csv, os, datetime


def create_csv_pos(
    path_excel_dat, ccn, mas_loc, request_div, pur_loc, delivery, inspection
):
    # current date
    current_date = datetime.date.today()
    formatted_date = current_date.strftime("%Y/%m/%d")
    # read po template
    data_excel = pd.read_excel(path_excel_dat, sheet_name="po_data")
    # read df with vendor codes and change it to dictionary
    vendor_code_df = pd.read_excel(path_excel_dat, sheet_name="vendor_code")
    vendors_code = vendor_code_df.set_index("VENDOR")["CODE"].to_dict()
    # create df with all pos
    all_vendors_po = pd.DataFrame(
        {
            "CCN": ccn,
            "MAS_LOC": mas_loc,
            "REQUEST_DIV": request_div,
            "PLACED_DATE": formatted_date,
            "REQUESTOR": data_excel["REQUESTOR"],
            "VENDOR": data_excel["VENDOR"],
            "PUR_LOC": pur_loc,
            "PR_ITEM": data_excel["PR_ITEM"],
            "PR_REVISION": " ",
            "REQD_DATE": data_excel["REQD_DATE"].dt.strftime("%Y/%m/%d"),
            "ORDER_QTY": data_excel["ORDER_QTY"],
            "FJ_SEIBAN": " ",
            "DELIVERY": delivery,
            "INSPECTION": inspection,
            "APPLY_SEC_CODE": " ",
        }
    )
    directory_path = os.path.dirname(path_excel_dat)

    suppliers_list = all_vendors_po["VENDOR"].unique().tolist()
    for supplier in suppliers_list:
        one_vendor_pos = all_vendors_po[all_vendors_po["VENDOR"] == supplier]
        one_vendor_pos.loc[:, "VENDOR"] = vendors_code[supplier]
        now = datetime.datetime.now()
        filename = f"PO_{supplier}_{now.strftime('%d%m%Y_%H%M')}.csv"
        one_vendor_pos["VENDOR"] = "000" + one_vendor_pos["VENDOR"].astype("string")
        new_file_path = os.path.join(directory_path, filename)
        one_vendor_pos.to_csv(new_file_path, index=False, quoting=csv.QUOTE_ALL)
