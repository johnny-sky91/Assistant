import PySimpleGUI as sg
import pyperclip, os
from dotenv import load_dotenv

from assistant_scripts.read_data.read_components_data import ComponentsDataReader
from assistant_scripts.read_data.read_products_data import ProductsDataReader
from assistant_scripts.read_data.read_factory_data import FactoryDataReader
from assistant_scripts.read_data.read_supply_data import SupplyDataReader

from assistant_scripts.reports.report_components_balances import ComponentsBalances
from assistant_scripts.reports.report_groups_statuses import GroupsStatuses
from assistant_scripts.reports.report_supply_info import CreateSupplyInfo

from assistant_scripts.other_functions.create_pos import create_csv_pos
from assistant_scripts.other_functions.sort_my_data import sort_my_data

load_dotenv()

paths = {
    "po_template": os.getenv("PATH_PO_TEMPLATE"),
    "groups": os.getenv("PATH_GROUPS"),
    "db": os.getenv("PATH_DB"),
    "my_data": os.getenv("PATH_MY_DATA"),
    "products_files": os.getenv("PATH_PRODUCTS"),
    "components_info": os.getenv("PATH_COMPONENTS"),
}

po_data = {
    "ccn": os.getenv("PO_CCN"),
    "mas_loc": os.getenv("PO_MAS_LOC"),
    "request_div": os.getenv("PO_REQUEST_DIV"),
    "pur_loc": os.getenv("PO_PUR_LOC"),
    "delivery": os.getenv("PO_DELIVERY"),
    "inspection": os.getenv("PO_INSPECTION"),
}

supply_data = {
    "supplier_1": os.getenv("SUPPLIER_1"),
    "supplier_2": os.getenv("SUPPLIER_2"),
    "supplier_3": os.getenv("SUPPLIER_3"),
    "incoterms": os.getenv("SUPPLY_INCOTERMS"),
    "t_mode": os.getenv("SUPPLY_T_MODE"),
}

passwords = {}
for i in range(1, 6):
    passwords[f"name_{i}"] = os.getenv(f"PASSWORD_{i}").split(", ")[0]
    passwords[f"pass_{i}"] = os.getenv(f"PASSWORD_{i}").split(", ")[1]

sg.theme("GreenTan")


def handle_report_components_data(values):
    data_handler = ComponentsDataReader(
        file_path=values["components_data_path"],
        fix_weeks=True,
        add_forecast=True,
        to_excel=True,
    )
    data_handler()


def handle_report_products_data(values):
    data_handler = ProductsDataReader(
        file_path=None, folder_path=paths["products_files"], to_excel=True
    )
    data_handler()


def handle_report_factory_data(values):
    data_handler = FactoryDataReader(
        factory_data_path=values["factory_path"], db_path=paths["db"]
    )
    data_handler()


def handle_create_pos(values):
    create_csv_pos(
        path_excel_dat=paths["po_template"],
        ccn=po_data["ccn"],
        mas_loc=po_data["mas_loc"],
        request_div=po_data["request_div"],
        pur_loc=po_data["pur_loc"],
        delivery=po_data["delivery"],
        inspection=po_data["inspection"],
    )


def handle_sort_my_data(values):
    sort_my_data(directory=paths["my_data"])


def handle_get_password(values, password):
    pyperclip.copy(password)


def handle_groups_statuses(values):
    report = GroupsStatuses(
        path_components=values["components_data_path"],
        path_products=paths["products_files"],
        path_groups=paths["groups"],
        path_supply=values["supply_path"],
        path_db=paths["db"],
    )
    report()


def handle_components_balances(values):
    report = ComponentsBalances(
        path_groups=paths["groups"],
        path_components=values["components_data_path"],
        path_supply=values["supply_path"],
    )
    report()


def handle_report_supply_info(values):
    report = CreateSupplyInfo(
        path_supply=values["supply_path"],
        path_components=paths["components_info"],
        supplier1=supply_data["supplier_1"],
        supplier2=supply_data["supplier_2"],
        supplier3=supply_data["supplier_3"],
        incoterms=supply_data["incoterms"],
        t_mode=supply_data["t_mode"],
    )
    report()


def handle_new_supply_info(values):
    supply = SupplyDataReader(
        path_supply=values["supply_path"], path_groups=paths["groups"]
    )
    supply()


main_functions_layout = [
    [
        sg.Column(
            [
                [sg.Text("Files browser", font=20)],
                [
                    sg.Text("Components_data_path:", size=(20, 1)),
                    sg.InputText(key="components_data_path", size=(20, 1)),
                    sg.FileBrowse(),
                ],
                [
                    sg.Text("Products_data_path:", size=(20, 1)),
                    sg.InputText(key="products_data_path", size=(20, 1)),
                    sg.FileBrowse(),
                ],
                [
                    sg.Text("Supply_file_path:", size=(20, 1)),
                    sg.InputText(key="supply_path", size=(20, 1)),
                    sg.FileBrowse(),
                ],
                [
                    sg.Text("Factory_data_path:", size=(20, 1)),
                    sg.InputText(key="factory_path", size=(20, 1)),
                    sg.FileBrowse(),
                ],
            ],
            element_justification="left",
            vertical_alignment="top",
        ),
        sg.Column(
            [
                [sg.Text("Reports", font=20)],
                [sg.Button("Groups_statuses", size=(20, 1))],
                [sg.Button("Groups_balances", size=(20, 1))],
                [sg.Button("Components_data", size=(20, 1))],
                [sg.Button("Products_data", size=(20, 1))],
                [sg.Button("Factory_data", size=(20, 1))],
                [sg.Button("Supply_info", size=(20, 1))],
            ],
            element_justification="left",
            vertical_alignment="top",
        ),
        sg.Column(
            [
                [sg.Text("Others", font=20)],
                [sg.Button("Create_POs", size=(20, 1))],
                [sg.Button("Sort_My_Data", size=(20, 1))],
                [sg.Button("New_supply_info", size=(20, 1))],
            ],
            element_justification="top",
            vertical_alignment="top",
        ),
        sg.Column(
            [
                [sg.Text("Password manager", font=20)],
                [sg.Button(passwords["name_1"], size=(20, 1))],
                [sg.Button(passwords["name_2"], size=(20, 1))],
                [sg.Button(passwords["name_3"], size=(20, 1))],
                [sg.Button(passwords["name_4"], size=(20, 1))],
                [sg.Button(passwords["name_5"], size=(20, 1))],
            ],
            element_justification="top",
            vertical_alignment="top",
        ),
    ]
]

minor_functions_layout = [[]]

layout = [
    [
        sg.TabGroup(
            [
                [
                    sg.Tab("Tab I", main_functions_layout),
                    sg.Tab("Tab II", minor_functions_layout),
                ],
            ]
        )
    ]
]
window = sg.Window("Assistant", layout, default_element_size=(12, 1))


event_handlers = {
    "Components_data": handle_report_components_data,
    "Products_data": handle_report_products_data,
    "Factory_data": handle_report_factory_data,
    "Groups_statuses": handle_groups_statuses,
    "Groups_balances": handle_components_balances,
    "Create_POs": handle_create_pos,
    "Sort_My_Data": handle_sort_my_data,
    "Supply_info": handle_report_supply_info,
    "New_supply_info": handle_new_supply_info,
    passwords["name_1"]: lambda x: handle_get_password(x, passwords["pass_1"]),
    passwords["name_2"]: lambda x: handle_get_password(x, passwords["pass_2"]),
    passwords["name_3"]: lambda x: handle_get_password(x, passwords["pass_3"]),
    passwords["name_4"]: lambda x: handle_get_password(x, passwords["pass_4"]),
    passwords["name_5"]: lambda x: handle_get_password(x, passwords["pass_5"]),
}

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == "Exit":
        break
    if event in event_handlers:
        # event_handlers[event](values)
        try:
            event_handlers[event](values)
        except Exception as e:
            error_message = (
                f"An error occurred while processing the '{event}' event:\n{str(e)}"
            )
            sg.popup_error(error_message)

window.close()
