import PySimpleGUI as sg
import pyperclip, os
from dotenv import load_dotenv

from assistant_scripts.read_data.read_components_data import ComponentsDataReader
from assistant_scripts.read_data.read_products_data import ProductsDataReader
from assistant_scripts.read_data.read_factory_data import FactoryDataReader
from assistant_scripts.read_data.read_supply_data import SupplyDataReader

from assistant_scripts.report_components_balances import ComponentsBalances
from assistant_scripts.report_groups_statuses import GroupsStatuses
from assistant_scripts.report_supply_info import CreateSupplyInfo

from assistant_scripts.create_pos import create_csv_pos
from assistant_scripts.minor_functions import sort_my_data

load_dotenv()

mem_comp_path = os.getenv("PATH_MEM_LIST")
pos_template = os.getenv("PATH_PO_TEMPLATE")

po_ccn = os.getenv("PO_CCN")
po_mas_loc = os.getenv("PO_MAS_LOC")
po_request_div = os.getenv("PO_REQUEST_DIV")
po_pur_loc = os.getenv("PO_PUR_LOC")
po_delivery = os.getenv("PO_DELIVERY")
po_inspection = os.getenv("PO_INSPECTION")

path_groups = os.getenv("PATH_GROUPS")
path_db = os.getenv("PATH_DB")
path_my_data = os.getenv("PATH_MY_DATA")
path_products_files = os.getenv("PATH_PRODUCTS")
path_components_info = os.getenv("PATH_COMPONENTS")

supplier1 = os.getenv("SUPPLIER_1")
supplier2 = os.getenv("SUPPLIER_2")
supplier3 = os.getenv("SUPPLIER_3")
supply_incoterms = os.getenv("SUPPLY_INCOTERMS")
supply_t_mode = os.getenv("SUPPLY_T_MODE")


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
        file_path=None, folder_path=path_products_files, to_excel=True
    )
    data_handler()


def handle_report_factory_data(values):
    data_handler = FactoryDataReader(
        factory_data_path=values["factory_path"], db_path=path_db
    )
    data_handler()


def handle_create_pos(values):
    create_csv_pos(
        path_excel_dat=pos_template,
        ccn=po_ccn,
        mas_loc=po_mas_loc,
        request_div=po_request_div,
        pur_loc=po_pur_loc,
        delivery=po_delivery,
        inspection=po_inspection,
    )


def handle_sort_my_data(values):
    sort_my_data(directory=path_my_data)


def handle_give_pass_main(values):
    pyperclip.copy(os.getenv("PASS_MAIN"))


def handle_give_pass_bible(values):
    pyperclip.copy(os.getenv("PASS_BIBLE"))


def handle_give_pass_bright(values):
    pyperclip.copy(os.getenv("PASS_BRIGHT"))


def handle_give_pass_mole(values):
    pyperclip.copy(os.getenv("PASS_MOLE"))


def handle_give_pass_rpg(values):
    pyperclip.copy(os.getenv("PASS_RPG"))


def handle_groups_statuses(values):
    report = GroupsStatuses(
        path_components=values["components_data_path"],
        path_products=path_products_files,
        path_groups=path_groups,
        path_supply=values["supply_path"],
        path_db=path_db,
    )
    report()


def handle_components_balances(values):
    report = ComponentsBalances(
        path_groups=path_groups,
        path_components=values["components_data_path"],
        path_supply=values["supply_path"],
    )
    report()


def handle_report_supply_info(values):
    report = CreateSupplyInfo(
        path_supply=values["supply_path"],
        path_components=path_components_info,
        supplier1=supplier1,
        supplier2=supplier2,
        supplier3=supplier3,
        incoterms=supply_incoterms,
        t_mode=supply_t_mode,
    )
    report()


def handle_new_supply_info(values):
    supply = SupplyDataReader(
        path_supply=values["supply_path"], path_groups=path_groups
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
                [sg.Button("Pass_Main", size=(20, 1))],
                [sg.Button("Pass_Bible", size=(20, 1))],
                [sg.Button("Pass_Bright", size=(20, 1))],
                [sg.Button("Pass_Mole", size=(20, 1))],
                [sg.Button("Pass_RPG", size=(20, 1))],
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
    "Pass_Main": handle_give_pass_main,
    "Pass_Bible": handle_give_pass_bible,
    "Pass_Bright": handle_give_pass_bright,
    "Pass_Mole": handle_give_pass_mole,
    "Pass_RPG": handle_give_pass_rpg,
    "Supply_info": handle_report_supply_info,
    "New_supply_info": handle_new_supply_info,
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
