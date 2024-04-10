import os, shutil


def sort_my_data(directory):
    dir_mapping = {
        "C502_ProductsSupplyPlanDisp": "DOWNLOADS\Products_data",
        "C711_PartsSupplyPlanDisp": "DOWNLOADS\Components_data",
        "zpp066_": "DOWNLOADS\Factory_data",
        "AttachRatePlanning-": "DOWNLOADS\AR_downloads",
        "EMS_Forecast_": "RESULTS\EMS_Forecast",
        "Report_components_data_": "RESULTS\Component_report",
        "Report_factory_inventory_": "RESULTS\Factory_inventory",
        "Results_AR_check_": "RESULTS\AR_check",
        "Report_products_data_": "RESULTS\Product_report",
        "Report_groups_statuses_": "RESULTS\Groups_statuses",
        "Report_groups_balances_": "RESULTS\Groups_balances",
    }
    files = os.listdir(directory)
    for file in files:
        for file_type in dir_mapping:
            if file_type in file:
                destination_directory = os.path.join(directory, dir_mapping[file_type])
                source_file_path = os.path.join(directory, file)
                shutil.move(source_file_path, destination_directory)
    directory2 = os.path.join(directory, "products_files_to_report")
    files2 = os.listdir(directory2)
    new_directory = os.path.join(directory, "DOWNLOADS\Products_data")
    for file in files2:
        destination_directory = os.path.join(new_directory, file)
        source_file_path = os.path.join(directory2, file)
        shutil.move(source_file_path, destination_directory)
