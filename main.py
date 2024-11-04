from funcs import *
from guis import create_multiple_selection_window, main_gui_window

# user file selection
file_path_main, ext_file_path, save_location_path = main_gui_window()

# getting dataframes data
main_df = get_file_structure(file_path_main)

# getting df for extension file
extension_file_df = get_file_structure(ext_file_path)

# checking for matching sheet and column names that can be extended
matching_sheets_cols = {}
for sheet_name_main_file, sheet_df in main_df.items():
    for ext_sheet_name, ext_sheet_df in extension_file_df.items():
        if sheet_name_main_file == ext_sheet_name:
            matching_sheets_cols[sheet_name_main_file] = find_matching_columns(sheet_df, ext_sheet_df)

# present a user choice which columns to update
cols_to_update: dict = create_multiple_selection_window(matching_sheets_cols)

# update the selection
write_file(main_df, extension_file_df, cols_to_update, save_location_path)


