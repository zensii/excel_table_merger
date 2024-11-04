import os
import pandas as pd


def get_file_structure(file_path: str) -> dict:
    """extract file structure and return a dictionary with sheet names as key and sheet content as df"""

    file_sheets_dict = {}
    xl_df = pd.ExcelFile(file_path)
    for sheet_name in xl_df.sheet_names:
        file_sheets_dict[sheet_name] = xl_df.parse(sheet_name)
    return  file_sheets_dict


def find_matching_columns(df1, df2) -> list:
    """
    Finds all columns with matching names between two df
    Returns a list of matching column names
    """
    matches = []
    for col1 in df1.columns:
        for col2 in df2.columns:
            if col1.lower().strip() == col2.lower().strip():
                matches.append(col1)
    return matches


def extend_selected_columns(df1, df2, selected_columns_dict):
    """
    Extend specific columns in df1 with data from df2 based on user selection.
    unselected columns and filled with NaN.
    """
    # Create a copy of df1 to avoid modifying the original
    result_df = df1.copy()

    # Initialize an empty DataFrame to hold the rows to append
    rows_to_add = pd.DataFrame(index=range(len(df2)), columns=result_df.columns)

    # Process each sheet's selected columns
    for sheet_name, columns in selected_columns_dict.items():
        for col in columns:
            if col in df1.columns and col in df2.columns:
                # Copy data from df2's selected columns into rows_to_add
                rows_to_add[col] = df2[col].values

    # Concatenate the new rows with the original dataframe
    result_df = pd.concat([result_df, rows_to_add], ignore_index=True)

    return result_df


def write_file(main_df, extension_file_df, cols_to_update, output_file_path):

    output_file = os.path.join(output_file_path, "extended_output.xlsx")

    with pd.ExcelWriter(output_file) as writer:
        for sheet_name_main_file, sheet_df in main_df.items():
            if sheet_name_main_file in cols_to_update:
                ext_sheet_df = extension_file_df[sheet_name_main_file]
                # Extend the selected columns based on user choices
                updated_main_df = extend_selected_columns(sheet_df, ext_sheet_df,
                                                          {sheet_name_main_file: cols_to_update[sheet_name_main_file]})
                updated_main_df.to_excel(writer, sheet_name=sheet_name_main_file, index=False)
                print(f"Data for sheet '{sheet_name_main_file}' has been written to {output_file}")
            else:
                sheet_df.to_excel(writer, sheet_name=sheet_name_main_file, index=False)
                print(f"No extension needed for sheet '{sheet_name_main_file}'. Original data written.")

    # Close the ExcelWriter to save the file
    writer.close()
    print(f"All sheets have been successfully written to {output_file}")

