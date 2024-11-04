import sys
from pathlib import Path
import FreeSimpleGUI as Sg

def main_gui_window():
    # Set theme for better appearance
    Sg.theme('LightGrey1')

    # Common parameters for consistent styling
    INPUT_SIZE = (40, 1)
    BUTTON_SIZE = (10, 1)

    # Define layouts with better organization and consistent spacing
    file_selection_layout = [
        # Main file selection
        [
            Sg.Text("Select the main file", size=(15, 1), expand_x=True),
            Sg.InputText(tooltip='Select the file to be updated', size=INPUT_SIZE, key='_open_main_'),
            Sg.FileBrowse('Browse', file_types=(("Excel Files", "*.xlsx *.xls"),))
        ],
        # Files to add selection
        [
            Sg.Text("Select the files to add from", size=(15, 1), expand_x=True),
            Sg.InputText(tooltip='Select all files for the update', size=INPUT_SIZE, key='_open_files_'),
            Sg.FilesBrowse('Browse', file_types=(("Excel Files", "*.xlsx *.xls"),))
        ],
        # Save location selection
        [
            Sg.Text("Select save location", size=(15, 1), expand_x=True),
            Sg.InputText(tooltip='Select the output file save location.', size=INPUT_SIZE, key='_save_location_'),
            Sg.FolderBrowse('Browse')
        ]
    ]

    # Bottom section layout
    bottom_layout = [
        [Sg.Push(), Sg.Button('Confirm', size=BUTTON_SIZE, tooltip="Press to confirm selection")],
        [Sg.Output(size=(65, 3), key='_output_', expand_x=True)],
        [
            Sg.Button('Execute', size=BUTTON_SIZE, tooltip='Press to start the script'),
            Sg.Push(),
            Sg.Button('Exit', size=BUTTON_SIZE, tooltip='Press to exit the program', button_color=('white', '#FF0000'))
        ]
    ]

    # Combine layouts
    layout = [
        *file_selection_layout,
        *bottom_layout
    ]

    # Create window with padding and proper sizing
    window = Sg.Window(
        'Excel Updater',
        layout,
        finalize=True,
        resizable=True,
        return_keyboard_events=True
    )

    file_path_main = None
    new_file_paths = None
    save_location_path = None

    while True:
        event, values = window.read()
        window['_output_'].update(value='')  # Clear output

        if event in (Sg.WIN_CLOSED, 'Exit'):
            window.close()
            sys.exit()

        elif event == 'Confirm':
            # Validate main file
            if values['_open_main_']:
                file_path_main = Path(values['_open_main_'])
                if not file_path_main.exists():
                    print('Error: Main file does not exist!')
                    continue
                print('Main File loaded!')
            else:
                print('Please select the main Excel file to update!')
                continue

            # Validate input files
            if values['_open_files_']:
                new_file_paths = Path(values['_open_files_'])
                # new_file_paths = [Path(p) for p in values['_open_files_'].split(';')]
                # if not all(p.exists() for p in new_file_paths):
                #     print('Error: One or more input files do not exist!')
                #     continue
                print('Files loaded!')
            else:
                print('Please select the files to add!')
                continue

            # Validate save location
            if values['_save_location_']:
                save_location_path = Path(values['_save_location_'])
                if not save_location_path.exists():
                    print('Error: Save location does not exist!')
                    continue
                print('Output destination confirmed!')
            else:
                print('Please select the output file save destination.')
                continue

        elif event == 'Execute':
            if all([file_path_main, new_file_paths, save_location_path]):
                window.minimize()
                return str(file_path_main), str(new_file_paths), str(save_location_path)
            else:
                print('Please confirm your selections first!')


def create_multiple_selection_window(matching_cols_dict: dict) -> dict:
    """
    Presents the user with a GUI with choice which columns from which sheets to extend.
    """
    Sg.theme('LightGrey1')

    # Create a tab group layout where each tab represents a sheet
    tab_group_layout = []

    for sheet_name, columns in matching_cols_dict.items():
        # Create checkbox layout for current sheet
        checkbox_layout = [
            [Sg.Checkbox(col, key=f"{sheet_name} --> {col}")] for col in columns
        ]

        # Create a scrollable column for the checkboxes
        scrollable_column = Sg.Column(
            checkbox_layout,
            scrollable=True,
            vertical_scroll_only=True,
            size=(350, 200),
            expand_x=True
        )

        # Create a tab for the current sheet
        sheet_tab = [
            [Sg.Text(f'"Select columns from {sheet_name}:', font=('Helvetica', 12))],
            [scrollable_column]
        ]

        # Add the tab to the tab group
        tab_group_layout.append(Sg.Tab(sheet_name, sheet_tab))

    # Create the main layout with the tab group
    layout = [
        [Sg.TabGroup([tab_group_layout], enable_events=True, key='-TABGROUP-', expand_x=True)],
        [Sg.Button('Select All', size=(10, 1)),
         Sg.Button('Deselect All', size=(10, 1)),
         Sg.Push(),
         Sg.Button('Submit', size=(10, 1)),
         Sg.Button('Cancel', size=(10, 1))]
    ]

    window = Sg.Window('Sheet and Column Selection', layout, size=(420, 320), resizable=True)

    selected_items = {}
    current_tab = next(iter(matching_cols_dict.keys()))  # Start with first sheet

    while True:
        event, values = window.read()

        if event == Sg.WIN_CLOSED or event == 'Cancel':
            # Return empty dict if cancelled
            break

        if event == '-TABGROUP-':
            current_tab = values['-TABGROUP-']

        if event == 'Select All':
            # Select all checkboxes in current tab
            for key in values:
                if isinstance(key, str) and key.startswith(f"{current_tab} --> "):
                    window[key].update(True)

        if event == 'Deselect All':
            # Deselect all checkboxes in current tab
            for key in values:
                if isinstance(key, str) and key.startswith(f"{current_tab} --> "):
                    window[key].update(False)

        if event == 'Submit':
            # Create dictionary of selected items for each sheet
            for key, value in values.items():
                if isinstance(key, str) and ' --> ' in key and value:
                    sheet_name, col_name = key.split(' --> ')
                    if sheet_name not in selected_items:
                        selected_items[sheet_name] = []
                    selected_items[sheet_name].append(col_name)
            break

    window.close()
    return selected_items
