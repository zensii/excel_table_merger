from funcs import *
import FreeSimpleGUI as Sg


label_save = Sg.Text("Select save location", expand_x=True)
input_box_save = Sg.InputText(tooltip='Select the output file save location.')
add_button_save = Sg.Button('Confirm')

label_new = Sg.Text("Select the files", expand_x=True)
input_box_new = Sg.InputText(tooltip='Select all files for the update')
add_button_new = Sg.Button('Confirm', tooltip="Press to confirm selection")

save_location = Sg.FolderBrowse('Open', key='_save_location_', tooltip='Press to open file browser')
open_explorer_new = Sg.FilesBrowse('Open', key='_open_files_', tooltip='Press to open file browser')
button_go = Sg.Button('Execute', size=(10, 2), tooltip='Press to start the script')
button_exit = Sg.Button('Exit', size=(10, 2), tooltip='Press to exit the program')

window = Sg.Window('Excel Updater', layout=[[label_new, input_box_new, open_explorer_new, ],
                                            [label_save, input_box_save, save_location],
                                            [Sg.Push(), add_button_new],
                                            [Sg.Output(size=(50, 3), key='_output_', expand_x=True)],
                                            [button_go, Sg.Push(), button_exit],
                                            ])


file_paths = None
save_location = None

while True:

    event, file_selection = window.read()
    window['_output_'].update(value='')
    match event:

        case 'Confirm':
            if file_selection['_open_files_'] != '':
                file_paths = file_selection['_open_files_'].split(';')
                print('Files loaded!')
            else:
                print('Please select the files to add!')

            if file_selection['_save_location_'] != '':
                save_location = file_selection['_save_location_']
                print('Output destination confirmed!')
            else:
                print('Please select the output file save destination.')

        case 'Exit':
            window.close()
            sys.exit()

        case Sg.WIN_CLOSED:
            window.close()
            sys.exit()

        case 'Execute':

            if file_paths and save_location:
                print('Script starting...')
                sleep(1)
                cols = ['Company Code', 'G/L Account', 'Vendor', 'Name 1']

                dataframes, df_main, main_file = get_df(file_paths)
                file_year, file_period = get_period(dataframes)
                clean_data(dataframes, cols)
                check_for_df(dataframes)
                pv_tables = create_sub_pivots(dataframes)
                pivot = create_main_pivot(pv_tables, file_year, file_period)
                data = prep_update_package(pivot)
                sheet = get_main_sheet(main_file)
                start_row = get_start_row(df_main)
                update(data, sheet, start_row)

                print('Saving...')
                main_file.save(f'{save_location}/(UPDATED){main_file.name}')  # save the Excel file
                print('All DONE!')
                sleep(0.5)
                print('Quitting...')
                sleep(1)
                app = xw.apps.active  # To quit Excel if needed
                app.quit()
                window.close()
                sys.exit()
            else:
                print("First select the files to work with and press 'Confirm'!")
