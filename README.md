


<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/zensii/excel_table_merger/edit/main/README.md">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>



<!-- ABOUT THE PROJECT -->
## About The Project

### A small program that takes two excel files with similar structure and allows for extending of the 'main' file with data from the other.
#
### How it works

The program utilises a fork of PySimpleGui -> FreeSimpleGui for the frontend GUI
The Backend is written in Python and uses Pandas library for some data manipulation.

Upon starting the program, the user is presented with an interface to select the files that need to be update and a save location:

![image](https://github.com/user-attachments/assets/fe6deb7f-9036-444d-94c4-ce9a1a494609)

#

After the needed locations are selected the "Confirm" button must be pressed in order for the backend to execute the nessesary checks.
If all is good a confirmation message is displayed.
#
![image](https://github.com/user-attachments/assets/828d1f39-d0d8-4e47-9f96-4fe4bee38004)

#

Next by pressing "Execute" another window is presented to the user showing column matches between the two files:
#
![image](https://github.com/user-attachments/assets/19dea226-6baf-415c-b7dd-f0a65cb455da)

If multiple sheets are present in the files, multiple tabs are displayed each suggesting the possible merges.
The user then can select which or all columns he wants to update and press "Submit"
#
In the end a new file with the merged columns is saved at the specified location.
#

### Installation

_Just download and run the provided binary executable for a windows OS._

or alternatively:

_download and run the python script directly or compile your own executable._

