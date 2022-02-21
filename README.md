# BlockImportExport
Block Import and Export autolisp for AutoCAD

This is AutoLISP application for AutoCAD. It will create window for easy block import to drawing or export from drawing. 

### This application needs OpenDCL to work. OpenDCL can be downloaded from https://opendcl.com/

Tested on OpenDCL 9.1.4.0 and AutoCAD Civil 3D 2022.

All block data needs to be in this format:
(BlockName; Y coordinate; X coordinate; Z coordinate; Attribut1; Attribut2; ...; AttributN)
Only BlockName and YXZ coordinates are necessary.

Need to call BlockImportExport command to open Block Import and Export window.

## Usage

Main Window
![image](https://user-images.githubusercontent.com/27780893/155004382-b0497989-69d2-4fd1-822d-ca62914d3509.png)

On the left top is shown list of all blocks in current drawing. Double clicking on block name will open new window for block preview. Block preview only shows elements whoes layers are shown. At the bottom is button for importing missing blocks to current drawing from external dwg file.

At left from Blocks field is Delimiter field with all supported delimiters and text box for custom delimiter. Most often Auto delimiter will work well.

Button "Import" will open window for data import from csv files (or other text files). Here you can specify directory of csv files. On left side will be shown files and folders of current directory. Double clicking folder will open this folder. Double clicking ".." will go up by one directory. Double clicking file it will be moved to right side into selected files list. Many files can be selected at once and moved to right with "Add Files" button. Files can be removed from selected files list by double clicking them or selecting them and pressing "Remove Files" button. Files can be filtered by selecting CSV and TXT options. If CSV and TXT options are not selected all files will be shown. If CSV option is selected, only CSV files will be shown.

![image](https://user-images.githubusercontent.com/27780893/155005226-2c9eafee-6f90-4f98-b3fc-723b4691f3a8.png)

Data can also be imported from clipboard. To import data from clipboard press on button "Past" in Main window. And past data to new window.

![image](https://user-images.githubusercontent.com/27780893/155006345-7f4ed354-a089-4002-b525-90a0e061fbe4.png)

After importing block data to Main window it will look like this:

![image](https://user-images.githubusercontent.com/27780893/155007644-b7730701-88ca-4b60-b78c-8be3adac609d.png)

Infromation section shows information about block data. Mainly number of times block will be imported and if block with Name don't exist in current drawing it will be shown as comment. In this case you will need to import block from external drawing file.

If you choose to import block data using Import button or Past button, Export button will be disabled. If you choose to export blocks using Export button, Import and Past buttons will be disabled.

You can export blocks by two ways: all blocks of same name or by selecting blocks manualy. Blocks data will be entered to block data grid. You can copy all block data using Copy button (New window with formated block data will open, just copy it), or by pressing Ok button and saving all block data as CSV file.

