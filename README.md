# FileHarvester
A small DFIR triage tool to collect files of interest.

**FileHarvester** is a small utility used to collect files from a specific target directory and all of its sub-directories. The user points the tool to the target directory and chooses what type of file category he wants to 'harvest' from it (Document Files and Media Files are supported). The tool traverses the directory and all of its sub-directories copying all the files of interest to the output selected folder.

The tool will then create a CSV report file containing the copied file's original path, created and last modified timestamp, the name it has in the output folder (given the fact that files with same names cannot exist in the same folder and thus are being renamed) as well as a flag for providing more info on the copy action. 

The output folder, containing the copied harvested files and the CSV report, is then being zipped and MD5 hashed. The whole procedure is being reported in the verbose analysis pane of the tool and a log file is being created in the output folder after the completion of the analysis. 

Similar tools exist for sure but this one is by far the quickest implementation when you want a specific collection from a targeted folder (e.g. the whole user folder on a Windows machine). Original file's metadata are not being altered and thus the whole harvesting procedure leaves the computer files intact.

## Installation

This is a tool written in Python (version 3.8.5 used). The .exe file (**FileHarvester.exe**) works on Microsoft Windows based machines by just double clicking.

The source code file (**FileHarvester.py**) can be run on a system with python 3 installed (Version 3.6 and above needed). It only needs two additional libraries to run successfully. Use the package manager [pip](https://pip.pypa.io/en/stable/) to install them.

```bash
pip install textract
```
for the [textract](https://pypi.org/project/textract/) module from Dean Malmgren at https://pypi.org/project/textract/

```bash
pip install PySimpleGUI
```
for the [PySimpleGUI](https://pypi.org/project/PySimpleGUI/) module from  MikeTheWatchGuy at https://pypi.org/project/PySimpleGUI/

## Usage

![GitHub Logo](/MainGUI.PNG)

The tool comes with a GUI interface. User has to point the tool to a target folder and choose a file type category from the checkboxes (Documents, Media or both). User then has the following options:

**1.** Run a **Calculate action** to get an estimation of the total size the chosen files are. In this way he can check whether an external media can hold all the files after their collection and thus if the media is suited as an output folder destination.

**2.** Run a **Harvest action** only after providing the tool with an **output** folder destination. Files of the chosen file type are going to be harvested throughout the folder and its sub-folders and copied to the output folder. The CSV report will be created, providing info about the files and the output folder will be zipped and MD5 hashed. A log file will then be created to the output folder.

**3.** Run a **Keyword Harvest action** only after providing the tool with an **output** folder destination and providing keyword/keywords in the specific input widget. The keyword harvest action is only supported for the Document File type and is going to be ignored if a media file type has been selected for collection. Document Files are going to be harvested throughout the folder and its sub-folders and copied to the output folder **only** if they contain one of the keywords in their content (flag: '**KEYWORD FOUND**'). The tool will try to read all of the files' contents but in the case it fails it will copy the unread file anyway and provide a flag that 'KEYWORD NOT SEARCHED'. The CSV report will be created, providing info about the files and the output folder will be zipped and MD5 hashed. A log file will then be created to the output folder.

## IMPORTANT INFO

- ALWAYS **run the tool as an administrator**, otherwise it can crash on folders/ sub-folders where the permissions deny the read action.

- Failure to copy a certain file will be logged both in the csv report and the final log file. Such a failure can be handled gracefully and so the tool will continue working on the rest of the files even if a corrupted/protected file is found.Flag in the report CSV file will be '**NOT COPIED**', '**ERROR**'.

- On a **Keyword Harvest Action**, files that were successfully read but no keyword was found in their content, won't be copied in the output folder but they will be logged in the CSV report with the proper flag, '**NOT COPIED**', '**NO KEYWORD FOUND**'.

- Files with the **same name** are being **renamed** when copied to the output folder with the format: SameFileName_RN.ext and the rename action is being logged in the log file. As many times the pattern (underscore_RN) is found on a copied file's name in the output folder these many times the file has been found with the same name throughout the folder traversal (e.g file_RN_RN_RN_RN.ext has been found 4 times with the same name).

- In the case of multiple keywords to be searched in the files, the **keywords are to be given with a comma between them** and no space (e.g. keyword1,keyword2,keyword3). 

- Document Files collected: doc, docx, xls, xlsx, pdf, pptx, odt, odf.

- Media Files collected: jpg, png, avi, mov, mp4, 3gp. 

## SUPER IMPORTANT INFO

- **DO NOT** choose an output folder inside the target collection folder...unless you liked the movie Inception!

## License
[MIT](https://github.com/D-Kats/FileHarvester/blob/main/LICENSE)

