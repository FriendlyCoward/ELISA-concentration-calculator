# ELISA concentration calculator
A program that calculates unknown plate well concentrations based on well light absorbance and known well concentrations

## Description
ELISA (Enzyme-Linked ImmunoSorbent Assay) is a type of experiment which uses protein-specific antibodies (conjugated with an enzyme) to detect or quantify proteins (antigens or antibodies) via colourimetry.  
To calculate the concentrations of analyzed proteins, plate wells with known concentrations are required (control wells).  
This program fits a graph formula with various parameters to the control well concentrations and their light absorbance and then uses the fitted parameters to calculate unknown well concentrations.  

## Program usage
Because the program was made in part to help me work in my lab it expects some things to work properly:
* Data files are in .xlsx format.
* Data is located in a sheet whose name starts with "Basic Calculation", there can only be one sheet with this kind of name.
* Control wells must be in every plate, in the same locations and with the same concentrations.  

There is an example template and a "data" folder with example data files.  

1. Open "template.ods" and change it for use.
+ In "template": write single characters in cells (C - control wells, E - experimental wells), Groups can be made by also writing a number after the symbol (C1 is control group 1, E1 is experimental group 1). Making experimental groups with one control is fine (C for E1, C for E2, and so on), but if you are using control well groups, then make sure that for each control group there is a matching experimental group (for C1 there is E1, for C2 there is E2, and so on).
+ In "concentrations": write concentrations in the cells which were marked as control wells.

2. Place data files in the "data" folder

### Running with python
1. Navigate to the program folder using the command line.
2. Write in command line:
```
python calculator.py
```
OR
* Instead of going through command line, just double click calculator.py and wait till the terminal closes.

### Running with jupyter lab
1. Navigate to the program folder using the command line.
2.  Write in command line:
```
jupyter lab
```
3. Jupyter lab should launch in your browser (if not, then ctrl+click one of the links provided in the terminal).
4. On the left side double click the file named "calculator.ipynb".
5. Up top click the menu "Run" and then "Run All"

### Results
After running the program the results will be generated in .xlsx and .ods files in the program folder.  
Inside the result file there are sheets for each data file processed.  

Empty cells mean that those wells absorbance values were higher or lower than the maximum or minimum absorbance of control wells. To potentially fix this, start from a higher concentration for control wells (so that the maximum light absorbance would be higher).  

.xlsx file also has coloured cells: green - control wells, red - absorbance over maximum or minimum, yellow - concentration effectively zero.
Graph pictures are also generated in the pictures folder.

## Installation
If you dont have python installed, you can download and install it [here](https://www.python.org/downloads/).  
This program was built using python 3.13.1.  

Download calculator.py, calculator.ipynb (optional), template.xlsx, requirements.txt and put them into a fresh folder.  

Create 2 folders in the program folder and name them "data" and "pictures".
Inside the pictures folder also create 2 folders named "controls" and "experimentals".  

Download required packages. To do this:  
1. Open a terminal (linux, mac) or command prompt (windows).
2. Navigate to the program folder using `cd` commands. If you put the program folder in your desktop the commands could look like this:
```
cd Desktop
cd my_folder
```
3. When in the folder, write:
```
pip install -r requirements.txt
```

(optional) If you want to use the notebook (.ipynb) then download jupyter lab:
```
pip install jupyterlab
```

The program is ready for use!
