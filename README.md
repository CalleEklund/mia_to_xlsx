# mia_to_xlsx
Created by Carl Eklund, Linköping, 2021

### General Info
Used to convert a .mia file in the course TSEA28 at Linköpings University to specific formated .xlsx file. Built using Python 3.9 and the [openpyxl](https://pypi.org/project/openpyxl/) module.

### Setup
Clone the git repo
Change the fileencoding from .mia .txt or the file will not be found.
```
run by executing ./miatoxlsx.py 'name1' 'name2' 'nameofthemiafile'
```
If no arguments are given the default values are:  
name1 & name2 = 'no name given'  
nameofthemiafile = 'mia_demo.txt'
