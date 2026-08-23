# dataSyncer

↘️
A tool can read multi source files and update their values to the target files you select(.csv and xlsx supported)

## How to use?

1.running this code via main.py  
2.select source files  
3.select a target file  
4.preview changes or execute updates  

## ⏰Note

The running time depends on the file type of export you selected, normally the .csv file will run fastly
so I recommend you'd better select the .csv file as the export type unless you want to keep the formulas!

## 🥤bilibili cheers !
Please feel free to reach out if you have any questions!


## Requirements

- PySide6>=6.5.0
- polars>=0.19.0
- fastexcel>=0.10.0
- openpyxl>=3.1.0
- xlsxwriter>=3.0.0
- chardet>=5.0.0

## Installation

```bash
pip install -r requirements.txt
