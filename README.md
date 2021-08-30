# python-AddrToLatLng [![Python 3.x](https://img.shields.io/badge/python-3.7-green.svg)](https://www.python.org/)[![License](https://img.shields.io/badge/license-Public_domain-green.svg)](https://wiki.creativecommons.org/wiki/Public_domain)
Convert Address to Geometry(Latitude &amp; Longtitude) using Google Geocoding API with python and xlsx file.

## Requirements

This project is developed with python 3.7, and need to install some packages.

```
$ pip install requests
$ pip install openpyxl
```

## Usage

```
$ python python_AddrToLatLng.py -h

usage: python_AddrToLatLng.py [-h] -k API_KEY [-i INPUT_FILE] [-o OUTPUT_FILE]
                              [-wb WORKBOOK] [--addr_col ADDR_COLUMN]
                              [--lat_col LAT_COLUMN] [--lng_col LNG_COLUMN]

Convert your address to Geometry using Google Geocoding API.

optional arguments:
  -h, --help            show this help message and exit
  -k API_KEY, --api-key API_KEY
                        Input your google api key
  -i INPUT_FILE, --input-file INPUT_FILE
                        Input xlsx file(e.g. "./data/input.xlsx")
  -o OUTPUT_FILE, --output-file OUTPUT_FILE
                        output xlsx file(e.g. "./outputs/result.xlsx")
  -wb WORKBOOK, --workbook WORKBOOK
                        Select a workbook in xlsx file. Default='工作表1'
  --addr_col ADDR_COLUMN
                        Address column in xlsx file. Default=4
  --lat_col LAT_COLUMN  Latitude column going to add in xlsx file. Default=11
  --lng_col LNG_COLUMN  Longitude column going to add in xlsx file. Default=12
```

## Sample

```
$python python_AddrToLatLng.py -k {YOUR_API_KEY}
```

Run with default option. 

- Input file path : "./data/input.xlsx"
- Output file path : "./outputs/result.xlsx" 
- Workbook : "工作表1"
- Address column : 4
- Latitude column : 11
- Longitude column : 12

