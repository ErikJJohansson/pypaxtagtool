# PythonPlantPAX



## Description

This project takes the PlantPAX tag configuration spreadsheet and reads/writes to a PLC using Python instead of VBA.

This version can poke/find all the AOI instances of a certain type and populate the spreadsheet so you don't have to make it yourself.

## Usage

The tool requires a few command line arguments to work. a properly formatted command is shown below

```
pypax.py -r ProcessLibraryOnlineConfigTool.xlsm 10.10.16.20/5

pypax.py -w ProcessLibraryOnlineConfigTool_CVM.xlsm 10.10.16.20/5

```

-r/-w - switch between read and write mode

with read mode. A new file will be created with the input file with the PLC name on the end of it.
with write mode. Use the desired file with tag information

10.10.16.20/5 is the PLC IP address and slot number of the PLC, without the slot number it will default to slot 0

## Installation

Please ensure you have the python packages installed as specified in the requirements.txt file.