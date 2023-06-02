# PythonPlantPAX



## Description

This project takes the PlantPAX tag configuration spreadsheet and reads/writes to a PLC using Python instead of VBA.

When using read mode, the script will find all the instances of the AOI types in the PLC.

## Usage

The tool requires a few command line arguments to work. a properly formatted command is shown below

```
pypax.py -R ProcessLibraryOnlineConfigTool.xlsm 10.10.16.20/5

pypax.py -W ProcessLibraryOnlineConfigTool_CVM.xlsm 10.10.16.20/5

```

-R/-W - switch between read and write mode

For read mode. A new file will be created with the input file with the PLC name on the end of it.
For write mode. Use the desired file with tag information you wish to write to the PLC. Maybe i'll add that the file name matches the PLC name before writing!

10.10.16.20/5 is the PLC IP address and slot number of the PLC, without the slot number and just the IP address (like '10.10.16.20' it will default to slot 0)

## Installation

Please ensure you have the python packages installed as specified in the requirements.txt file.

Navigate to the directory where you cloned the repo and run the command below

```
pip3 install requirements.txt

```

## Troubleshooting

Can you ping the PLC you are trying to read from? Ensure you have network connectivity to the PLC before running this script.

If you do not know how to ping, run the command below, it should be the same for Mac/Unix and Windows. Replace the IP address with the PLC you wish to ping

```
ping 10.10.16.20

```

