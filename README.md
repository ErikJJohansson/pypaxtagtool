# PythonPlantPAX



## Description

This project takes the PlantPAX tag configuration spreadsheet and reads/writes to a PLC using Python instead of VBA.

## Motivation

The original PlantPAX tag configuration tool runs on excel using VBA. The computer the macro is being run on also needs RSLinx classic installed.

Using a Python based version we can forgo the requirement of needing excel on the computer. this script can also be run on computers that don't have RSLinx classic installed or you can't change the driver settings.

Furthermore. The original tool requires you to supply all the tag names for the instances inside the PLC. this will populate the tags upon running the read command. So you won't miss any instances!

Also, this seemed like a fun project with a clearly defined scope.

## Reading tags from PLC

The tool requires a few command line arguments to work. a properly formatted command is shown below

```
pypax.py 10.10.17.10/4 read [ProcessLibraryOnlineConfigTool.xlsm]

```

10.10.17.10/4 is the PLC IP address and slot number of the PLC, without the slot number and just the IP address (like '10.10.17.10' it will default to slot 0)

If no file is specified, the default file in the repo 'ProcessLibraryOnlineConfigTool.xlsm' will be used.

For read mode. A new file will be created using the template file. The file will be named PLCNAME_ConfigTags.xlsx.

Edit the values you want to change in the newly outputted file.

The provided .xlsm repo has all the PlantPAX AOI types and their configuration tag values in it. Please be careful when modifying this file. It's preferable if you don't

The script will loop through each sheet that looks like a PlantPAX AOI and poke the PLC for the number of instances of each type. This then gets written to the "Setup" sheet in the workbook.

For each tag instance, a bulk read will be done to the PLC to get all the tag data. This then gets written to the spreadsheet for the AOI type.

## Writing tags to PLC.

The tool requires a few command line arguments to work. a properly formatted write command is shown below

```
pypax.py 10.10.17.10/4 write PLCNAME_ConfigTags.xlsx 

```

For write mode. Use the desired file with tag information you wish to write to the PLC. If the filename DOES NOT contain the name of the PLC being written to, the script will exit. This is to ensure we are writing to the correct PLC.

Write mode will only write the changes detected between the PLC tag value and the spreadsheet tag value. The number of instances to write is determined by the rows in each AOI spreadsheet. There must be no spaces between any names in the names column otherwise the blank will be treated like the last row to write.

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
ping 10.10.17.10

```

