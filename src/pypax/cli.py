"""Command-line interface for the PlantPAX tag configuration tool.

Thin orchestration over the ``pypax`` core library: parse arguments,
connect to the PLC, and run the read or write flow. All spreadsheet,
PLC, and diff logic lives in :mod:`pypax.spreadsheet`, :mod:`pypax.plc`,
and :mod:`pypax.diff`.

Usage (unchanged from the legacy CLI)::

    pypax.py 10.10.17.10/4 read [ProcessLibraryOnlineConfigTool.xlsm]
    pypax.py 10.10.17.10/4 write PLCNAME_ConfigTags.xlsx
"""

from __future__ import annotations

import argparse

import openpyxl
from pycomm3 import LogixDriver
from tqdm import tqdm

from .diff import row_differences as find_row_differences
from .layout import SETUP_SHEET_NAME, START_ROW, get_aoi_list
from .plc import (
    failed_tag_formatter,
    get_aoi_tag_instances,
    get_failed_tags,
    make_tag_list,
    read_aoi_instances,
    read_plc_row,
    write_plc_row,
)
from .spreadsheet import (
    get_aoi_setup,
    get_subtag_list,
    read_data_sheet_row,
    set_num_instances,
    write_sheet_row,
)


def main():
    # default filename of template file included in the repo
    default_excelfile = "ProcessLibraryOnlineConfigTool.xlsm"

    # Parse arguments
    parser = argparse.ArgumentParser(
        description="Python-based PlantPAX tag configuration tool.",
        epilog="This tool works on both Windows and Mac.",
    )

    # Add command-line arguments
    parser.add_argument("commpath", help="Path to PLC")
    subparsers = parser.add_subparsers(dest="mode", help="Select read/write mode")

    # parsing read commands, filename is optional and will default to default_excelfile value
    read_parser = subparsers.add_parser("read", help="Read tags from PLC into spreadsheet")
    read_parser.add_argument(
        "excelfile", nargs="?", default=default_excelfile, help="Path to excel file"
    )

    # parsing write commands, excelfile is required
    write_parser = subparsers.add_parser("write", help="Write data from spreadsheet into PLC tags")
    write_parser.add_argument("excelfile", help="Path to excel file")

    args = parser.parse_args()

    # Access the parsed arguments
    commpath = args.commpath
    excelfile = args.excelfile
    mode = args.mode

    # open connection to PLC

    plc = LogixDriver(commpath, init_tags=True, init_program_tags=True)

    print("Connecting to PLC.")
    try:
        plc.open()
        plc_name = plc.get_plc_name()

        print("Connected to " + plc_name + " PLC at " + commpath)
    except Exception:
        print("Unable to connect to PLC at " + commpath)
        exit()

    # open excel file

    # filename check
    if mode == "write" and excelfile.find(plc_name) == -1:
        print(
            "Filename mismatch. The file '" + excelfile + "' does not contain '" + plc_name + "'."
        )
        exit()

    print("Opening " + excelfile)
    try:
        if mode == "write":
            book = openpyxl.load_workbook(excelfile, keep_vba=False, data_only=True, read_only=True)
        elif mode == "read":
            book = openpyxl.load_workbook(excelfile, keep_vba=False)
    except Exception:
        print("Unable to open excel file " + excelfile)
        plc.close()
        exit()

    print("Opened file named " + excelfile)

    # get list of AOI sheet names
    aoi_sheet_names = get_aoi_list(book)

    # should be the first sheet in the workbook
    setup_sheet = book[SETUP_SHEET_NAME]

    # read from PLC
    if mode == "read":
        print("Reading tags from " + plc_name + " PLC.")

        for aoi in aoi_sheet_names:
            # get setup info from PLC tags, write to spreadsheet
            base_tags = get_aoi_tag_instances(plc, aoi)
            num_instances = len(base_tags)
            set_num_instances(setup_sheet, aoi, num_instances)

            if num_instances > 0:
                # get subtag list for given AOI
                sub_tags = get_subtag_list(book[aoi])

                failed_read_tags = []

                # Read all instances of this AOI in chunked bulk reads
                # (<=100 tags per request) instead of one request per row.
                rows = read_aoi_instances(plc, base_tags, sub_tags)

                # write rows to spreadsheet
                for i in tqdm(range(num_instances), "Reading instances of " + aoi):
                    tag_data, read_result = rows[i]
                    tag_list = make_tag_list(base_tags[i], sub_tags)

                    # add to failed tags list if we can't find the tag
                    if not all(read_result):
                        failed_read_tags = failed_read_tags + get_failed_tags(tag_list, read_result)

                    write_sheet_row(book[aoi], START_ROW + i, base_tags[i], tag_data)

                # print to command line if we couldn't read any tags
                if failed_read_tags:
                    print(failed_tag_formatter(failed_read_tags, False))
            else:
                print("No instances of " + aoi + " found in " + plc_name + " PLC.")

        # add plc name to file and save to new file
        outfile = plc_name + "_ConfigTags." + "xlsx"
        print("Finished reading from " + plc_name + " PLC.")
        print("Saving to file " + outfile)
        book.save(outfile)
        print("file saved to " + outfile)

    # Write to PLC
    elif mode == "write":
        print("Writing tags to " + plc_name + " PLC.")

        for aoi in aoi_sheet_names:
            # get aoi info from sheet and plc
            num_instances_in_sheet = get_aoi_setup(book[aoi])
            base_tags = get_aoi_tag_instances(plc, aoi)

            # Check to make sure there are instances in sheet
            if num_instances_in_sheet > 0:
                # get subtags of datatype
                sub_tags = get_subtag_list(book[aoi])

                # reset lists
                tag_data_differences = []
                failed_read_tags = []
                failed_write_tags = []

                # read AOI data from each row, collect differences
                for i in tqdm(range(num_instances_in_sheet), "Comparing instances of " + aoi):
                    # data for one tag and all sub tags from sheet
                    base_tag, tag_data_sheet = read_data_sheet_row(
                        book[aoi], START_ROW + i, sub_tags
                    )

                    # data for one tag and all sub tags from plc
                    tag_list = make_tag_list(base_tag, sub_tags)
                    tag_data_plc, read_result = read_plc_row(plc, tag_list)

                    # add to failed tags list if we can't find the tag
                    if not all(read_result):
                        failed_read_tags = failed_read_tags + [base_tag]

                    else:
                        # compare PLC row to spreadsheet row, add differences
                        row_differences = find_row_differences(tag_data_sheet, tag_data_plc)
                        if row_differences:
                            tag_data_differences += row_differences

                # print to command line if we couldn't read any tags for that instance
                if failed_read_tags:
                    print(failed_tag_formatter(failed_read_tags, False))

                # calculate number of changes
                num_tag_changes = len(tag_data_differences)

                # write data to plc
                if num_tag_changes > 0:
                    # format output based on number of changes
                    if num_tag_changes >= 2:
                        print(
                            "Writing "
                            + str(num_tag_changes)
                            + " tag changes to instances of "
                            + aoi
                        )
                    else:
                        print(
                            "Writing " + str(num_tag_changes) + " tag change to instances of " + aoi
                        )

                    # store results of PLC write
                    write_result = write_plc_row(plc, tag_data_differences)

                    # add to failed tags list if we can't find the tag
                    if not all(write_result):
                        # extract tag name from list of tuples
                        tag_difference_list = [t[0] for t in tag_data_differences]

                        failed_write_tags = get_failed_tags(tag_difference_list, write_result)

                        # print to command line if we couldn't write any tags
                        print(failed_tag_formatter(failed_write_tags, True))

                else:
                    print("No differences for instances of " + aoi)
            else:
                print("No instances of " + aoi + " in " + plc_name + " PLC.")

        print("Finished writing to " + plc_name + " PLC.")

    plc.close()
    book.close()


if __name__ == "__main__":
    main()
