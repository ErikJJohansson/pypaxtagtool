from pycomm3 import LogixDriver
from sys import argv
import openpyxl
from tqdm import trange, tqdm

'''

    argv 1 path to excel file (REQUIRED to be backwards compatible
    argv 2 PLC path if we want to be backwards compatible without changing spreadsheet
    argv 3 Check for -r or -w for read/write

'''
# These are constants based on the PlantPAX spreadsheets containing AOI data
START_ROW = 10
START_COL = 5
NAME_COL = 3
TOP_TAG_ROW = 7
BOTTOM_TAG_ROW = 8

# These are constants based on the PlantPAX setup sheet at the beginning of the workbook

NUM_INSTANCES_COL = 4 # in setup sheet
NUM_SUBTAGS_COL = 5  # in setup sheet

def get_aoi_tag_instances(plc, tag_type):
    """
    function returns list of tag names matching struct type
    """
    #return tag_list

    tag_list = []

    for tag, _def in plc.tags.items():
        if _def['data_type_name'] == tag_type:
            if _def['dim'] > 0:
                tag_list = tag_list + get_dim_list(tag,_def['dimensions'])
            else:
                tag_list.append(tag)

    return tag_list

def get_aoi_list(excel_book):
    aoi_list = []

    # PlantPAX AOI's have an _ for second character
    for sheet in excel_book.sheetnames:
        if sheet[1] == '_':
            aoi_list.append(sheet)

    return aoi_list

def get_subtag_list(sheet):
    '''
    function gets all subtags in a given sheet, returns a list of subtags
    '''
       
    sub_tag_list = [] 
    i = START_COL
    sub_tag = get_subtag(sheet,i)

    while sub_tag != 'NoneNone':            
        
        sub_tag_list.append(sub_tag)
        
        #update iterator
        i+=1
        sub_tag = get_subtag(sheet,i)

    return sub_tag_list

def get_subtag(sheet, column):
    '''
    function gets subtag based on column
    '''

    sub_tag = str(sheet.cell(TOP_TAG_ROW,column).value) + str(sheet.cell(BOTTOM_TAG_ROW,column).value)

    return sub_tag

def search_value_in_col(sheet, search_string, col_idx=1):
    '''
    search a column for the specific string, return row on match
    '''
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row,col_idx).value == search_string:
            return row
    
    return None

def get_aoi_setup(sheet, aoi_name):

    # 8 is col H "Worksheet tab name"
    aoi_row = search_value_in_col(sheet, aoi_name, 8)

    num_aoi_tags = sheet.cell(aoi_row,NUM_INSTANCES_COL).value
    num_sub_tags = sheet.cell(aoi_row,NUM_SUBTAGS_COL).value

    if aoi_row != None:
        return num_aoi_tags, num_sub_tags
    else:
        return 0,0

def set_num_instances(sheet, aoi_name, num):
    '''
    Function updates the num instances in the Setup page of spreadsheet
    '''

    # 7 is col H "Worksheet tab name"
    aoi_row = search_value_in_col(sheet, aoi_name, 8)
    
    if aoi_row != None:
        sheet.cell(aoi_row,4).value = num



def get_dim_list(base_tag, dim_list):
    '''
    function takes a list which has the array size and turns it into a list with all iterations
    '''
    # remove 0's
    filtered_list = list(filter(lambda num: num != 0, dim_list))

    temp = []

    # this can totally be better, my brain just started hurting
    # idea is to get a single dimension list of strings with all the indexes so that can be concatenated with base tag

    if len(filtered_list) == 1: # one dimension
        for i in range(dim_list[0]):
            temp.append(base_tag + '[' + str(i) + ']')
    elif len(filtered_list) == 2: # two dimension
        for i in range(dim_list[0]):
            for j in range(dim_list[1]):
                temp.append(base_tag + '[' + str(i) + '][' + str(j) + ']')
    elif len(filtered_list) == 3: # three dimension
        for i in range(dim_list[0]):
            for j in range(dim_list[1]):
                for k in range(dim_list[2]):
                    temp.append(base_tag + '[' + str(i) + '][' + str(j) + '][' + str(k) + ']')

    return temp

def make_tag_list(base_tag,sub_tags):
    '''
    returns the full tag path of a given base tag and sub tags
    '''
    # concatenate base tag
    read_list = [base_tag + s for s in sub_tags]

    return read_list

def read_plc_row(plc, tag_list):
    '''
    reads data from plc, returns list of tuples (tag_name, tag_value)
    '''
    
    if plc.connected:
        tag_data = plc.read(*tag_list)

    # tuple of tag name, data
    return [(s[0],s[1]) for s in tag_data]

def write_plc_row(plc, tag_data):
    if plc.connected:
        plc.write(*tag_data)

def write_sheet_row(sheet,row,base_tag,tag_data):
    '''
    writes tag data to a row in spreadsheet
    '''
    # write name
    sheet.cell(row,NAME_COL).value = base_tag

    # write data    
    for i in range(len(tag_data)):
        # change booleans to 0 and 1 in spreadsheet
        # this could probably be cleaner
        # index 1 is data in Tuple
        if tag_data[i][1] == True:
            sheet.cell(row,START_COL+i).value = 1
        elif tag_data[i][1] == False:
            sheet.cell(row,START_COL+i).value = 0
        else:
            sheet.cell(row,START_COL+i).value = tag_data[i][1]

def read_sheet_row(sheet,row,sub_tags):
    '''
    reads tag name and data from list
    '''
    base_tag = sheet.cell(row,NAME_COL).value

    tag_data = []

    # loop through subtags, get data
    for i in range(len(sub_tags)):
        tag_data.append((base_tag + sub_tags[i],sheet.cell(row,START_COL+i).value))

    return tag_data

if __name__ == "__main__":

    # Arguments checking
   
    if len(argv) == 4:
        mode = argv[1]
        excelfile = argv[2]
        commpath = str(argv[3])
    else:
        print('Cannot run script. Invalid number of arguments.')
        exit()

    # open connection to PLC

    plc = LogixDriver(commpath, init_tags=True,init_program_tags=True)

    print('Connecting to PLC.')
    try:
        plc.open()
        plc_name = plc.get_plc_name()

        print('Connected to ' + plc_name + ' PLC at ' + commpath)
    except:
        print('Unable to connect to PLC at ' + commpath)
        exit()

    # open excel file

    # filename check
    if mode == '-W' and excelfile.find(plc_name) == -1:
        print("Filename mismatch. The file '" + excelfile + "' does not contain '" + plc_name + "'.")
        exit()

    print('Opening ' + excelfile)
    try:
        book = openpyxl.load_workbook(excelfile,keep_vba=False,keep_links=True)

    except:
        print('Unable to open excel file ' + excelfile)
        plc.close()
        exit()
    
    print('Opened file named ' + excelfile)

    # get list of AOI sheet names
    aoi_sheet_names = get_aoi_list(book)

    setup_sheet = book["Setup"]

    # read from PLC
    if mode == '-R':
        print('Reading tags from ' + plc_name + ' PLC.')
        
        for aoi in aoi_sheet_names:
            # get setup info from PLC tags, write to spreadsheet
            base_tags = get_aoi_tag_instances(plc,aoi)
            num_instances = len(base_tags)
            set_num_instances(setup_sheet,aoi,num_instances)

            if num_instances > 0:

                # get subtag list for given AOI
                sub_tags = get_subtag_list(book[aoi])

                # read rows, write to spreadsheet
                for i in tqdm(range(num_instances),"Reading instances of " + aoi):
                    tag_list = make_tag_list(base_tags[i],sub_tags)

                    # data for one tag and all sub tags
                    tag_data = read_plc_row(plc,tag_list)

                    write_sheet_row(book[aoi],START_ROW+i,base_tags[i],tag_data)

            else:
                print("No instances of " + aoi + " found in " + plc_name + " PLC.")


        parsed_filename = excelfile.split('.')

        # add plc name to file and save
        outfile = parsed_filename[len(parsed_filename)-2] + "_" + plc_name + '.' + 'xlsx' #parsed_filename[len(parsed_filename)-1]

        print('Finished reading from ' + plc_name + ' PLC.')
        print('Saving to file ' + outfile)
        book.save(outfile)
        print('file saved to ' + outfile)

    # Write to PLC
    elif mode == '-W':
        print('Writing tags to ' + plc_name + ' PLC.')
        
        for aoi in aoi_sheet_names:

            # get aoi info
            num_instances, num_sub_tags = get_aoi_setup(setup_sheet,aoi)

            if num_instances > 0:

                # get subtags
                sub_tags = get_subtag_list(book[aoi])

                # read spreadsheet rows, write to plc
                for i in tqdm(range(num_instances),"Writing instances of " + aoi):
                    
                    # data for one tag and all sub tags
                    tag_data = read_sheet_row(book[aoi],START_ROW+i,sub_tags)

                    #write data to plc
                    write_plc_row(plc,tag_data)

            else:
                print("Skipping instances of " + aoi)

        print("Finished writing to " + plc_name + " PLC.")

    plc.close()
    book.close()