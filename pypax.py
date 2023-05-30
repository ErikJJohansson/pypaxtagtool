from pycomm3 import LogixDriver
from sys import argv
import openpyxl
from tqdm import trange, tqdm

'''

    argv 1 path to excel file (REQUIRED to be backwards compatible
    argv 2 PLC path if we want to be backwards compatible without changing spreadsheet
    argv 3 Check for -r or -w for read/write

'''

# check for -r for read or -w for write

def get_aoi_tags(plc, tag_type):
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
    # These are constants based on the PlantPAX spreadsheets
    START_ROW = 10
    START_COL = 5
    NAME_COL = 3
    TOP_TAG_ROW = 7
    BOTTOM_TAG_ROW = 8     
    
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
    # These are constants based on the PlantPAX spreadsheets
    START_ROW = 10
    START_COL = 5
    NAME_COL = 3
    TOP_TAG_ROW = 7
    BOTTOM_TAG_ROW = 8 

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

    num_aoi_tags = sheet.cell(aoi_row,4).value
    num_sub_tags = sheet.cell(aoi_row,5).value

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

def get_tag_value(plc,base_tag,sub_tag):

    # index 1 is tag value in list
    try:
        tag_value = plc.read(base_tag + sub_tag)[1]
    except Exception:
        tag_value = ''
        pass

    # boolean catch
    if tag_value == False:
        tag_value = 0
    elif tag_value == True:
        tag_value = 1
    
    return tag_value

def set_tag_value(plc,base_tag,sub_tag,tag_value):

    # index 1 is tag value in list
    try:
        plc.write(base_tag + sub_tag,tag_value)
    except Exception:
        pass

def read_aoi_tags_from_plc(plc,workbook,aoi_name):
    '''
    function will read the tag values for a given AOI
    '''
    
    # These are constants based on the PlantPAX spreadsheets
    START_ROW = 10
    START_COL = 5
    NAME_COL = 3
    TOP_TAG_ROW = 7
    BOTTOM_TAG_ROW = 8   
    
    aoi_tag_list = get_aoi_tags(plc,aoi_name)       
    aoi_sheet = workbook[aoi_name]
    setup_sheet = workbook['Setup']

    # for each tag

    num_aoi_tags = len(aoi_tag_list)

    # update spreadsheet with AOI count
    set_num_instances(setup_sheet,aoi_name,num_aoi_tags)

    if num_aoi_tags > 0:

        for i in tqdm(range(num_aoi_tags),"Reading instances of " + aoi):
            #hardcoded offsets
            # write the tag name in column c
            aoi_sheet.cell(START_ROW+i,NAME_COL).value = aoi_tag_list[i]
            
            # loop through colums to read individual tags, tag name is retrieved from column in spreadsheet
            j = START_COL

            sub_tag = get_subtag(aoi_sheet,j)

            # this means we have data in the cell
            # cells return None when no value, we are concatenating the value of two cells, not the best but it works
            while sub_tag != 'NoneNone':
                aoi_sheet.cell(START_ROW+i,j).value = get_tag_value(plc,aoi_tag_list[i],sub_tag)
                
                #update iterator
                j+=1
                sub_tag = get_subtag(aoi_sheet,j)
    else:
        print("No instances of " + aoi_name)

def write_aoi_tags_to_plc(plc,workbook,aoi_name):
    '''
    write to PLC!
    '''
    # These are constants based on the PlantPAX spreadsheets
    START_ROW = 10
    START_COL = 5
    NAME_COL = 3
    TOP_TAG_ROW = 7
    BOTTOM_TAG_ROW = 8   

    aoi_sheet = workbook[aoi_name]
    setup_sheet = workbook['Setup']

    num_aoi_tags, num_sub_tags = get_aoi_setup(setup_sheet,aoi_name)

    if num_aoi_tags > 0:

        # loop through rows
        for i in tqdm(range(num_aoi_tags),"Writing instances of " + aoi):

            # loop through colums to write individual tags, tag name is retrieved from column in spreadsheet
            j = START_COL

            base_tag = str(aoi_sheet.cell(START_ROW+i,NAME_COL).value)
            sub_tag = get_subtag(aoi_sheet,j)

            while sub_tag != "NoneNone":

                tag_value = aoi_sheet.cell(START_ROW+i,j).value

                set_tag_value(plc,base_tag,sub_tag,tag_value)

                #update iterator
                j += 1
                sub_tag = get_subtag(aoi_sheet,j)

    else:
        print("Skipping instances of " + aoi_name)


if __name__ == "__main__":

    # open connection to PLC

    #excelfile = argv[1]
    #commpath = argv[2]
    #mode = argv[3]

    # default to read mode
    #if len(argv[3] == 0):
    #    mode = 0


    commpath = '10.10.16.20/5'
    excelfile = 'TEST.xlsm' #'ProcessLibraryOnlineConfigTool.xlsm'
    outfile = 'TEST.xlsm'
    mode = 1

    # open connection to PLC

    plc = LogixDriver(commpath, init_tags=True,init_program_tags=True)

    print('Connecting to PLC.')
    try:
        plc.open()
        print('Connected to ' + plc.get_plc_name() + ' PLC.')
    except:
        print('Unable to connect to PLC at ' + commpath)

    # open excel file

    print('Opening ' + excelfile)
    try:
        book = openpyxl.load_workbook(excelfile,keep_vba=True)

    except:
        print('Unable to open excel file ' + excelfile)
        plc.close()
    
    print('Opened file named ' + excelfile)

    # get list of AOI sheet names
    aoi_sheet_names = get_aoi_list(book)

    setup_sheet = book["Setup"]

    # read from PLC
    if mode == 0:

        for aoi in aoi_sheet_names:
            read_aoi_tags_from_plc(plc,book,aoi)
            #print(aoi + ': ' + str(get_aoi_setup(setup_sheet,aoi)))
            #print(aoi + ': ' + str(search_value_in_col(setup_sheet,aoi,8)))

    # write to PLC
    elif mode == 1:

        for aoi in aoi_sheet_names:
            write_aoi_tags_to_plc(plc,book,aoi)
            #print(aoi + str(get_aoi_setup(setup_sheet,aoi)))

    print('Saving file')
    book.save(outfile)
    print('file saved to ' + outfile)

    plc.close()
    book.close()