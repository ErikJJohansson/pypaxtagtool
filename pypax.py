from pycomm3 import LogixDriver
from sys import argv
import openpyxl

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

def get_dim_list(base_tag, dim_list):
    '''
    function takes a list which has the array size and turns it into a single dimension list with all the indexes
    '''
    # remove 0's
    filtered_list = list(filter(lambda num: num != 0, dim_list))

    temp = []

    # this can totally be better, my brain just started hurting
    # idea is to get a single dimension list fo strings with all the indexes so that can be concatenated with base tag

    if len(filtered_list) == 1: # one dimension
        for i in range(dim_list[0]):
            temp.append(base_tag + '[' + str(i) + ']')
    elif len(filtered_list) == 2: # two dimension
        for i in range(dim_list[0]):
            for j in range(dim_list[1]):
                temp.append(base_tag + '[' + str(i) + '][' + str(j) + ']')
    elif len(filtered_list) == 3: # two dimension
        for i in range(dim_list[0]):
            for j in range(dim_list[1]):
                for k in range(dim_list[2]):
                    temp.append(base_tag + '[' + str(i) + '][' + str(j) + '][' + str(k) + ']')

    return temp

def read_tags_from_plc(plc,workbook,aoi_name):
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
    # for each tag

    for i in range(len(aoi_tag_list)):
        #hardcoded offsets
        # write the tag name in column c
        aoi_sheet.cell(START_ROW+i,NAME_COL).value = aoi_tag_list[i]
        
        # loop through colums to read individual tags, tag name is retrieved from column in spreadsheet
        j = START_COL

        sub_tag =str(aoi_sheet.cell(TOP_TAG_ROW,j).value) + str(aoi_sheet.cell(BOTTOM_TAG_ROW,j).value)
        # this means we have data in the cell
        # cells return None when no value, we are concatenating the value of two cells, not the best but it works
        while sub_tag != 'NoneNone':
            aoi_sheet.cell(START_ROW+i,j).value = get_tag_value(plc,aoi_tag_list[i],sub_tag)
            
            #update iterator
            j+=1
            sub_tag =str(aoi_sheet.cell(TOP_TAG_ROW,j).value) + str(aoi_sheet.cell(BOTTOM_TAG_ROW,j).value)

def write_tags_to_plc(plc,workbook,aoi_name):
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

    i = START_ROW
    base_tag = str(aoi_sheet.cell(i,NAME_COL).value)

    # loop through rows
    while base_tag != 'None':

        # loop through colums to write individual tags, tag name is retrieved from column in spreadsheet
        j = START_COL

        sub_tag = str(aoi_sheet.cell(TOP_TAG_ROW,j).value) + str(aoi_sheet.cell(BOTTOM_TAG_ROW,j).value)

        while sub_tag != "NoneNone":

            tag_value = aoi_sheet.cell(i,j).value

            set_tag_value(plc,base_tag,sub_tag,tag_value)

            #update iterator
            j += 1
            sub_tag = str(aoi_sheet.cell(TOP_TAG_ROW,j).value) + str(aoi_sheet.cell(BOTTOM_TAG_ROW,j).value)


        # update iterator
        i += 1
        base_tag = str(aoi_sheet.cell(i,NAME_COL).value)


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



if __name__ == "__main__":

    # open connection to PLC

    #excelfile = argv[1]
    #commpath = argv[2]
    #mode = argv[3]

    # default to read mode
    #if len(argv[3] == 0):
    #    mode = 0


    commpath = '10.10.16.20/0'
    excelfile = 'ProcessLibraryOnlineConfigTool.xlsm'
    outfile = 'TEST.xlsm'
    mode = 0

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
    
    print('Opened file named ' + excelfile)

    # get list of AOI sheet names
    aoi_sheet_names = get_aoi_list(book)

    # read from PLC
    if mode == 0:

        for aoi in aoi_sheet_names:
            print("Reading instances of " + aoi)
            read_tags_from_plc(plc,book,aoi)

    # write to PLC
    elif mode == 1:

        for aoi in aoi_sheet_names:
            print("Writing instances of " + aoi)
            #write_tags_to_plc(plc,book,aoi)

    print('Saving file')
    book.save(outfile)
    print('file saved to ' + outfile)

    plc.close()
    book.close()