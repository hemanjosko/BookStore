import xlrd, xlwt

def read_data():
    '''Read data from excel sheet and make dictionary'''

    #location of of excel sheet data
    loc = ("resource/BookList.xlsx")

    #open excel sheet
    wb = xlrd.open_workbook(loc)

    #all first sheet data
    sheet = wb.sheet_by_index(0)

    headers = []
    whole_data = []
    dict_whole_data = []

    for c_rows in range(0,sheet.nrows):
        m_data = []
        for c_cols in range(0,sheet.ncols):
            if c_rows == 0:
                headers.append(sheet.cell_value(c_rows,c_cols))
            else:
                m_data.append(sheet.cell_value(c_rows,c_cols))
        if c_rows > 0:
            whole_data.append(m_data)

    for x in whole_data:
        a = dict(zip(headers,x))
        dict_whole_data.append(a)

    return dict_whole_data

if __name__== "__main__":
    print read_data.__doc__
    #reading data from excel
    whole_data_from_excel = read_data()




