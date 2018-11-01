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
                headers.append((sheet.cell_value(c_rows,c_cols))
                               .replace(" ","_").replace("-","_").lower())
            else:
                m_data.append(sheet.cell_value(c_rows,c_cols))
        if c_rows > 0:
            whole_data.append(m_data)

    for x in whole_data:
        a = dict(zip(headers,x))
        dict_whole_data.append(a)

    print "Headers"
    print headers

    return dict_whole_data

def make_three_main_category_boxes(data_set):
    '''Three main category of boxes
    1. Adult + Kids Bundles - 5 books each
        At least 1 fiction Adult
        At least 1 Non-fiction Adult
        At least 1 Kid coloring
        At least 1 Kid fiction

    2. Adult Only Bundles - 5 books each
        At least 1 fiction
        At least 1 non fiction
        At least and only 1 colouring Adult

    3. Kids Only Bundles - 5 books each
        At least 2 Colouring books

    The rules are
    Every bundle should have 5 books
    No bundle should have the same 2 titles
    No bundle can exceed 1800 gms
    Kids bundle cannot have any adult books '''

    # Three main box categories
    adult_kids_category_boxes = []
    adult_only_category_boxes = []
    kids_only_category_boxes = []

    #only main two category data adult and kids
    clean_up_data_set = []
    for data in data_set:
        if data['bos_category'] in ('ADULT', 'KIDS'):
            clean_up_data_set.append(data)
    #print len(clean_up_data_set)

    # create bundles for adults and kids category
    bundle = set()
    for book in clean_up_data_set:
        if len(bundle) == 5:
            bundle = set()
        #     At least 1 fiction Adult
        if book['genre'] == 'FICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5:
            bundle.add(book['title'])
        #     At least 1 Non-fiction Adult
        if book['genre'] == 'NONFICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5:
            bundle.add(book['title'])
        #     At least 1 Kid coloring
        if book['sub_genre'] == 'Colouring' and book['bos_category'] == 'KIDS' and len(bundle) < 5:
            bundle.add(book['title'])
        #     At least 1 Kid fiction
        if book['genre'] == 'FICTION' and book['bos_category'] == 'KIDS' and len(bundle) < 5:
            bundle.add(book['title'])
        adult_kids_category_boxes.append(bundle)

    # create bundles for adults only category
    bundle = set()
    for book in clean_up_data_set:
        if len(bundle) == 5:
            bundle = set()
        #     At least 1 fiction
        if book['genre'] == 'FICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5:
            bundle.add(book['title'])
        #     At least 1 Non-fiction
        if book['genre'] == 'NONFICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5:
            bundle.add(book['title'])
        #     At least and only 1 colouring Adult
        if book['sub_genre'] == 'Colouring' and book['bos_category'] == 'ADULT'and len(bundle) < 5:
            bundle.add(book['title'])
        adult_only_category_boxes.append(bundle)

    # create bundles for kids category
    bundle = set()
    for book in clean_up_data_set:
        if len(bundle) == 5:
            bundle = set()
        #     At least 2 Colouring books
        if book['genre'] == 'Colouring' and book['bos_category'] == 'KIDS' and len(bundle) < 5:
            bundle.add(book['title'])
        kids_only_category_boxes.append(bundle)

    return adult_kids_category_boxes, \
           adult_only_category_boxes, \
           kids_only_category_boxes


if __name__== "__main__":
    print read_data.__doc__
    #reading data from excel
    whole_data_from_excel = read_data()
    #print len(whole_data_from_excel)
    adult_kids_category, adult_only_category, kids_only_category = make_three_main_category_boxes(whole_data_from_excel)
    #print len(adult_only_category) + len(kids_only_category)
    for x in adult_kids_category:
        print len(x)



