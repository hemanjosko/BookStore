# -*- coding: utf-8 -*-
import xlrd, xlwt
import sys

def read_data():
    '''Read data from excel sheet and make dictionary...

    '''

    #location of of excel sheet data
    loc = ("resource/BookList.xlsx")

    #open excel sheet
    wb = xlrd.open_workbook(loc)

    #all first sheet data
    sheet = wb.sheet_by_index(0)

    headers = ['id','used']
    whole_data = []
    dict_whole_data = []

    i = 0
    for c_rows in range(0,sheet.nrows):
        m_data = [i,0]
        for c_cols in range(0,sheet.ncols):
            if c_rows == 0:
                headers.append((sheet.cell_value(c_rows,c_cols))
                               .replace(" ","_").replace("-","_").lower())
            else:
                m_data.append(sheet.cell_value(c_rows,c_cols))
        i += 1
        if c_rows > 0:
            whole_data.append(m_data)


    for x in whole_data:
        a = dict(zip(headers,x))
        dict_whole_data.append(a)

    # print "Headers"
    # print headers

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
    adult_kids_category_box = []
    adult_only_category_box = []
    kids_only_category_box = []
    rest_category_box = []

    #only main two category data adult and kids
    clean_up_data_set = []
    for data in data_set:
        if data['bos_category'] in ('ADULT', 'KIDS'):
            clean_up_data_set.append(data)
    print 'Total books lies between adults and kids category: %s' % len(clean_up_data_set)

    # create bundles for adults and kids category
    ####################################################################################################################
    bundle = set()
    weight = []
    clean_me = []
    validator = set()
    #print 'start adult books out of adult kids'
    for book in clean_up_data_set:
        #in this loop only adult bundle we create
        #now in this section only removing entries while adding to bundle remain from data set
        if len(bundle) == 5:
            adult_kids_category_box.append(bundle)
            validator = set()
            bundle = set()
            weight = []
            clean_me = []
        #     At least 1 fiction Adult
        if book['genre'] == 'FICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5 and book['used'] == 0 \
                and 'one' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('one')
                book['used'] = 1
                clean_me.append(book['id'])
        #     At least 1 Non-fiction Adult
        if book['genre'] == 'NONFICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5 and book['used'] == 0 \
                and 'two' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('two')
                book['used'] = 1
                clean_me.append(book['id'])
        if len(validator) < 5 and book['bos_category'] == 'ADULT' and book['used'] == 0:# if your validation says yours bundle has less than 5 books then you can add more book
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                book['used'] = 1
                clean_me.append(book['id'])
    else:
        #completed loop remove books from bundle if any
        # print bundle
        # print validator
        # print len(adult_kids_category_boxes)
        # print sum(weight)
        #print clean_me
        for clean in clean_me:
            for clean_data in clean_up_data_set:
                if clean_data['id'] == clean:
                    clean_data['used'] = 0

        #print 'complete adult loop out of adult kids'

    bundle = set()
    weight = []
    clean_me = []
    validator = set()
    #print 'start kids out of kids adult books'
    for book in clean_up_data_set:  # now in this section only removing entries while adding to bundle remain from data set
        if len(bundle) == 5:
            adult_kids_category_box.append(bundle)
            validator = set()
            bundle = set()
            weight = []
            clean_me = []
        #     At least 1 Kid coloring
        if book['sub_genre'] == 'Colouring' and book['bos_category'] == 'KIDS' and len(bundle) < 5 and book['used'] == 0 \
                and 'three' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('three')
                book['used'] = 1
                clean_me.append(book['id'])
        #     At least 1 Kid fiction
        if book['genre'] == 'FICTION' and book['bos_category'] == 'KIDS' and len(bundle) < 5 and book['used'] == 0 \
                and 'four' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('four')
                book['used'] = 1
                clean_me.append(book['id'])
        if len(validator) < 5 and book['bos_category'] == 'KIDS' and book['used'] == 0:  # if your validation says yours bundle has less than 5 books then you can add more book
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                book['used'] = 1
                clean_me.append(book['id'])
    else:
        # completed loop remove books from bundle if any
        # print bundle
        # print validator
        # print len(adult_kids_category_boxes)
        # print sum(weight)
        #print clean_me
        for clean in clean_me:
            for clean_data in clean_up_data_set:
                if clean_data['id'] == clean:
                    clean_data['used'] = 0
        #print 'complete kids out of kids adult loop'
    #######################################################################################################################

    # create bundles for adults only category
    bundle = set()
    weight = []
    clean_me = []
    validator = set()
    #print 'start adult loop'
    for book in clean_up_data_set:
        if len(bundle) == 5:
            adult_only_category_box.append(bundle)
            validator = set()
            bundle = set()
            weight = []
            clean_me = []
        #     At least 1 fiction
        if book['genre'] == 'FICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5 and book['used'] == 0 \
                and 'one' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('one')
                book['used'] = 1
                clean_me.append(book['id'])
        #     At least 1 Non-fiction
        if book['genre'] == 'NONFICTION' and book['bos_category'] == 'ADULT' and len(bundle) < 5 and book['used'] == 0 \
                and 'two' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('two')
                book['used'] = 1
                clean_me.append(book['id'])
        #     At least and only 1 colouring Adult
        if book['sub_genre'] == 'Colouring' and book['bos_category'] == 'ADULT' and len(bundle) < 5 and book['used'] == 0 \
                and 'three' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('three')
                book['used'] = 1
                clean_me.append(book['id'])
        if len(validator) < 5 and book['bos_category'] == 'ADULT' and book['used'] == 0:  # if your validation says yours bundle has less than 5 books then you can add more book
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                book['used'] = 1
                clean_me.append(book['id'])
    else:
        # completed loop remove books from bundle if any
        # print bundle
        # print validator
        # print len(adult_only_category_boxes)
        # print sum(weight)
        #print clean_me
        for clean in clean_me:
            for clean_data in clean_up_data_set:
                if clean_data['id'] == clean:
                    clean_data['used'] = 0
        #print 'complete adult loop'

    ####################################################################################################################

    # create bundles for kids category
    bundle = set()
    weight = []
    clean_me = []
    validator = set()
    #print 'start kids loop'
    for book in clean_up_data_set:  # now in this section only removing entries while adding to bundle remain from data set
        if len(bundle) == 5:
            kids_only_category_box.append(bundle)
            validator = set()
            bundle = set()
            weight = []
            clean_me = []
        #     At least 2 Colouring books
        if book['genre'] == 'Colouring' and book['bos_category'] == 'KIDS' and len(bundle) < 5 and book['used'] == 0 \
                and 'one' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('one')
                book['used'] = 1
                clean_me.append(book['id'])
        elif book['genre'] == 'Colouring' and book['bos_category'] == 'KIDS' and len(bundle) < 5 and book['used'] == 0 \
                and 'two' not in validator:
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                validator.add('two')
                book['used'] = 1
                clean_me.append(book['id'])
        if len(validator) < 5 and book['bos_category'] == 'KIDS' and book['used'] == 0:  # if your validation says yours bundle has less than 5 books then you can add more book
            if float(sum(weight)) + float(book['stock_weight']) < 1800:
                bundle.add(book['title'])
                weight.append(float(book['stock_weight']))
                book['used'] = 1
                clean_me.append(book['id'])
    else:
        # completed loop remove books from bundle if any
        # print bundle
        # print validator
        # print len(kids_only_category_boxes)
        # print sum(weight)
        #print clean_me
        for clean in clean_me:
            for clean_data in clean_up_data_set:
                if clean_data['id'] == clean:
                    clean_data['used'] = 0
        #print 'complete kids loop'
    # print 'test start'
    # for clean in clean_me:
    #     for clean_data in clean_up_data_set:
    #         if clean_data['id'] == clean:
    #             print clean_data
    # print 'test complete'
    #print 'start keeping seperate books in a rest box'
    for book in clean_up_data_set:
        if book['used'] == 0:
            rest_category_box.append(book['title'])
    return adult_kids_category_box, \
           adult_only_category_box, \
           kids_only_category_box, \
           rest_category_box


if __name__== "__main__":
    print read_data.__doc__
    #reading data from excel
    whole_data_from_excel = read_data()
    # for x in whole_data_from_excel:
    #     print x
    #print len(whole_data_from_excel)
    adult_kids_category, adult_only_category, kids_only_category, rest_category = make_three_main_category_boxes(whole_data_from_excel)
    #print len(adult_only_category) + len(kids_only_category)
    a = len(adult_kids_category)
    b = len(adult_only_category)
    c = len(kids_only_category)
    d = len(rest_category)

    sys.stdout = open('output/file', 'w')

    print '\n1. Kids and adult box category has (%s)Bundles\n' % a
    bundle = 1
    for x in adult_kids_category:
        print "Bundle %s" % bundle
        for y in x:
            print y.encode("utf8")
        print "\n"
        bundle += 1

    print '\n2. Adult box ategory has (%s)Bundles\n' % b
    bundle = 1
    for x in adult_only_category:
        print "Bundle %s" % bundle
        for y in x:
            print y.encode("utf8")
        print "\n"
        bundle += 1

    print '\n3. Kids box category has (%s)Bundles\n' % c
    bundle = 1
    for x in kids_only_category:
        print "Bundle %s" % bundle
        for y in x:
            print y.encode("utf8")
        print "\n"
        bundle += 1

    print "------------------------------------------------------"

    print 'Total bundles count of kids and adult box category: %s' % a
    print 'Total bundles count of adult box category: %s' % b
    print 'Total bundles count of kids box category: %s' % c
    print 'Total count of rest books: %s' % d
    # for x in adult_kids_category:
    #    print len(x)
    print '%s + %s + %s = %s bundles' % (a,b,c,a+b+c)
    print '%s x 5 = %s books' % (a+b+c,(a+b+c)*5)
    print '%s + %s: %s Total Books' % ((a+b+c)*5,d,((a+b+c)*5)+d)

    print "------------------------------------------------------"