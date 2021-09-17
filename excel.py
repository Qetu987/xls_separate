import xlrd
import xlwt

workbook = xlrd.open_workbook('Excel.xls')
workbook2 = xlrd.open_workbook('excel2.xls')


# get names from elfy file
def get_names_elfy(data):
    worksheet = data.sheet_by_index(0)
    names = []
    for rx in range(1, worksheet.nrows):
        names.append(worksheet.cell_value(rowx=rx, colx=1).lower())
    return names


# hmmm separate names from-data table
def get_names_data(data):
    worksheet = data.sheet_by_index(0)
    names = {}
    for rx in range(1, worksheet.nrows):
        names[rx] = worksheet.cell_value(rowx=rx, colx=4).lower() + ' ' + \
                    worksheet.cell_value(rowx=rx, colx=5).lower() + ' ' + \
                    worksheet.cell_value(rowx=rx, colx=6).lower()
    return names


# get create dict with index value curent name from parent table
def filter_list(list_elfy, dict_data):
    data = {}
    for index, val in dict_data.items():
        if val in list_elfy:
            data[index] = val
    return data


# create dict with data from data-table (key is title of colum)
def find_data(data, workbook_data):
    worksheet = workbook_data.sheet_by_index(0)
    SBK_FIO = {}
    SBK_INN = {}
    SBK_NUM = {}
    SBK_SUM = {}
    IBAN_NUM = {}
    for index, val in data.items():
        SBK_FIO[index] = val.title()
        SBK_INN[index] = worksheet.cell_value(rowx=index, colx=3)
        SBK_NUM[index] = worksheet.cell_value(rowx=index, colx=7)
        IBAN_NUM[index] = worksheet.cell_value(rowx=index, colx=8)
    return {'SBK_FIO':SBK_FIO , 'SBK_INN':SBK_INN, 'SBK_NUM':SBK_NUM, 'IBAN_NUM':IBAN_NUM}
        

def separate(val_data):
    data = {}
    for index, SBK_FIO in val_data['SBK_FIO'].items():
        data[index] = {'SBK_FIO': SBK_FIO}

    for index, SBK_INN in val_data['SBK_INN'].items():
        data[index].update({'SBK_INN': SBK_INN})

    for index, SBK_NUM in val_data['SBK_NUM'].items():
        data[index].update({'SBK_NUM': SBK_NUM})

    for index, IBAN_NUM in val_data['IBAN_NUM'].items():
        data[index].update({'IBAN_NUM': IBAN_NUM})

    return data


def add_celery_saparate(data, elfy_workbook):
    worksheet = elfy_workbook.sheet_by_index(0)
    elfy_data = {}
    for rx in range(1, worksheet.nrows):
        elfy_data[worksheet.cell_value(rowx=rx, colx=1).title()] = worksheet.cell_value(rowx=rx, colx=2)
    for item in data.values():
        item.update({'SBK_SUM': elfy_data[item['SBK_FIO']]})
    return data


# create new file xls
def create_data(data):
    row = 1
    
    writebook = xlwt.Workbook()
    sheet = writebook.add_sheet("Data")
    
    sheet.write(0, 0, 'SBK_FIO')
    sheet.write(0, 1, 'SBK_INN')
    sheet.write(0, 2, 'SBK_NUM')
    sheet.write(0, 3, 'SBK_SUM')
    sheet.write(0, 4, 'IBAN_NUM')

    for item in data.values():
        print(item)
        sheet.write(row, 0, item['SBK_FIO'])
        sheet.write(row, 1, item['SBK_INN'])
        sheet.write(row, 2, item['SBK_NUM'])
        sheet.write(row, 3, item['SBK_SUM'])
        sheet.write(row, 4, item['IBAN_NUM'])
        row += 1    

    writebook.save('contacts.xls')

    


data_set = filter_list(get_names_elfy(workbook), get_names_data(workbook2))
f = find_data(data_set, workbook2)
sort_data = separate(f)
correct_data = add_celery_saparate(sort_data, workbook)
create_data(correct_data)

