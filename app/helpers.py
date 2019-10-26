from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from flask import render_template, send_file

def gen_file(index_rows):
    with open("file.html", "w") as f:
        #f = open('index_assiut.html', 'w')
        f.write(render_template('main_index.html', row_list=index_rows))
        #f.filename = 'indexmmm.html'
        #f.close()
        return f



def index_init():
    file = load_workbook('app/templates/079254C-0000-ML-100_0 - SECT. 5.xlsx')
    sheet_list = file.sheetnames
    
    new_file = Workbook()
    new_ws = new_file.active
    print('')
    #print(main_folder)
    print('')
    index_rows=[]
    #f = open('index_assiut.html', 'w')
    for sheet in sheet_list:
        sheet = file[sheet]
        sec1 = sheet['B13'].value
        sec2 = sheet['B15'].value
        sec3 = sheet['B17'].value
        for row in sheet.iter_rows(min_row=2):
            cell = row[2]
            if cell.value is not None and cell.value[0:7] == '079254C':
                '''
                cell.hyperlink = Hyperlink('link',
                        target='5.1.1/' + cell.value + "_" + row[3].value,
                        #location='5.1.1/' + cell.value + "_" + row[3].value,
                        display='something')
                '''
                cell_start_value = cell.value
                #cell.value = '=HYPERLINK("5.1.1\\' + cell_start_value + "_" + row[3].value +'"'+',"' + cell.value + '")'
                #cell.value = "=HYPERLINK('5.1.1/" + cell.value + "_" + row[3].value + "','" + cell.value + "')"
                cell.hyperlink = Hyperlink('link',
                        target='5.1.1\\' + cell_start_value,
                        #location='5.1.1/' + cell.value + "_" + row[3].value,
                        display='something')
                print(cell.value)
                link1 = row[3]
                link1.hyperlink = Hyperlink('link',
                        target='5.1.1\\' + cell_start_value,
                        #location='5.1.1/' + cell.value + "_" + row[3].value,
                        display='something')
                
                index_rows.append((cell_start_value,cell.value))
                new_ws.append((sec1,sec2,sec3,row[1].value,cell_start_value,row[3].value))  

    file.save('test2.xlsx')
    new_file.save('testof1.xlsx')
    '''
    f = gen_file(index_rows)
    
    return send_file(f, as_attachment=True, attachment_filename='indexmain.html')
    '''
 