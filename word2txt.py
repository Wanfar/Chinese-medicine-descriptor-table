import xlrd

# word=dict = {'叙词': ' ', '首字笔画数': '', '功效': '',
#              'Y':'','D':'','S':'','F':'','C':''}
def extract_word_info(file_path):
    word_list=[]
    x1 = xlrd.open_workbook(file_path)
    sheet = x1.sheet_by_name('Sheet1')
    row_Num = sheet.nrows
    col_Num = sheet.ncols

    tittle=sheet.row_values(0)#key
    j=1
    for i in range(row_Num-1):
        word_dict = dict()
        values = sheet.row_values(j)
        for x in range(col_Num):
            word_dict[tittle[x]]=values[x]
        j+=1
        word_list.append(word_dict)
    return word_list

def print_word(word_list):
    fw = open("叙词.txt", 'w+', encoding='utf-8')
    for i in word_list:
        fw.write('\n'+i['叙词']+'\n       '+i['首字笔画数']+'  '+i['功效']+'\n')
        if(i['D']):
            i['D'] = i['D'].replace('、', '\n' + '   ')
            fw.write('D'+'  '+i['D']+'\n')
        if (i['F']):
            i['F']=i['F'].replace(',', '\n'+'   ')
            i['F'] = i['F'].replace('、', '\n' + '   ')
            fw.write('F' +'  '+ i['F'] + '\n')
        i['C']=i['C'].replace(',', '\n'+'   ')
        i['C'] = i['C'].replace('、', '\n' + '   ')
        fw.write('C' + '  ' + i['C'] + '\n')
    fw.close()


file_path = '中药叙词表_all.xlsx'
words=extract_word_info(file_path)
print_word(words)