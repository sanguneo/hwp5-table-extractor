from hwp5_table import HwpFile
import json
import csv
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def cli(path):
    hwp = HwpFile(path)
    tables = []
    section_idx = 0
    while hwp.ole.exists('BodyText/Section%d' % section_idx):
        tables.extend(hwp.get_tables(section_idx))
        section_idx += 1
    doc = []
    for table in tables:
        onetable = []
        table.rows.pop(0)
        for row in table.rows:
            onerow = []
            for cell in row:
                onerow.append("".join(cell.lines))
            onetable.append(onerow)
        doc = doc + onetable
    return doc


def gethering(dirname, list):
    filenames = os.listdir(dirname)
    for filename in filenames:
        full_filename = os.path.join(dirname, filename)
        if os.path.isdir(full_filename):
            gethering(full_filename, list)
        else:
            ext = os.path.splitext(full_filename)[-1]
            if ext == '.hwp':
                list.append(full_filename)

def listing(dirname) :
    list = []
    gethering(dirname, list)
    return list

def processing(dirname) :
    ret = [['번호','이름','일련번호','날짜','끝날짜','회사','비고']]
    filelist = listing(dirname)
    for file in filelist:
        ret += cli(file)
    return [ret,len(filelist)]

def tojson (dirname):
    whole = processing(dirname)
    with open(os.path.join(dirname, 'result.json'), mode='w+', encoding='utf8') as json_file:
        json.dump(whole[0], json_file, ensure_ascii=False, indent="\t")

    return [len(whole[0]), whole[1]]

def tocsv (dirname):
    whole = processing(dirname)
    with open('result.csv', 'w+') as csvfile:
        csvv = csv.writer(csvfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        for row in whole[0]:
            csvv.writerow(row)
    return [len(whole[0]),whole[1]]


if __name__ == '__main__':
    result = tocsv(BASE_DIR)
    print('처리한 파일 수 : {} \n처리된 데이터 수 : {}'.format(result[1], result[0]))