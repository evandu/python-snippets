#coding=gbk
"""
save_file_path, check_file_path, sheet_name
install  module xlwt
demo:
  python kedi_days.py  save_file_path  check_file_path  sheet_name
"""
from xlwt import *
import sys
import logging
if __name__ == '__main__':
    if len(sys.argv) != 4:
        logging.error("""
          Usage: python input_name output_name
          demo:
                python kedi_days.py  save_file_path  check_file_path  sheet_name
          """)
        exit(1)
    save_file_path = sys.argv[1]
    check_file_path = sys.argv[2]
    sheet_name = sys.argv[3]
    fontA = Font()
    fontA.name = u'宋体'
    fontA.height = 220
    styleA = XFStyle()
    styleA.font = fontA

    fontB = Font()
    fontB.name = u'宋体'
    fontB.bold = True
    fontB.height = 220
    styleB = XFStyle()
    styleB.font = fontB
    headings = [u'店名', u'编号', u'所属公司', u'日期', u'e卡通号', u'交易金额',u'商户号']
    import codecs
    from itertools import groupby
    from collections import defaultdict
    fh = codecs.open(filename=check_file_path, mode='r', encoding='gbk')
    data1 = map(lambda x: x[0], [[line.strip().split(',')] for line in fh.readlines()])
    fh.close()
    c = groupby(sorted(data1, key=lambda s: s[1]), lambda f: (f[2], f[1]))
    c_dict, c_p_dict = {}, {}
    for k, g in c:
        (c_dict[k[1]], c_p_dict[k[1]]) = (list(g), k[0])
    p_c_dict = defaultdict(list)
    for key in c_p_dict:
        p_c_dict[c_p_dict[key]].append(key)
    book = Workbook(encoding='cp1251')
    sheet = book.add_sheet(sheet_name + u'日明细表')
    rowx = 0
    for colx, value in enumerate(headings):
        sheet.write(rowx, colx, value, styleB)
    if len(data1) == 0:
        logging.error('check file is empty' + check_file_path)
    else:
        for p_key in p_c_dict:
            p_count = rowx + 2
            for c_e in p_c_dict[p_key]:
                c_count = rowx+2
                for row in c_dict[c_e]:
                    rowx += 1
                    for colx, value in enumerate(row):
                        if colx == 5:
                            value = float(value)
                        sheet.write(rowx, colx, value, styleA)
                rowx += 1
                for colx, value in enumerate(['', c_e+u' 汇总', '', '', '',\
                                              Formula("SUBTOTAL(9,F%d:F%d)" % (c_count, rowx)), '']):
                    if colx == 5:
                        sheet.write(rowx, colx, value, styleA)
                    else:
                        sheet.write(rowx, colx, value, styleB)
            rowx += 1
            for colx, value in enumerate(['', '',  p_key+u' 汇总', '', '',\
                                          Formula("SUBTOTAL(9,F%d:F%d)" % (p_count, rowx-1)), '']):
                if colx == 5:
                    sheet.write(rowx, colx, value, styleA)
                else:
                    sheet.write(rowx, colx, value, styleB)
        rowx += 1
        for colx, value in enumerate(['', '', u'总计', '', '',\
                                      Formula("SUBTOTAL(9,F%d:F%d)" % (2, rowx-2)), '']):
            if colx == 5:
                sheet.write(rowx, colx, value, styleA)
            else:
                sheet.write(rowx, colx, value, styleB)
    logging.info(u'执行完成，请查看：'+save_file_path)
    book.add_sheet(u"差异调整表")
    book.save(save_file_path)

