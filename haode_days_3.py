#coding=gbk
"""
save_file_path, check_file_path, sheet_name
install  module xlwt
demo:
  python haode_days_3.py  save_file_path  check_file_path sheet_name
"""
from xlwt import *
import sys
import logging
if __name__ == '__main__':
    if len(sys.argv) != 4:
        logging.error("""
           Usage: python input_name output_name
           demo  python haode_days_3.py  save_file_path  check_file_path  sheet_name
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
    headings = [u'店名', u'编号', u'日期', u'e卡通号', u'交易金额',u'商户号']
    import codecs
    from itertools import groupby
    fh = codecs.open(filename=check_file_path, mode='r', encoding='gbk')
    data1 = map(lambda x: x[0], [[line.strip().replace(u"N0", u"NO").split(',')] for line in fh.readlines()])
    fh.close()
    c = groupby(sorted(data1, key=lambda s: s[1]), lambda f: f[1])
    p_dict = {}
    for k, g in c:
        p_dict[k] = list(g)
    book = Workbook(encoding='cp1251')
    sheet = book.add_sheet(sheet_name + u'日明细表')
    rowx = 0
    for colx, value in enumerate(headings):
        sheet.write(rowx, colx, value, styleB)
    if len(data1) == 0:
        logging.warning(u'生成{prefix}日明细表,error: {error}'.format(prefix=sheet_name, error=u'数据为空'))
    else:
        for p_key in p_dict:
             c_count = rowx+2
             for row in p_dict[p_key]:
                rowx += 1
                for colx, value in enumerate(row):
                        if colx == 4:
                            value = float(value)
                        sheet.write(rowx, colx, value, styleA)
             rowx += 1
             for colx, value in enumerate(['',  p_key+u' 汇总', '', '',\
                                           Formula("SUBTOTAL(9,E%d:E%d)" % (c_count, rowx)), '']):
                if colx == 4:
                    sheet.write(rowx, colx, value, styleA)
                else:
                    sheet.write(rowx, colx, value, styleB)
        rowx += 1
        for colx, value in enumerate(['',  u'总计', '', '',\
                                       Formula("SUBTOTAL(9,E%d:E%d)" % (2, rowx-1)), '']):
            if colx == 4:
                sheet.write(rowx, colx, value, styleA)
            else:
                sheet.write(rowx, colx, value, styleB)
    logging.info(u'执行完成，请查看：'+save_file_path)
    book.add_sheet(u"差异调整表")
    book.save(save_file_path)
