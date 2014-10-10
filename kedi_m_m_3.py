#coding=gbk
"""
install  module xlwt
demo:
  python kedi_days.py  abs:output_path  abs:input_path(kedi_m.yyyymmdd) abs:input_dir startDate(yyyymmdd)  endDate(yyymmdd)
"""
from xlwt import *
import sys
import logging
if __name__ == '__main__':
    if len(sys.argv) != 6:
        logging.error("""
        Usage: python input_name output_name
        demo:
          python kedi_days.py  abs:output_path  abs:input_path(kedi_m.yyyymmdd) abs:input_dir startDate(yyyymmdd)  endDate(yyymmdd)
        """)
        exit(1)
    import time
    save_file_path = sys.argv[1]
    check_file_path = sys.argv[2]
    m_dir = sys.argv[3]
    startDate = time.strptime(sys.argv[4], '%Y%m%d')
    endDate = time.strptime(sys.argv[5], "%Y%m%d")
    prefix = time.strftime("%y.%m.%d", startDate) + "-" + time.strftime("%y.%m.%d", endDate)
    import codecs
    from itertools import groupby
    from collections import defaultdict
    book = Workbook(encoding='cp1251')
    book.add_sheet(u'差异调整表')
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

    """
    month to days summary
    """
    import datetime
    sd = datetime.datetime(startDate.tm_year,startDate.tm_mon, startDate.tm_mday)
    ed = datetime.datetime(endDate.tm_year, endDate.tm_mon, endDate.tm_mday)
    m_total_data = []
    import os
    try:
        while sd <= ed:
            d_file = os.path.join(m_dir, 'kedi_d.{sdtime}'.format(sdtime = sd.strftime("%Y%m%d")))
            if os.path.exists(d_file):
                logging.debug("parse ..ing " + d_file)
                fhm = codecs.open(filename=d_file, mode='r', encoding='gbk')
                map(lambda x: m_total_data.append(x[0]), [[line.strip().split(',')] for line in fhm.readlines()])
                fhm.close()
            sd = sd + datetime.timedelta(days=1)
        m_c = groupby(sorted(m_total_data, key=lambda s: (s[2], s[1])), lambda f: (f[2], f[1]))
        m_c_dict, m_c_p_dict = {}, {}
        for k, g in m_c:
            (m_c_dict[k[1]], m_c_p_dict[k[1]]) = (list(g), k[0])
        m_p_c_dict = defaultdict(list)
        for key in m_c_p_dict:
            m_p_c_dict[m_c_p_dict[key]].append(key)
        m_headings = [u'店名', u'编号', u'所属公司', u'日期', u'交易金额',u'调整金额',u'合计']
        m_sheet = book.add_sheet(prefix + u'日汇总表')
        m_rowx = 0
        for colx, value in enumerate(m_headings):
            m_sheet.write(m_rowx, colx, value, styleB)
        if len(m_total_data) == 0:
            logging.error(u'生成{prefix}日汇总表, error: {error}'.format(prefix=prefix, error=u'数据为空'))
        else:
            for p_key in m_p_c_dict:
                p_count = m_rowx+2
                for c_e in m_p_c_dict[p_key]:
                    c_count = m_rowx+2
                    for row in m_c_dict[c_e]:
                        m_rowx += 1
                        """
                         remove cardNo ,merCode
                        """
                        row.pop(4)
                        row.pop(5)
                        for colx, value in enumerate(row):
                            if colx == 4:
                                value = float(value)
                            m_sheet.write(m_rowx, colx, value, styleA)
                    m_rowx += 1
                    for colx, value in enumerate(['', c_e+u' 汇总', '', '',\
                                                  Formula("SUBTOTAL(9,E%d:E%d)" % (c_count, m_rowx))]):
                        """
                         money diff style
                        """
                        if colx == 4:
                            m_sheet.write(m_rowx, colx, value, styleA)
                        else:
                            m_sheet.write(m_rowx, colx, value, styleB)
                m_rowx += 1
                for colx, value in enumerate(['', '',  p_key+u' 汇总', '', \
                                              Formula("SUBTOTAL(9,E%d:E%d)" % (p_count, m_rowx-1))]):
                    if colx == 4:
                        m_sheet.write(m_rowx, colx, value, styleA)
                    else:
                        m_sheet.write(m_rowx, colx, value, styleB)
            m_rowx += 1
            for colx, value in enumerate(['', '', u'总计', '',\
                                          Formula("SUBTOTAL(9,E%d:E%d)" % (2, m_rowx-2))]):
                if colx == 4:
                    m_sheet.write(m_rowx, colx, value, styleA)
                else:
                    m_sheet.write(m_rowx, colx, value, styleB)
    except Exception, e:
        logging.error(u'生成{prefix}日汇总表,error: {error}'.format(prefix=prefix, error=e))
    """
     month to month
    """
    try:
        headings = [u'店名', u'编号', u'所属公司', u'交易金额']
        fh = codecs.open(filename=check_file_path, mode='r', encoding='gbk')
        data1 = map(lambda x: x[0], [[line.strip().split(',')] for line in fh.readlines()])
        fh.close()
        c = groupby(data1, lambda f: (f[2], f[1]))
        c_dict, c_p_dict = {}, {}
        for k, g in c:
            (c_dict[k[1]], c_p_dict[k[1]]) = (list(g), k[0])
        p_c_dict = defaultdict(list)
        for key in c_p_dict:
            p_c_dict[c_p_dict[key]].append(key)
        sheet = book.add_sheet(prefix + u'月明细表')
        rowx = 0
        for colx, value in enumerate(headings):
            sheet.write(rowx, colx, value, styleB)
        if len(data1) == 0:
            logging.error(u'生成{prefix}月汇总表报错,error: {error}'.format(prefix=prefix, error=u'数据为空'))
        else:
            for p_key in p_c_dict:
                p_count = rowx + 2
                for c_e in p_c_dict[p_key]:
                    c_count = rowx+2
                    for row in c_dict[c_e]:
                        rowx += 1
                        """
                            money diff style
                        """
                        for colx, value in enumerate(row):
                            if colx == 3:
                                value = float(value)
                            sheet.write(rowx, colx, value, styleA)
                    rowx +=1
                    for colx, value in enumerate(['', c_e+u' 汇总', '', Formula("SUBTOTAL(9,D%d:D%d)" % (c_count, rowx))]):
                        if colx == 3:
                            sheet.write(rowx, colx, value, styleA)
                        else:
                            sheet.write(rowx, colx, value, styleB)
                rowx += 1
                for colx, value in enumerate(['', '',  p_key+u' 汇总', Formula("SUBTOTAL(9,D%d:D%d)" % (p_count, rowx))]):
                    if colx == 3:
                        sheet.write(rowx, colx, value, styleA)
                    else:
                        sheet.write(rowx, colx, value, styleB)
            rowx += 1
            for colx, value in enumerate(['', '', u'总计', Formula("SUBTOTAL(9,D%d:D%d)" % (2, rowx-1))]):
                if colx == 3:
                    sheet.write(rowx, colx, value, styleA)
                else:
                    sheet.write(rowx, colx, value, styleB)
        logging.info(u'执行完成，请查看：'+save_file_path)
        book.add_sheet(prefix + u'手续费月汇总表')
        book.save(save_file_path)
    except Exception, e:
        logging.error(u'生成{prefix}月汇总表报错,error: {error}'.format(prefix=prefix, error=e))


