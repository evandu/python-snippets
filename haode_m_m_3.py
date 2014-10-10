#coding=gbk
"""
demo:
  python haode_m_m_3.py  abs:output_path  abs:input_path(haode_m.yyyymmdd) abs:input_dir startDate(yyyymmdd)  endDate(yyymmdd)
"""
from xlwt import *
import sys
import logging
if __name__ == '__main__':
    if len(sys.argv) != 6:
        logging.error("""
        Usage: python input_name output_name
         demo
         python haode_m_m_3.py  abs:output_path  abs:input_path(haode_m.yyyymmdd) abs:input_dir startDate(yyyymmdd)  endDate(yyymmdd)
        """)
        exit(1)
    save_file_path = sys.argv[1]
    check_file_path = sys.argv[2]
    m_dir = sys.argv[3]
    import time
    startDate = time.strptime(sys.argv[4], '%Y%m%d')
    endDate = time.strptime(sys.argv[5], "%Y%m%d")
    prefix = time.strftime("%y.%m.%d", startDate) + "-" + time.strftime("%y.%m.%d", endDate)
    headings = [u'����', u'���', u'���׽��']
    book = Workbook(encoding='cp1251')
    book.add_sheet(u'���������')
    fontA = Font()
    fontA.name = u'����'
    fontA.height = 220
    styleA = XFStyle()
    styleA.font = fontA

    fontB = Font()
    fontB.name = u'����'
    fontB.bold = True
    fontB.height = 220
    styleB = XFStyle()
    styleB.font = fontB

    import codecs
    """
    month to days summary
    """
    try:
        import datetime
        sd = datetime.datetime(startDate.tm_year,startDate.tm_mon, startDate.tm_mday)
        ed = datetime.datetime(endDate.tm_year, endDate.tm_mon, endDate.tm_mday)
        m_total_data = []
        import os
        while sd <= ed:
            d_file = os.path.join(m_dir, 'haode_d.{sdtime}'.format(sdtime = sd.strftime("%Y%m%d")))
            if os.path.exists(d_file):
                logging.debug("parse ..ing " + d_file)
                fhm = codecs.open(filename=d_file, mode='r', encoding='gbk')
                map(lambda x: m_total_data.append(x[0]),
                    [[line.strip().replace(u"N0", u"NO").split(',')] for line in fhm.readlines()])
                fhm.close()
            sd = sd + datetime.timedelta(days=1)
        from itertools import groupby
        c = groupby(sorted(m_total_data, key=lambda s: s[1]), lambda f: f[1])
        p_dict = {}
        for k, g in c:
            p_dict[k] = list(g)
        m_sheet = book.add_sheet(prefix + u'�ջ��ܱ�')
        m_headings = [u'����', u'���', u'����', u'���׽��',u'�������',u'�ϼ�']
        m_rowx = 0
        for colx, value in enumerate(m_headings):
            m_sheet.write(m_rowx, colx, value, styleB)
        if len(m_total_data) == 0:
            logging.warn(u'����{prefix}�������,error: {error}'.format(prefix=prefix, error=u'����Ϊ��'))
        else:
            for p_key in p_dict:
                 c_count = m_rowx+2
                 for row in p_dict[p_key]:
                    m_rowx += 1
                    row.pop(3)
                    row.pop(4)
                    for colx, value in enumerate(row):
                        if colx == 3:
                            value = float(value)
                        m_sheet.write(m_rowx, colx, value, styleA)
                 m_rowx += 1
                 for colx, value in enumerate(['',  p_key+u' ����', '', \
                                               Formula("SUBTOTAL(9,D%d:D%d)" % (c_count, m_rowx))]):
                    if colx == 3:
                        m_sheet.write(m_rowx, colx, value, styleA)
                    else:
                        m_sheet.write(m_rowx, colx, value, styleB)
            m_rowx += 1
            for colx, value in enumerate(['',  u'�ܼ�', '', \
                                           Formula("SUBTOTAL(9,D%d:D%d)" % (2, m_rowx-1))]):
                if colx == 3:
                    m_sheet.write(m_rowx, colx, value, styleA)
                else:
                    m_sheet.write(m_rowx, colx, value, styleB)
    except Exception, e:
        logging.warn(u'����{prefix}�ջ��ܱ�,error: {error}'.format(prefix=prefix, error=e))
    """
     month to month
    """
    try:
        fh = codecs.open(filename=check_file_path, mode='r', encoding='gbk')
        data1 = map(lambda x: x[0], [[line.strip().split(',')] for line in fh.readlines()])
        fh.close()
        sheet = book.add_sheet(prefix + u'����ϸ��')
        rowx = 0
        for colx, value in enumerate(headings):
            sheet.write(rowx, colx, value, styleB)
        if len(data1) == 0:
            logging.warn(u'����{prefix}����ϸ��,error: {error}'.format(prefix=prefix, error=u'����Ϊ��'))
        else:
            for row in data1:
                rowx += 1
                for colx, value in enumerate(row):
                    if colx == 2:
                        value = float(value)
                    sheet.write(rowx, colx, value, styleA)
            rowx += 1
            for colx, value in enumerate(['',  u'�ܼ�', Formula("SUBTOTAL(9,C%d:C%d)" % (2, rowx))]):
                if colx == 2:
                    sheet.write(rowx, colx, value, styleA)
                else:
                    sheet.write(rowx, colx, value, styleB)
        logging.info(u'ִ����ɣ���鿴��'+save_file_path)
        book.add_sheet(prefix + u'�������»��ܱ�')
        book.save(save_file_path)
    except Exception, e:
        logging.error(u'����{prefix}����ϸ��,error: {error}'.format(prefix=prefix, error=e))

