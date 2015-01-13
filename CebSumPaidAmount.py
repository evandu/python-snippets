#coding=gbk


def read_check_file(start_date, end_date):
    import os
    import linecache
    sd = datetime.datetime(start_date.tm_year, start_date.tm_mon, start_date.tm_mday)
    ed = datetime.datetime(end_date.tm_year, end_date.tm_mon, end_date.tm_mday)
    while sd <= ed:
        d_file = os.path.join(".", u'{strdate}_xxxx_suc.txt'.format(strdate=sd.strftime("%Y%m%d")))
        if os.path.exists(d_file):
            data = linecache.getline(d_file, 2).split("|")
            yield (sd.strftime("%Y%m"), int(data[1]), int(data[3]))
        sd += datetime.timedelta(days=1)


def read_merge(start_date, end_date):
    from itertools import groupby

    for date, m_data in groupby(sorted(read_check_file(start_date, end_date), key=lambda s: s[0]), lambda f: f[0]):
        yield reduce(lambda x, y: (x[0], x[1] + y[1], x[2] + y[2]), list(m_data))


def println(x):
    print "{date} 笔数:{count}，金额:{sum} 元".format(date=x[0], count=x[1], sum=x[2]/100.00)
    return x[1], x[2]

import logging
import sys

if __name__ == '__main__':
    if len(sys.argv) != 3:
        logging.error(
            """
             Usage: python CebSumPaidAmount start_date(yyyymmdd) end_date(yyyymmdd)
            """)
        exit(1)
    import time
    import datetime
    startDate = time.strptime(sys.argv[1], '%Y%m%d')
    endDate = time.strptime(sys.argv[2], "%Y%m%d")
    if startDate > endDate:
        logging.error(
            """
              ||startDate  <  endDate
            """)
        exit(1)
    d = reduce(lambda x, y: (x[0] + y[0], x[1] + y[1]), map(println, list(read_merge(startDate, endDate))))
    print "合计：笔数:{count}，金额:{amount_sum} 元".format(
        startDate=sys.argv[1], endDate=sys.argv[2], count=d[0], amount_sum=d[1]/100.00)