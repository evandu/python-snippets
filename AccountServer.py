#!/usr/bin/env python
#coding=utf8

import BaseHTTPServer
import httplib
import datetime
import logging

global ACCOUNT_BALANCE
ACCOUNT_BALANCE = None

host = "180.166.52.155"
userId = ""
port = 8080

"""
  init load query smartpay balance
"""


def refresh_account_balance():
    global ACCOUNT_BALANCE
    http_client = None
    try:
        logging.info("refresh_account_balance .......... ")
        http_client = httplib.HTTPConnection(host, 8080, timeout=30)
        http_client.request('GET', '/mptopup/agent/agentApiQuerySva.htm?userId=20006221')
        response = http_client.getresponse()
        xml = response.read()
        logging.info("response {status}, reason = {reason},{xml}"
                     .format(status=response.status, reason=response.reason, xml=xml))
        if response.status == 200:
            ACCOUNT_BALANCE = AccountBalance(xml, datetime.datetime.now())
    finally:
        if http_client:
            http_client.close()


class AccountBalance:
    def __init__(self, xml, date):
        self.xml = xml
        self.date = date

    def is_balance_out_date(self):
        return ('<errcode>1000</errcode>' not in self.xml) or (datetime.datetime.now() - self.date).seconds > 6*60

"""
http server
"""


class WebRequestHandler(BaseHTTPServer.BaseHTTPRequestHandler):
    """
      proxy smartpay account balance
    """
    def do_GET(self):
        self.send_response(200)
        self.timeout = 10
        self.end_headers()
        if ACCOUNT_BALANCE is None or ACCOUNT_BALANCE.is_balance_out_date():
            print "balance is out date"
            refresh_account_balance()
        self.wfile.write(ACCOUNT_BALANCE.xml)


if __name__ == '__main__':
    refresh_account_balance()
    server = BaseHTTPServer.HTTPServer(('0.0.0.0', port), WebRequestHandler)
    server.serve_forever()