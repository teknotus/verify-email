#!/usr/bin/python

import os
import sys
import email.parser
import dateutil.parser
import dkim
from openpyxl import Workbook
from openpyxl.compat import range

wb = Workbook()
xl_filename = '../verify.xlsx'
ws = wb.active
ws.title = "email checks"
ws.append(['id', 'verified', 'date', 'date text', 'from', 'to', 'subject', 'message-id'])

for filesplit in sorted(map(os.path.splitext, os.listdir(os.getcwd())),key = lambda f: int(f[0])):
    row = [int(filesplit[0])]
    filename = "".join(filesplit)
    print filename
    f = open(filename, 'r')
    data = f.read()
    try:
        verified = dkim.verify(data)
        if verified:
            print "verified"
            row.append("verified")
        else:
            print "failed"
            row.append("failed")
    except:
        print "verify exception"
        pass

    try:
        msg = email.message_from_string(data)
        date = dateutil.parser.parse(msg['date'])
        meta = [date, msg['date'], msg['from'], msg['to'], msg['subject'], msg['message-id']]
        row += meta
    except:
        print "parse email exception"
        pass

    try:
        ws.append(row)
    except:
        print "worksheet append exception"
        pass

wb.save(xl_filename)

