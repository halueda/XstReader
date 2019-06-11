#!/usr/bin/env python
# coding: utf-8

import os
import pathlib
import glob
import re
import datetime as dt
import pytz
from dateutil.tz import tzlocal

def no_CtrlM( text ) :
    # ^M をすべて除く
    # ここでやるのが正しいのか、XstReaderでやるのが正しいのか？
    if isinstance(text, str) :
        text = text.replace("\r", '')
    return text

def reformat_header(header):
    header1 = []
    cur = header[0]
    for l in header[1:]:
        if re.match(r" [a-zA-Z][-_a-zA-Z0-9]*:", l) :
            header1.append( cur )
            cur = l[1:]
        elif re.match(r" [ \t]", l) :
            header1.append( cur )
            cur = l[1:]
        else :
            cur += ";"+l
    header1.append( cur )
    return header1


def csv2eml( f, outdir ):
    import pandas as pd

    sub_outdir = outdir / f.stem

    columns_needed = ('Subject', 'ClientSubmitTime', 'DisplayTo', 'TransportMessageHeaders', 'DisplayCc', 'SenderEmailAddress',
                      '0x1000', '0x1013', '0x1016', '0x1009', 'UserEntryId')

    data = pd.read_csv(f, parse_dates=True, infer_datetime_format=True, usecols=(lambda c: c in columns_needed) )

    # Subject などのカラムがないことがある。
    for c in columns_needed:
        if c not in data.columns :
            data[c] = pd.np.nan

    # utf-8 header encoding.
    from email.header import Header
    def convert_header(txt):
        if isinstance(txt, str):
#            txt1 = txt.encode().decode("utf-8", errors="replace")
            header = Header(txt, 'utf-8')
            txt2 = header.encode()
            return txt2
        else :
            return ''

    # ここでするのがよいか？ヘッダーがないことの方が多いので、出力先でやった方がよいかも
    # fillna だけここでやっておくという手もある data.fillna('', inplace=True)
#    data['alt_header'] = (
#            "Subject: " + data['Subject'].fillna('').map(convert_header) + "\n" +
#            "Date: "    + data['ClientSubmitTime'].fillna('') +"\n" +              # Wed, 29 May 2019 03:59:29 +0000 形式じゃない...
#            "From: "    + data['SenderEmailAddress'].fillna('') +"\n" +
#            "Cc: "      + data['DisplayCc'].fillna('').map(convert_header) +"\n" +
#            "To: "      + data['DisplayTo'].fillna('').map(convert_header) + "\n"  # DisplayTo は Outlook仕様で ; 区切り。emlは , 区切り。メールアドレスがない場合も多い。Exchangeスペシャルで、アドレスを変換することも考える
#        )
#
#    data0=data.fillna({'TransportMessageHeaders': data['alt_header']})

    (col_plain_text,
     col_html_base64,
     col_html_utf8,
     col_rtf,
     col_header,
     col_summary) = ( '0x1000',
                      '0x1013',
                      '0x1016',
                      '0x1009',
                      'TransportMessageHeaders',
                      'UserEntryId')
#    data1 = data0.loc[:, (col_plain_text, col_html_base64, col_html_utf8, col_rtf, col_header, col_summary)] # text、html(base64), html(utf-8), RTF
#    data1 = data0

#    os.makedirs(sub_outdir, exist_ok=True)
    sub_outdir.mkdir(parents=True, exist_ok=True)

    BOUNDARY="CSV2EML-3.141592653589793-2.718281828459045" # it should be make sure this does not happen in contents. Assume this never happen to occur in contents.
    os.linesep="\n"
    for index in range(0, len(data)):
        a_line = data.loc[index, : ]
        with (sub_outdir / ('%04d.eml'%index)).open( 'wt',encoding='utf-8',errors='replace') as fout:
            # output header
            header = a_line.TransportMessageHeaders
#            if True :
            if pd.isnull(header) :
                header1 = ""
                # ヘッダ一つずつ、if文とかで分岐しながら、convert_header() とかもかけて、header1変数に追記する。
                if not pd.isnull(a_line.Subject) :
                    header1 += "Subject: " + convert_header(a_line.Subject) + "\n"
                if not pd.isnull(a_line.ClientSubmitTime) :
                    date = dt.datetime.strptime(a_line.ClientSubmitTime, '%Y/%m/%d %H:%M:%S')
                    date = date.replace(tzinfo=tzlocal())
                    rfcdate = date.strftime("%a, %d %b %Y %H:%M:%S %z")
                    header1 += "Date: " + rfcdate + "\n"
                    # Wed, 29 May 2019 03:59:29 +0000 形式じゃない... こんな形式: 2012/10/6  7:13:30
                if not pd.isnull(a_line.SenderEmailAddress) :
                    header1 += "From: " + a_line.SenderEmailAddress + "\n"
                if not pd.isnull(a_line.DisplayTo) :
                    # DisplayTo は Outlook仕様で ; 区切り。emlは , 区切り。
                    # メールアドレスがない場合も多い。Exchangeスペシャルで、アドレスを変換することも考える
                    addrs = ",".join([convert_header(i) for i in a_line.DisplayTo.split(";")])
                    header1 += "To: " + addrs + "\n"
                if not pd.isnull(a_line.DisplayCc) :
                    # DisplayTo は Outlook仕様で ; 区切り。emlは , 区切り。
                    # メールアドレスがない場合も多い。Exchangeスペシャルで、アドレスを変換することも考える
                    addrs = ",".join([convert_header(i) for i in a_line.DisplayCc.split(";")])
                    header1 += "Cc: " + addrs + "\n"
                header = header1

            print( no_CtrlM(header), file=fout, end='' )

            # BOUNDARY=CSV2EML-001a1149d018653acc05655c3e96
            # Add extra header of Content-Type: multipart/mixed; boundary="$BOUNDARY"
            print( 'Content-Type: multipart/mixed; boundary="%s"' % (BOUNDARY ,), file=fout)

            # Then  output empty line.
            print( '', file=fout)

            # Now output body of a message. It is multipart/mixed; boundary="$BOUNDARY".
            body_exist = False

            # if there is a plain text, which is a value in a column 0x10000, 
            # then output boundary for beginning of a part
            if ( not pd.isnull(a_line[col_plain_text])) :
                print( '--' + BOUNDARY, file=fout)									# --$BOUNDARY
                print( 'Content-Type: text/plain; charset="utf-8"', file=fout)		# header in a part
                print( '', file=fout) 												# empty line
                print( no_CtrlM(a_line[col_plain_text]), file=fout, end='') 		# content
                body_exist = True

            # if there is html body,
            if ( not pd.isnull(a_line[col_html_base64])) :
                print( '--' + BOUNDARY, file=fout)									# --$BOUNDARY
                print( no_CtrlM(a_line[col_html_base64]), file=fout)		# content(header of the part and empty line is included in the content)
                body_exist = True

            # if there is RTF body,
            if ( not pd.isnull(a_line[col_rtf])) :
                print( '--' + BOUNDARY, file=fout)									# --$BOUNDARY
                print( no_CtrlM(a_line[col_rtf]), file=fout)		# content(header of the part and empty line is included in the content)
                body_exist = True

            # if there is EntryId ,
            if ( not pd.isnull(a_line[col_summary]) and not body_exist) :
                print( '--' + BOUNDARY, file=fout)									# --$BOUNDARY
                print( 'Content-Type: text/plain; charset="utf-8"', file=fout)		# header in a part
                print( '', file=fout) 												# empty line
                print( no_ctrlM(a_line[col_summary]), file=fout, end='') 			# content

            # Finally, close multipart
            # --$BOUNDARY--
            print( '--' + BOUNDARY + '--' , file=fout)								# --$BOUNDARY--

if __name__ == '__main__':
    # just for debug. 
    indir = pathlib.Path('c:/Users/ueda/Desktop/email-recover/recovered_20190530/')
    targets = indir.glob( '01-*.csv' )

    for f in targets:
        csv2eml(f, indir)
