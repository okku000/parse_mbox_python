#! /usr/bin/env python3
# File name: parse_mbox.py
# Description: Parse .mbox file. Decode options are base64 and quoted-printable for charsets of UTF-8 and ISO-2022-JP(JIS)
# Usage: python3 parse_mbox.py -i mails.mbox -o test.xlsx
# Version: Python 3.8.10
# Author: Yuya Okumura
# Date: 02-12-2021

import mailbox
from typing import Dict, List, Literal, Any
import bs4
import base64
import re
import quopri
import xlsxwriter
import argparse
import sys
import logging
from logging import critical, error, info, warning, debug
import traceback


def parse_arguments():
    parser = argparse.ArgumentParser(
        description='Arguments get parsed via --commands')
    parser.add_argument('-v', metavar='verbosity', type=int, default=3,
                        help='Verbosity of logging: 0 -critical, 1- error, 2 -warning, 3 -info, 4 -debug')
    parser.add_argument('-i', metavar='input', type=str, required=True,
                        help='Provide .mbox file to parse')
    parser.add_argument('-o', metavar='output', type=str, required=True,
                        help="Provide a .xlsx file for output")

    args = parser.parse_args()
    verbose = {0: logging.CRITICAL, 1: logging.ERROR,
               2: logging.WARNING, 3: logging.INFO, 4: logging.DEBUG}
    logging.basicConfig(format='%(message)s',
                        level=verbose[args.v], stream=sys.stdout)

    return args


def get_html_text(html):
    try:
        return bs4.BeautifulSoup(html, 'lxml').body.get_text(' ', strip=True)
    except AttributeError:  # message contents empty
        return None


class Decoder(object):
    def __init__(self, msg: str, charset: Literal['utf-8', 'iso-2022-jp', 'iso-8859-2', 'us-ascii'], transfer_method: Literal['base64', 'quoted-printable']):
        self.msg = msg
        self.charset = charset
        self.transfer_method = transfer_method
        self.bynary = self._encode_charset()

    def call(self):
        decoded_transfer = self._decode_with_transfer()
        return self._decode_charset(decoded_transfer)

    def _encode_charset(self):
        return self.msg.encode(self.charset)

    def _decode_charset(self, msg: bytes):
        return msg.decode(self.charset)

    def _decode_with_transfer(self):
        if self.transfer_method == 'base64':
            return self._decode_with_base64()
        elif self.transfer_method == 'quoted-printable':
            return self._decode_with_quoted_printable()

    def _decode_with_base64(self):
        return base64.b64decode(self.bynary)

    def _decode_with_quoted_printable(self):
        return quopri.decodestring(self.bynary)


class EmailDecoder(Decoder):
    def __init__(self, msg: str, charset: Literal['utf-8', 'iso-2022-jp', 'iso-8859-2', 'us-ascii'], transfer_method: Literal['base64', 'quoted-printable']):
        super().__init__(msg, charset, transfer_method)
        # chain emails can be divided in to parts by spliting with 'From:' notation
        chained_email = r'..From:'

        self.regex = re.compile(chained_email)

    def fetch_first_email(self):
        decoded_text = super().call()
        text_wo_re = self.regex.split(decoded_text)
        return text_wo_re[0].replace('\r', '').replace('\n', '')


class SubjectDecoder(Decoder):
    def __init__(self, subject: str):
        self.regex = self._set_regex()
        self.subject = subject

    def _set_regex(self):
        # .mbox contains subject encoded in the format
        # where first argument explains charset and encode method (=?UTF-8?Q?)
        # e.g. subject: =?UTF-8?Q?Marqu=C3=A9s_de_Vargas_=26_Baud?=
        # Q for quoted-printable B for base64
        regex = r"""
        \=\?
        (utf\-8|iso\-2022\-jp)    # charset
        \?                        
        .                         # encode-transfer-method 
        \?"""
        return re.compile(regex, re.VERBOSE)

    def call(self):
        if not self._is_required_to_decode():
            return self.subject

        return self._decode_subject()

    def _decode_subject(self):
        lines = self.subject.split("\n")
        decoded_subject = []
        for line in lines:
            # regex takes at least 10 chars, so < 10 is not a target
            if len(line) < 10:
                continue
            encoding = self._identify_decode_method(line)
            decoded_subject.append(self._decode_line(encoding))
        return ''.join(decoded_subject)

    def _decode_line(self, obj: Dict[str, str]):
        super().__init__(obj['text'], obj['charset'], obj['transfer_method'])
        return super().call()

    def _is_required_to_decode(self):
        return self.regex.search(self.subject) is not None

    def _identify_decode_method(self, line: str):
        line = line.strip()
        search = self.regex.search(line)
        charset = search.group(1)
        decoded_way_char = search.group(0)[-2]

        transfer_method = ''
        if decoded_way_char == 'B':
            transfer_method = 'base64'
        else:
            transfer_method = 'quoted-printable'

        decodable_text = self.regex.split(line)[-1]
        return {"transfer_method": str(transfer_method), "charset": str(charset), "text": str(decodable_text)}


class GmailMboxMessage():
    def __init__(self, email_data):
        if not isinstance(email_data, mailbox.mboxMessage):
            raise TypeError('Variable must be type mailbox.mboxMessage')
        self.email_data = email_data

    def parse_email(self):
        email_date = self.email_data['Date']
        email_from = self.email_data['From']
        email_to = self.email_data['To']
        email_subject = SubjectDecoder(str(self.email_data['Subject'])).call()
        content = self._read_email_payload()[0]
        print("\nDate: ", email_date, "\nFrom: ", email_from, "\nTo: ",
              email_to, "\nSubject: ", email_subject, "\nContent: ", content)

        content_type = content[0]
        charset = content[1]
        transfer = content[2]
        text = content[3]
        return {'Date': email_date, 'From': email_from, 'To': email_to, 'Subject': email_subject, 'Content_type': content_type, 'Charset': charset, 'Transfer': transfer, 'Text': text}

    def _read_email_payload(self):
        email_payload = self.email_data.get_payload()
        if self.email_data.is_multipart():
            email_messages = list(self._get_email_messages(email_payload))
        else:
            email_messages = [email_payload]
        return [self._read_email_text(msg) for msg in email_messages]

    def _get_email_messages(self, email_payload):
        for msg in email_payload:
            if isinstance(msg, (list, tuple)):
                for submsg in self._get_email_messages(msg):
                    yield submsg
            elif msg.is_multipart():
                for submsg in self._get_email_messages(msg.get_payload()):
                    yield submsg
            else:
                yield msg

    def _fetch_content_type(self, msg):
        return 'NA' if isinstance(msg, str) else msg.get_content_type()

    def _fetch_charset(self, msg):
        charset = 'NA' if isinstance(
            msg, str) else msg.get('Content-Type', 'NA')
        if "utf-8" in charset.lower():
            charset = "utf-8"
        if "iso-2022-jp" in charset.lower():
            charset = "iso-2022-jp"
        if "iso-8859-2" in charset.lower():
            charset = "iso-8859-2"
        if "us-ascii" in charset.lower():
            charset = "us-ascii"
        return charset

    def _fetch_encoding_method(self, msg):
        return 'NA' if isinstance(msg, str) else msg.get(
            'Content-Transfer-Encoding', 'NA')

    def _create_readable_text(self, msg, content_type, encoding, charset):
        if self._is_readable_text(content_type, encoding, charset):
            return EmailDecoder(
                msg.get_payload(), charset, encoding).fetch_first_email()
        else:
            return 'NA'

    def _read_email_text(self, msg):
        content_type = self._fetch_content_type(msg)
        encoding = self._fetch_encoding_method(msg)
        charset = self._fetch_charset(msg)
        msg_text = self._create_readable_text(
            msg, content_type, encoding, charset)

        return (content_type, charset,  encoding, msg_text)

    def _is_readable_text(self, content_type, encoded_way, charset):
        return True if 'text/plain' in content_type.lower() and encoded_way.lower() in ['base64', 'quoted-printable'] and charset in ['utf-8', 'iso-2022-jp', 'iso-8859-2', 'us-ascii'] else False


class ExcelSheet():
    def __init__(self, worksheet):
        self.worksheet = worksheet
        self.content = {}
        self.row = 0

    def call(self, content):

        self._set_content(content)
        if self.row == 0:
            self._write_title()
            self.row += 1
        self._write_content()

    def close(self):
        self.workbook.close()

    def _set_content(self, content):
        self.content = content

    def _fetch_titles(self):
        l = []
        for title in self.content.keys():
            l.append(title)
        return l

    def _write_title(self):
        column = 0
        titles = self._fetch_titles()
        for title in titles:
            self.worksheet.write(self.row, column, title)
            column += 1

    def _write_content(self):
        column = 0
        for item in self.content.values():
            self.worksheet.write(self.row, column, item)
            column += 1
        self.row += 1


def main(args):
    try:
        mbox_obj = mailbox.mbox(args.i)
        workbook = xlsxwriter.Workbook(args.o)
        worksheet = workbook.add_worksheet()
        spreadsheet = ExcelSheet(worksheet)
        num_entries = len(mbox_obj)

        i = 0
        for idx, email_obj in enumerate(mbox_obj):
            email_data = GmailMboxMessage(email_obj)
            content = email_data.parse_email()
            spreadsheet.call(content)
            info(
                '=-=-=-=-=-=-=-=-=-=-Parsing email {0} of {1}-=-=-=-=-=-=-=-=-=-=-=-=-=-=-'.format(idx+1, num_entries))
            # i += 1
            # if i == 10:
            #     break
    except Exception as e:
        error(str(e))
        traceback.print_exc()
    finally:
        workbook.close()


if __name__ == '__main__':
    args = parse_arguments()
    main(args)
