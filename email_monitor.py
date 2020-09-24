import os
import json
import time
import poplib
import smtplib
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr, formataddr
from email.mime.text import MIMEText
from email.header import Header
import re
import logging

import config

# config log
LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
DATE_FORMAT = "%m/%d/%Y %H:%M:%S %p"
logging.basicConfig(filename='running.log', level=logging.INFO, format=LOG_FORMAT, datefmt=DATE_FORMAT)
#logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, datefmt=DATE_FORMAT)

# bacis config
EMAIL = config.EMAIL
PASSWD = config.PASSWD
POP3_SERVER = config.POP3_SERVER
POP3_PORT = config.POP3_PORT
SMTP_SERVER = config.SMTP_SERVER
SMTP_PORT = config.SMTP_PORT
SENDER = formataddr((Header(config.SENDER, 'utf-8').encode(), EMAIL))
BASIC_PATH = config.BASIC_PATH
JSON_PATH = config.JSON_PATH
PATTERN = re.compile(r'^(20S\d{6})-(.*)-A(\d+)..+')

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
        value = value.decode(charset)
    return value


def accpetEmail(to):
    msg = MIMEText("我们已经收到了你的作业:)", 'plain', 'utf-8')
    msg['From'] = SENDER
    msg['To'] = _format_addr(to)
    msg['Subject'] = Header("来自现代密码学助教的提醒", 'utf-8').encode()
    return msg

def rejectEmail(to):
    msg = MIMEText("无法自动收取你的附件，已拒绝接收:(", 'plain', 'utf-8')
    msg['From'] = SENDER
    msg['To'] = _format_addr(to)
    msg['Subject'] = Header("来自现代密码学助教的提醒", 'utf-8').encode()
    return msg

def getHeader(msg):
    headers = {}
    value = msg.get('Subject', '')
    subject = decode_str(value)
    headers['Subject'] = subject
    value = msg.get('From', '')
    hdr, addr = parseaddr(value)
    name = decode_str(hdr)
    headers['From'] = u'%s <%s>' % (name, addr)
    headers['FromAddr'] = addr
    return headers

def getContent(msg):
    for part in msg.walk():
        logging.debug(part.get_content_maintype())
        if part.get_content_maintype() == 'multipart':
                continue
        name = part.get_filename()
        if not name:
            continue
        filename = decode_str(name)
        logging.debug('getContent():'+filename)
        match_res = PATTERN.match(filename)
        if match_res:
            id, name, times = match_res.groups()
            path = os.path.join(BASIC_PATH, 'A'+times)
            if not os.path.exists(path):
                os.makedirs(path)
            path = os.path.join(path, filename)
            with open(path, 'wb') as fp:
                fp.write(part.get_payload(decode=True))
                logging.info('Save file [{}]'.format(path))
            return True
        else:
            return False
    return False

def monitorEmail():
    # load record from json
    with open(JSON_PATH, 'r') as f:
        db = json.load(f)
    record_size = db['record']
    sleep_time = db['sleep']
    while True:
        # connect to POP3 Server
        server = poplib.POP3_SSL(host=POP3_SERVER, port=POP3_PORT)
        logging.info(server.getwelcome().decode('utf-8'))

        server.user(EMAIL)
        server.pass_(PASSWD)

        logging.info("Message: %s. Size: %s" % server.stat())

        # check any new emails
        resp, mails, octets = server.list()
        logging.debug(resp)
        logging.debug(mails)
        logging.debug(octets)
        cur_size = len(mails)
        # sleep for some time if there are no new emails
        if cur_size <= record_size:
            server.quit()
            time.sleep(sleep_time)
            sleep_time *= 2
            sleep_time = min(sleep_time*2, 300)
            continue

        # reset sleep time
        sleep_time = db['sleep']
        # connect to SMTP server
        smtp_server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        smtp_server.login(EMAIL, PASSWD)

        email_parser = Parser()
        for i in range(record_size+1, cur_size+1):
            # get email data and parser
            resp, lines, octets = server.retr(i)
            msg_content = '\r\n'.join([x.decode('utf-8', 'ignore') for x in lines])
            msg = email_parser.parsestr(msg_content)
            # parser header
            headers = getHeader(msg)
            logging.debug(headers['From'])
            # check attachments
            res = getContent(msg)
            # reply emails
            if not res:
                logging.error('Error email [{}] from [{}]'.format(headers['Subject'], headers['From']))
                smtp_server.sendmail(EMAIL, [headers['FromAddr']], rejectEmail(headers['From']).as_string())
            else:
                smtp_server.sendmail(EMAIL, [headers['FromAddr']], accpetEmail(headers['From']).as_string())
            logging.info("handle No.{} email complete.".format(i))

        # save current handle emails
        record_size = cur_size
        db['record'] = cur_size
        with open(JSON_PATH, 'w') as f:
            json.dump(db, f)
        server.quit()
        smtp_server.quit()

if __name__ == "__main__":
    while True:
        try:
            monitorEmail()
        except Exception() as e:
            time.sleep(60*10)
