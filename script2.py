import sys
import os
import imaplib
import getpass
import email
import email.header
import datetime

EMAIL_ACCOUNT = "placeholder@placeholder" #mail adress
EMAIL_PASSWORD = "placeholder"
EMAIL_FOLDER = "folder/subfolder" 
svdir = 'c:/directory/subdirectory'

def process_mailbox(M):
    reload(sys)
    sys.setdefaultencoding('utf-8')
    rv, data = M.search(None,'UNSEEN', '(Subject "placeholder")') #set filters for Subject and read/unread
    if rv != 'OK':
        print "No messages found!"
        return

    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print "ERROR getting message", num
            return

        msg = email.message_from_string(data[0][1])
		
        if msg.get_content_maintype() != 'multipart':
            continue

        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename=part.get_filename('png')
            if filename is not None:
                print filename			
                sv_path = os.path.join(svdir, filename)
                if not os.path.isfile(sv_path):
                    print sv_path       
                    fp = open(sv_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
				
        decode = email.header.decode_header(msg['Subject'])[0]
        subject = unicode(decode[0])
        print 'Message %s: %s' % (num, subject)
        print 'Raw Date:', msg['Date']
        # Now convert to local date-time
        date_tuple = email.utils.parsedate_tz(msg['Date'])
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(
                email.utils.mktime_tz(date_tuple))
            print "Local Date:", \
                local_date.strftime("%a, %d %b %Y %H:%M:%S")


M = imaplib.IMAP4_SSL('outlook.office365.com')

try:
    rv, data = M.login(EMAIL_ACCOUNT, EMAIL_PASSWORD)
except imaplib.IMAP4.error:
    print "LOGIN FAILED!!! "
    sys.exit(1)

print rv, data

rv, mailboxes = M.list()
if rv == 'OK':
    print "Mailboxes:"
    print mailboxes

rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print "Processing mailbox...\n"
    process_mailbox(M)
    M.close()
else:
    print "ERROR: Unable to open mailbox ", rv

M.logout()
