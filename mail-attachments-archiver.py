#!/usr/bin/python2

# libraries import
import email, email.header, getpass, imaplib, os, time, re, argparse

parser = argparse.ArgumentParser(description='IMAP mail attachments archiver.')
parser.add_argument('USER', help='IMAP username')
parser.add_argument('PWD', help='IMAP password')
parser.add_argument('IMAPSERVER', help='IMAP server hostname')
parser.add_argument('--dump_dir', help='Folder to save attachments (default: ./archive)', default='./archive', required=False)

args = parser.parse_args()

# --- --- --- --- ---
# CONFIGURATION BEGIN
# --- --- --- --- ---

# only consider unread emails?
FILTER_UNREAD_EMAILS = False

# mark emails as read after their attachments have been archived?
MARK_AS_READ = False

# delete emails after their attachments have been archived?
DELETE_EMAIL = False

# if no attachment is found, mark email as read?
MARK_AS_READ_NOATTACHMENTS = False

# if no attachment is found, delete email?
DELETE_EMAIL_NOATTACHMENTS = False

# if no match is found, mark email as read?
MARK_AS_READ_NOMATCH = False

# if no match is found, delete email?
DELETE_EMAIL_NOMATCH = False

# you could filter using the IMAP rules here (check http://www.example-code.com/csharp/imap-search-critera.asp)
#searchstring = 'ALL'
searchstring = '(SEEN SINCE 01-Mar-2017)'

# --- --- --- --- ---
#  CONFIGURATION END
# --- --- --- --- ---

# source: https://stackoverflow.com/questions/12903893/python-imap-utf-8q-in-subject-string
def decode_mime_words(s): return u''.join(word.decode(encoding or 'utf8') if isinstance(word, bytes) else word for word, encoding in email.header.decode_header(s))

# connecting to the IMAP serer
m = imaplib.IMAP4_SSL(args.IMAPSERVER)
m.login(args.USER, args.PWD)
# use m.list() to get all the mailboxes
m.select("INBOX") # here you a can choose a mail box like INBOX instead

if FILTER_UNREAD_EMAILS: searchstring = 'UNSEEN'
resp, items = m.search(None, searchstring)
items = items[0].split() # getting the mails id
for emailid in items:
        # fetching the mail, "(RFC822)" means "get the whole stuff", but you can ask for headers only, etc
        resp, data = m.fetch(emailid, "(RFC822)")
        # getting the mail content
        email_body = data[0][1]
        # parsing the mail content to get a mail object
        mail = email.message_from_string(email_body)
        # check if any attachments at all
        if mail.get_content_maintype() != 'multipart':
                # marking as read and delete, if necessary
                if MARK_AS_READ_NOATTACHMENTS: m.store(emailid.replace(' ',','),'+FLAGS','\Seen')
                if DELETE_EMAIL_NOATTACHMENTS: m.store(emailid.replace(' ',','),'+FLAGS','\\Deleted')
                continue
        # checking sender
        sender = mail['from'].split()[-1]
        senderaddress = re.sub(r'[<>]','', sender)
        print "<"+str(mail['date'])+"> "+"["+str(mail['from'])+"] :"+str(mail['subject'])
        # check subject
        subject = mail['subject']
        outputdir = args.--dump_dir
        # we use walk to create a generator so we can iterate on the parts and forget about the recursive headach
        for part in mail.walk():
                # multipart are just containers, so we skip them
                if part.get_content_maintype() == 'multipart':
                        # marking as read and delete, if necessary
                        if MARK_AS_READ: m.store(emailid.replace(' ',','),'+FLAGS','\Seen')
                        if DELETE_EMAIL: m.store(emailid.replace(' ',','),'+FLAGS','\\Deleted')
                        continue
                # is this part an attachment?
                if part.get('Content-Disposition') is None:
                        # marking as read and delete, if necessary
                        if MARK_AS_READ: m.store(emailid.replace(' ',','),'+FLAGS','\Seen')
                        if DELETE_EMAIL: m.store(emailid.replace(' ',','),'+FLAGS','\\Deleted')
                        continue
                filename = part.get_filename()
                counter = 1
                # if there is no filename, we create one with a counter to avoid duplicates
                if not filename:
                        filename = 'part-%03d%s' % (counter, 'bin')
                        counter += 1
                # getting mail date
                filename = decode_mime_words(u''+filename)
                att_path = os.path.join(outputdir, filename)
                # check if output directory exists
                if not os.path.isdir(outputdir): os.makedirs(outputdir)
                # check if its already there
                if not os.path.isfile(att_path):
                        try:
                                print 'Saving to', str(att_path)
                                # finally write the stuff
                                fp = open(att_path, 'wb')
                                fp.write(part.get_payload(decode=True))
                                fp.close()
                                # marking as read and delete, if necessary
                                if MARK_AS_READ: m.store(emailid.replace(' ',','),'+FLAGS','\Seen')
                                if DELETE_EMAIL: m.store(emailid.replace(' ',','),'+FLAGS','\\Deleted')
                        except: pass
# Expunge the items marked as deleted... (Otherwise it will never be actually deleted)
if DELETE_EMAIL: m.expunge()
# logout
m.logout()
