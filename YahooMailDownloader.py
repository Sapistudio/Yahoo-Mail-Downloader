""" NOTE: temporarily turn on ‘Allow apps that use less secure sign-in’ via https://login.yahoo.com/account/security#other-apps for this script to work.
If mail is partially downloaded, script resumes from where last run left off. 
"""

import os, sys, getpass, imaplib, quopri


# app variables
IMAP_SERVER = 'imap.mail.yahoo.com'                                                             # change to access a different IMAP-supporting mail service like Gmail
MAILBOX_BLACKLIST = ['trash']
LOCAL_MAIL_FOLDER_NAME = 'Yahoo Mail'



def get_credentials():
    print('Yahoo Mail username: ')
    username = input().strip()
    psswd = getpass.getpass(prompt='Yahoo Mail password: ')
    return username, psswd


def login(username, psswd):
    imap_client = imaplib.IMAP4_SSL(IMAP_SERVER)
    try:
        login_status, login_summary = imap_client.login(username, psswd)
        if login_status == "OK":
            return imap_client
    except imaplib.IMAP4.error as e:
        print(str(e))
    return None


def download_mail(imap_client):
    # get list of all mailboxes
    list_status, mailboxes = imap_client.list()
    if list_status != "OK":
        print("Error: could not retrieve mailboxes")
        return 1
    # create local folder to save mail in
    if not os.path.isdir(LOCAL_MAIL_FOLDER_NAME):
        os.mkdir(LOCAL_MAIL_FOLDER_NAME)
    # download mail from each mailbox (excluding Trash)
    for mailbox in mailboxes:
        mailbox_name = get_mailbox_name(mailbox)
        if mailbox_name.lower() not in MAILBOX_BLACKLIST:
            try:
                open_status, num_emails = imap_client.select('"%s"' % mailbox_name, readonly=True)
                if open_status == "OK":
                    num_emails = num_emails_as_int(num_emails)
                    mailbox_folder_path = create_mailbox_folder(mailbox_name, num_emails)
                    search_status, message_set = imap_client.search(None, 'ALL')
                    for num in message_set[0].split():
                        download_message(imap_client, num, mailbox_folder_path)
                    imap_client.close()                                                         # close this mailbox
                else:
                    print("Error: could not open mailbox %s" % str(mailbox_name))
                    return 1
            except imaplib.IMAP4.error as e:
                print("%s   %s" % (mailbox_name, str(e)))
    return 0

def get_mailbox_name(mailbox):
    mailbox = mailbox.decode('utf-8')
    mailbox_name = mailbox.split(' "')[-1:][0].replace('"', '')
    return mailbox_name

def num_emails_as_int(num_emails):
    return int(num_emails[0].decode('utf-8'))

def create_mailbox_folder(mailbox_name, num_emails):
    mailbox_folder_name = "%s (%s)" % (str(mailbox_name), str(num_emails))
    mailbox_folder_path = LOCAL_MAIL_FOLDER_NAME + '/' + mailbox_folder_name
    if not os.path.isdir(mailbox_folder_path):
        os.mkdir(mailbox_folder_path)
    return mailbox_folder_path                                           

def download_message(imap_client, message_num, folder_path):
    message_num = message_num.decode('utf-8')
    file_path = folder_path + '/' + message_num + '.html'
    if not os.path.isfile(file_path):
        fetch_status, message_data = imap_client.fetch(message_num, '(RFC822)')
        if type(message_data[0][1]) is bytes:
            message_body = quopri.decodestring(message_data[0][1])
            with open(file_path, 'wb') as f:
                f.write(message_body)



if __name__ == '__main__':

    username, psswd = get_credentials()

    imap_client = login(username, psswd)
    if not imap_client:
        sys.exit(0) 
    download_status = download_mail(imap_client)

    imap_client.logout()













