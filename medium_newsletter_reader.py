# Medium newsletter reader

# Author: Oriol Ordi

# This program accesses the e-mail and searches for e-mails in the inbox from the sender noreply@medium.com
# then it takes one of e-mails from that sender and accesses its contents.
# Within the contents, it searches for the first link (the main one (and subject) in the email)
# and opens it in an maximized window incognito version of google chrome.
# Finally, it deletes the e-mail from google chrome's inbox (and sends it to the trash folder)

# IMPORTANT:
# This code only works for GMAIL
# The variables email_username and email_password (hopefully around line 34 of the code) have to be changed
# to the desired email and password
# The variable chrome path (hopefully around line 101 of the code) has to be changed to the installation location of Chrome
# In order for this code to work, 2 things have to be changed in Gmail preferences:
# the first is going to Gmail and then Settings --> Settings --> Forwarding and POP/IMAP --> IMAP access --> Enable IMAP
# the second is going to Account --> Manage your google account --> Security --> Less secure app access --> ON


# Import necessary libraries
import imaplib
import email
from collections import namedtuple
import re
import os
import webbrowser


# Get the text of the body of the oldest medium email
def get_medium_oldest_email_body_text():
    # Import email username and password
    try:
        email_username = os.environ.get('EMAIL_ADDRESS')  # get the email address from the EMAIL_ADDRESS environmental variable
        email_password = os.environ.get('EMAIL_PASSWORD')  # get the email password from the EMAIL_PASSWORD environmental variable
    except:
        print('Could not get email address and password from environmental variables')
        print('Press enter to exit')
        input()
        return False, '', '', ''

    # Instantiate IMAP4_SSL object from imap module and login to the email
    try:
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(email_username, email_password)
    except imaplib.IMAP4.error:
        print('Could not log into the e-mail')
        print('Press enter to exit')
        input()
        return False, '', '', ''

    # Get a byte list of the number of e-mails on the selected folder (in this case the inbox folder)
    mail.select('Inbox')  # selects the inbox folder
    type_, [data] = mail.search(None, 'FROM "noreply@medium.com"')  # read only e-mails received from noreply@medium.com
    id_list = data.split()

    # If the id_list is empty, there are no e-mails from noreply@medium.com, thus there is nothing to show. End the function
    if len(id_list) == 0:
        print('No email to read')
        print('Press enter to exit')
        input()
        return False, '', '', ''

    # Fetch the first email on the list (which, to my understanding thus far, seems to be a random one)
    id_email = id_list[0]#.decode('utf-8')
    type_, data = mail.fetch(id_email, '(RFC822)')

    # Iterate the responses part of the e-mails
    for response_part in data:
        if isinstance(response_part, tuple):
            # Get the e-mail
            msg = email.message_from_string(response_part[1].decode('utf-8'))
            # Iterate over the parts of the e-mail body
            for i, part in enumerate(msg.walk()):
                # If the part of the multipart e-mail body is text, get the body of the part with the get_payload method
                # and add that text body part to the list containing all possible body parts of each e-mail
                if part.get_content_type() == 'text/plain':
                    email_text = part.get_payload(decode=True).decode('utf-8')

    # End the function and return true because the email existed. Also returns:
    # the body [email_text] (the text) of the email to process and find links to open
    # the mail object [mail] (to later use it to store the email in the trash)
    # the e-mail id [id_email] in bytes type to later know which email has been opened and, thus, has to be stored in the trash
    return True, email_text, mail, id_email


# Open the links from the body text of the email
def open_links_from_email_body_text(email_text):
    # Find all urls on the email_text
    try:
        #links = re.findall(r'https://medium.+\)', email_text)  # finds all links
        first_link = re.search(r'(https://medium.+)\)', email_text)  # finds the first link
        first_link = first_link.group(1)  # the link is in the first group of the re.search (https://medium.+)
        # remember that .group(0) is the whole regular expression, while group(1) is the first group (first thing in parenthesis)

        # Assign the url of interest to the variable
        url = first_link

        try:
            # Open chrome with the specified url
            chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s --incognito --start-maximized'
            chrome = webbrowser.get(chrome_path)
            try:
                chrome.open_new(url)
            except:
                print("Couldn't open the specified link")
                print('Press enter to exit')
                input()
        except:
            print("Couldn't find Google chrome")
            print('Press enter to exit')
            input()
    except:
        print('No links found in the email')
        print('Press enter to exit')
        input()


# Move the email to the medium folder after it's been read
def move_email_to_medium_folder(mail, id_email):
    # Convert an int to bytes (need to test if it works as well for string, otherwise convert strint to int first with int(the_string)--> bytes([the_int_number])
    #mail.store('1', '+X-GM-LABELS', 'Medium')  # add the first medium e-mail on the list to the "Medium" label in Gmail
    mail.store(id_email, '+X-GM-LABELS', '\\Trash')  # remove the first medium e-mail on the list (send it to the trash)
    mail.expunge()  # make the store take effect


# Main function
def main():
    email_existing, email_text, mail, id_email = get_medium_oldest_email_body_text()

    if email_existing:
        open_links_from_email_body_text(email_text)
        move_email_to_medium_folder(mail, id_email)

    try:
        mail.logout()
    except:
        print('Not possible to logout')


# Execute main function
if __name__ == '__main__':
    main()
