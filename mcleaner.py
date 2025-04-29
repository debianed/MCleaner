#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# This script connects to an email server using IMAP, processes specified mail folders,
# and deletes emails older than a defined cutoff time. It uses logging to track actions
# and errors, and reads configuration details from a JSON file (mcleaner.json).

import email
import imaplib
import json
import logging
from os import path

def process_mailbox_folder(mail, user, folder, cut_time): 
    """
    Selects the specified folder and searches for emails older than the cutoff time.
    
    Parameters:
    mail: The IMAP connection object.
    user: The email user.
    folder: The folder to process.
    cut_time: The cutoff time for deleting emails.
    """
    # Select the specified folder and search for emails
    mail.select(folder, readonly=False)
    logging.info(f'{user}: {folder}: Search for messages older {cut_time}h')
    status, messages = mail.search(None, f'(OLDER {cut_time * 60 * 60})')

    # Process the emails found in the search
    if status == 'OK':
        messages = messages[0].split()
        for i, num in enumerate(messages):
            msg = None
            if (i == 0) or (i == len(messages) - 1):
                typ, msg_data = mail.fetch(num, '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        part = response_part[1].decode('utf-8')
                        msg = email.message_from_string(part)
                logging.info(f'{user}: {folder}: Add "Delete" mark for message #{num}, message date: {msg["Date"]}')
            else:
                logging.info(f'{user}: {folder}: Add "Delete" mark for message #{num}')
            mail.store(num, '+FLAGS', '\\Deleted')            
    else:
        logging.error(f"{user}: {folder}: Failed to search emails: {status}")

    # Expunge the mailbox
    logging.info(f'{user}: {folder}: Delete all marked messages')
    mail.expunge()    

def process_mailbox(mbox_config, cutoff_config): 
    """
    Connects to the mailbox and processes the specified folders.
    
    Parameters:
    mbox_config: Configuration for the mailbox connection.
    cutoff_config: Configuration for cutoff times for each folder.
    """
    try:
        # Connect to the mailbox
        mail = imaplib.IMAP4_SSL(mbox_config["imap"])
        user = mbox_config["username"]
        mail.login(user, mbox_config["password"])

        # Process folders
        for fld in cutoff_config:
            process_mailbox_folder(mail, user, fld, cutoff_config[fld])

        # Close the mailbox
        mail.close()
        mail.logout()

    except imaplib.IMAP4_SSL.error as e:
        logging.error(f"{user}: IMAP error: {e}")
    except Exception as e:
        logging.error(f"{user}: Error processing mailbox: {e}")

def main(): 
    """
    Main function to initialize logging, read the configuration file, and process mailboxes.
    """
    logging.basicConfig(filename='mcleaner.log',
                        filemode='w',
                        format='%(asctime)s,%(msecs)03d %(name)s %(levelname)s %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.DEBUG)
    
    logging.info("Start ifcm.com.ru mail cleaner")

    config_file = path.join(path.dirname(path.abspath(__file__)), 'mcleaner.json')

    try:
        with open(config_file, 'r', encoding='utf-8') as file:
            config = json.load(file)
    except FileNotFoundError as e:
        raise FileNotFoundError(f"Configuration file '{config_file}' not found.") from e
    except json.JSONDecodeError as e:
        raise json.JSONDecodeError(f"Invalid JSON in configuration file '{config_file}'.") from e

    for mbox in config["mboxes"]:
        process_mailbox(config["mboxes"][mbox], config["cutofftime"])

if __name__ == "__main__":
    main()
