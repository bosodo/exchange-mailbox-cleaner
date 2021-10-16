# -*- coding: UTF-8 -*-
"""
Cleanup tool for Microsoft Exchange mailbox.
You can use it to delete emails from exchange mailboxes, that are not managed by Outlook, OWA or etc.
You can also export deleted emails to .eml files. Default option is hard-delete e-mails.
(script homepage: https://github.com/bosodo/exchange-mailbox-cleaner/)

Usage:
  exch-mbox-cleaner.py <exch-server> <user-name> <user-pass> (--inbox | --inbox-subdir=<directory>) [--days=<number_days>] [--bckp=DIR] [--dry-run] [(--soft | --trash)]
  exch-mbox-cleaner.py (-h | --help)
  exch-mbox-cleaner.py --version

Examples:
  exch-mbox-cleaner.py exchange.example.org jan_kowalski Password! --inbox --days=180 --bckp='./deleted-emails'
  exch-mbox-cleaner.py exchange.example.org jan_kowalski Password! --inbox-subdir='Sample subdirectory' --days=180 --bckp='./deleted-emails'

Options:
  -h --help                     Show this screen.
  --version                     Show version.
  --days=<number_days>          Emails more than x days old will be deleted [default: 30]
  --inbox-subdir=<directory>    Subdirectory in INBOX (or select --inbox as root)
  --bckp=DIR                    Location to backup deleted emails (directory must exist!).
  --soft                        Soft-delete (keep a copy in the recoverable items folder).
  --trash                       Move message to the trash folder.
  --dry-run                     Only "dry-run". Run script without deleting emails.
"""
from exchangelib import DELEGATE, Account, Credentials, Configuration, UTC_NOW
from datetime import timedelta
from docopt import docopt


def connect(username, password, exch_server):
    credentials = Credentials(
        username=username,
        password=password
    )

    config = Configuration(
        server=exch_server,
        credentials=credentials
    )

    return Account(
        primary_smtp_address=username,
        config=config,
        autodiscover=False,
        access_type=DELEGATE
    )


def get_size_of_mailbox_folder(mailbox_folder, add_text):
    size = sum(mailbox_folder.all().values_list('size', flat=True))
    print('- size of mailbox folder (%s): %.2f MB' % (add_text, (size / 1024 / 1024)))


def main():
    args = docopt(__doc__, version='exchange-mailbox-cleaner v0.1')
    mbox_account = connect(args['<user-name>'], args['<user-pass>'], args['<exch-server>'])

    if args['--inbox']:
        processing_folder = mbox_account.inbox
    else:
        processing_folder = mbox_account.inbox / args['--inbox-subdir']

    if args['--dry-run']:
        print('Initiating dry-run...')

    get_size_of_mailbox_folder(processing_folder, 'before')
    since = UTC_NOW() - timedelta(days=int(args['--days']))
    print('- searching for e-mails older than: ' + str(since))
    emails = processing_folder.all().filter(datetime_received__lt=since).order_by('-datetime_received')
    print('- number of email to delete: ' + str(emails.count()))

    if args['--dry-run'] is False:
        for item in emails:
            if args['--bckp'] is not None:
                decoded_item = item.mime_content.decode("utf-8")
                with open(args['--bckp'] + '/' +
                          str(item.datetime_created.strftime("%Y-%m-%d-%H%M%S")) + '_' +
                          ''.join(e for e in item.subject[:25] if e.isalnum()) +
                          '.eml', 'w') as f:
                    f.write(decoded_item)
            if args['--soft']:
                item.soft_delete()
            elif args['--trash']:
                item.move_to_trash()
            else:
                item.delete()
    else:
        print('Dry-run completed. No email was deleted.')

    get_size_of_mailbox_folder(processing_folder, 'after')


if __name__ == "__main__":
    main()
