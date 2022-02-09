# Exchange mailbox cleaner

Cleanup tool for Microsoft Exchange mailbox (https://github.com/bosodo/exchange-mailbox-cleaner/).

You can use it to delete emails from exchange mailboxes, that are not managed by Outlook, OWA or etc. (e.g. technical e-mail accounts).
Deleted emails can be exported to .eml files. (default option is hard-delete e-mails).
- `exch-mbox-cleaner.py` based on `exchangelib`, which use Microsoft Exchange Web Services (EWS) to communicate with MS 
  Exchange server. 

### Options
```
-h --help                     Show this screen.
--version                     Show version.
--days=<number_days>          Emails more than x days old will be deleted [default: 30]
--inbox-subdir=<directory>    Subdirectory in INBOX (or select --inbox as root)
--bckp=DIR                    Location to backup deleted emails (directory must exist!).
--soft                        Soft-delete (keep a copy in the recoverable items folder).
--trash                       Move message to the trash folder.
--dry-run                     Only "dry-run". Run script without deleting emails.
```

### Dependency

Python 3.6+. Before run, install the required dependencies with:
```
pip install -r requirements.txt
```

### Usage
```
$ exch-mbox-cleaner.py <exch-server> <user-name> <user-pass> (--inbox | --inbox-subdir=<directory>) [--days=<number_days>] [--bckp=DIR] [--dry-run] [(--soft | --trash)]
$ exch-mbox-cleaner.py (-h | --help)
$ exch-mbox-cleaner.py --version
```
- `<user-name>` - `jan_kowalski` or with domain `jan_kowalski@example.org` depend on your exchange configuration

### Examples
```
$ exch-mbox-cleaner.py exchange.example.org jan_kowalski@example.org Password! --inbox --days=180 --bckp='./deleted-emails'
$ exch-mbox-cleaner.py exchange.example.org jan_kowalski@example.org Password! --inbox-subdir='Sample subdirectory' --days=180 
--bckp='./deleted-emails'
```

