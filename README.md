# outlook-comtool
Microsoft Outlook win32com script collection.

## Installation
```
pip install https://github.com/andiwand/outlook-comtool/archive/master.zip
```

## outlook-dumpcontacts
```
usage: outlook-dumpcontacts [-h] [--attributes ATTRIBUTES] [-e EXTRA] -i INPUT
                            -o OUTPUT

Filter script for exported Microsoft Outlook contacts.

optional arguments:
  -h, --help            show this help message and exit
  --attributes ATTRIBUTES
                        filter attributes (comma separated)
  -e EXTRA, --extra EXTRA
                        parse body for extra information (comma separated)
  -i INPUT, --input INPUT
                        input file
  -o OUTPUT, --output OUTPUT
                        output file
```

## outlook-filtercontacts
```
usage: outlook-filtercontacts [-h] [--attributes ATTRIBUTES] [-m MODE] [-o OUTPUT]
                              [-a ACCOUNT]

Export script for Microsoft Outlook contacts.

optional arguments:
  -h, --help            show this help message and exit
  --attributes ATTRIBUTES
                        filter attributes (comma separated)
  -m MODE, --mode MODE  opperation mode (dump, dump_photos, list_attr,
                        list_acc)
  -o OUTPUT, --output OUTPUT
                        output file/directory
  -a ACCOUNT, --account ACCOUNT
                        email address of the account
```