#!/usr/bin/env python2

import os
import sys
import time
import json
import base64
import datetime
import argparse
import pywintypes
import outlookcomtool.outlook as outlook

class GeneratorList(list):
    def __init__(self, gen):
        self.__gen = gen
    def __iter__(self):
        return self.__gen
    def __len__(self):
        return 1

def fix_contacts_birthdays(contacts, startAfter=None):
    do = not startAfter
    tmpbd = pywintypes.Time(time.time())

    for n, contact in enumerate(contacts):
        if not do:
            do = contact.FullName == startAfter
            continue
        tmp = contact.Birthday
        if tmp.year == 4501: continue
        print(n, contact.FullName.encode("utf8", errors="replace"))
        contact.Birthday = tmpbd
        contact.Save()
        contact.Birthday = tmp
        contact.Save()

def count_contacts_birthdays(contacts):
    result = 0
    for contact in contacts:
        if contact.Birthday.year == 4501: continue
        result += 1
    return result

def find_contact(name, contacts):
    name = name.lower()
    for contact in contacts:
        return contact
        if name in contact.FullName.lower():
            return contact
    pass

def dump_contacts(contacts, attributes=None):
    for n, contact in enumerate(contacts):
        print(n, contact.FullName.encode("utf8", errors="replace"))
        single = outlook.contact_attributes(contact, attributes)
        for key in single:
            value = single[key]
            if type(value) == datetime.datetime:
                value = value.isoformat(" ")
            if type(value) == buffer:
                value = base64.b64encode(value)
            single[key] = value
        yield single

def dump_contacts_photos(contacts, path):
    fs_enc = sys.getfilesystemencoding()
    for i, contact in enumerate(contacts):
        if not contact.HasPicture: continue
        name = contact.FullName.encode("utf8", errors="replace")
        print(i, name)
        file_name = contact.FullName.replace("\"", "")
        photo_path = os.path.join(path, file_name) + ".jpg"
        photo_path = photo_path.encode(fs_enc, errors="replace")
        for j in range(1, contact.Attachments.Count + 1):
            attachement = contact.Attachments.Item(j)
            if attachement.FileName != "ContactPicture.jpg": continue
            if os.path.exists(photo_path):
                print("error: file already exists: %s" % photo_path)
            attachement.SaveAsFile(photo_path)

def main():
    parser = argparse.ArgumentParser(description="Export script for Microsoft Outlook contacts.")
    parser.add_argument("--attributes", help="filter attributes (comma separated)", default="FullName,FirstName,LastName,Title,Suffix,HomeAddressCountry,BusinessAddressCountry,Categories,Email1Address,Email2Address,Email3Address,Body")
    parser.add_argument("-m", "--mode", help="opperation mode (dump, dump_photos, list_attr, list_acc)", default="dump")
    parser.add_argument("-o", "--output", help="output file/directory")
    parser.add_argument("-a", "--account", help="email address of the account")
    args = parser.parse_args()

    o = outlook.Outlook()
    accounts = o.get_accounts()
    if args.mode == "list_acc":
        for a in accounts:
			print(a)
        return 0
    if args.account not in accounts:
        print("account not found")
        return 2

    contacts = o.get_contacts(args.account)
    attributes = [a for a in args.attributes.split(",") if a]
    if not attributes: attributes = None

    out = sys.stdout
    close = False
    if args.output:
        out = open(args.output, "w")
        close = True

    if args.mode == "list_attr":
        contact = next(contacts)
        for attribute in outlook.contact_list_attributes(contact):
            print(attribute)
    elif args.mode == "dump":
        d = dump_contacts(contacts, attributes)
        g = GeneratorList(d)
        json.dump(g, out, sort_keys=True, indent=4, separators=(",", ": "))
    elif args.mode == "dump_photos":
        dump_contacts_photos(contacts, args.output)
    else:
        print("unknown option")
        return 1

    out.flush()
    if close: out.close()
    return 0

if __name__ == "__main__":
    code = main()
    sys.exit(code)
