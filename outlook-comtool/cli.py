import time
import json
import base64
import datetime
from pprint import pprint
import pywintypes
import outlook

def fix_contacts_birthdays(contacts, startAfter=None):
    do = not startAfter
    tmpbd = pywintypes.Time(time.time())
    
    for n, contact in enumerate(contacts):
        if not do:
            do = contact.FullName == startAfter
            continue
        tmp = contact.Birthday
        if tmp.year == 4501: continue
        print n, contact.FullName.encode("utf8", errors="replace")
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

def dump_contacts(contacts):
    result = []
    for n, contact in enumerate(contacts):
        print n, contact.FullName.encode("utf8", errors="replace")
        single = outlook.contact_attributes(contact)
        for key in single:
            value = single[key]
            if type(value) == datetime.datetime:
                value = value.isoformat(" ")
            if type(value) == buffer:
                value = base64.b64encode(value)
            single[key] = value
        result.append(single)
    return result

def dump_contacts_photos(contacts):
    pass

if __name__ == "__main__":
    o = outlook.Outlook()
    contacts = o.get_contacts("office@logsol.at")
    
    #print count_contacts_birthdays(contacts)
    
    #fix_contacts_birthdays(contacts, startAfter="")
    
    #contact = find_contact("andreas stefl", contacts)
    #pprint(outlook.contact_attributes(contact))
    
    d = dump_contacts(contacts)
    json.dump(d, open("contacts.json", "w"), sort_keys=True, indent=4, separators=(",", ": "))
