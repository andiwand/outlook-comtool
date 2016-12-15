import win32com.client
import outlookcomtool.util as util

OlInspectorClose_olDiscard = 1

class Outlook:
    def __init__(self):
        self._client = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        self._mapi = self._client.GetNamespace("MAPI") # https://msdn.microsoft.com/en-us/library/ff869848.aspx
    def get_accounts(self):
        accounts = [self._mapi.Accounts[i] for i in range(1, self._mapi.Accounts.Count + 1)]
        return {account.DisplayName: account for account in accounts}
    def get_contact_folder(self, store=None):
        if store is None: return self._mapi.GetDefaultFolder(win32com.client.constants.olFolderContacts)
        if type(store) is str: store = self._mapi.Stores[store]
        return store.GetDefaultFolder(win32com.client.constants.olFolderContacts)
    def get_contacts(self, store=None):
        folder = self.get_contact_folder(store)
        for i in range(1, folder.Items.Count + 1):
            contact = folder.Items.Item(i)
            if contact.Class != win32com.client.constants.olContact: continue
            yield contact
            contact.Close(win32com.client.constants.olDiscard)

def contact_list_attributes(contact):
    return [key for key in contact._prop_map_get_]

def contact_attributes(contact, attributes=None):
    result = {}
    for key in contact._prop_map_get_:
        if attributes is not None and key not in attributes: continue
        try:
            value = getattr(contact, key)
        except Exception as e:
            print(e)
        if not util.check_type(value): continue
        result[key] = util.fix_type(value)
    return result
