import win32com.client
import util

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
            yield folder.Items.Item(i)

def contact_attributes(contact):
    result = {}
    for key in contact._prop_map_get_:
        try:
            value = getattr(contact, key)
        except Exception, e:
            print e
        if not util.check_type(value): continue
        result[key] = util.fix_type(value)
    return result
