from win32com.client import Dispatch
import win32com.client
import pandas as pd

class Python_outlook:
    def __init__(self,send_to=None):
        self.send_to = send_to
        # self.connection = self.__enter__()
        self.outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        self.all_contacts_df = self.contacts()
        


    # def __enter__(self):
    #     self.connection = win32com.client.gencache.EnsureDispatch("Outlook.Application")

    #     return self.connection


    def contacts(self):
        # Outlook stuff
        # outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        outGAL = self.outApp.Session.GetGlobalAddressList()
        entries = outGAL.AddressEntries

        # Empty list to store contact info
        contact_list = list()

        # Iterates through your contact book and extracts/appends them to a list
        for entry in entries:
            if entry.Type == "EX":
                user = entry.GetExchangeUser()
                if user is not None:
                    if len(user.FirstName) > 0 and len(user.LastName) > 0:
                        row = list()
                        row.append(user.FirstName)
                        row.append(user.LastName)
                        row.append(user.PrimarySmtpAddress)
                        """print("First Name: " + user.FirstName)
                        print("Last Name: " + user.LastName)
                        print("Email: " + user.PrimarySmtpAddress)"""
                        contact_list.append(row)
            #Create dataframe from contact list
            contact_df = pd.DataFrame(contact_list,columns=['Name','Last Name','Email'])
            contact_df['Full Name'] = contact_df['Name'] +' '+ contact_df['Last Name'] 
        return contact_df
    outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
# cnn = Python_outlook().connection

# print(cnn.all_contacts_df)
dff = Python_outlook()
print( dff.all_contacts_df)
