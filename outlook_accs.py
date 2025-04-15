import win32com.client

# Initialize Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def get_accounts(namespace):
    """Retrieve all Outlook accounts on the machine."""
    accounts = namespace.Accounts
    account_list = []

    for account in accounts:
        account_list.append(account.SmtpAddress)
    
    return account_list

print(get_accounts(namespace))