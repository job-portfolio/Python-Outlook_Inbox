import win32com.client, re


# Function: find structure-  findThis 'any character'
def Finder_String(string, findThis, flags=0):
    pattern = re.compile(findThis)
    found = re.findall(pattern, string)
    return found


# Get Experian Risk Tracker emails and extract URLs and Company Names into a list
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

list_of_url_cn = []
temp = []

# Extract Web Address and Company Name from Message Body of an Experian email
for message in messages:
    foundURL = Finder_String(message.Subject, 'New Experian Business Express Alerts')
    print(foundURL)