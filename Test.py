import win32com.client

const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "I AM SUBJECT!!"
newMail.Body = "I AM IN THE BODY\nSO AM I!!!"
newMail.To = "Joel.Monza@phillyshipyard.com"
#attachment1 = r"E:\test\logo.png"

#newMail.Attachments.Add(Source=attachment1)
newMail.display()
newMail.send()