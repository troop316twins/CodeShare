import win32com.client

n = 1
while n <= 3:
    const = win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "I AM THE SUBJECT OF MSG %d !!" % (n)
    newMail.Body = "I AM IN THE BODY\nSO AM I!!!"
    newMail.To = "Joel.Monza@phillyshipyard.com"
    # attachment1 = r"E:\test\logo.png"

    # newMail.Attachments.Add(Source=attachment1)
    # newMail.display()
    newMail.send
    # time.sleep(1000)
    n += 1
