import win32com.client
#comment
n = 1
while n <= 1:
    const = win32com.client.constants
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = "Awesome Python Things :D"
    # newMail.Subject = "I AM THE SUBJECT OF MSG %d !!" % (n)
    newMail.Body = "I have figured out how to send basic e-mails from python. :D \n" \
                   "you do realize that I'm being super nice and not making a loop to send you e-mails >:D"
    newMail.To = "Christopher.Spurdle@phillyshipyard.com"
    attachment1 = r"C:\Users\gphl-JM5\Desktop\Test_Picture.jpg"

    newMail.Attachments.Add(Source=attachment1)
    # newMail.display()
    newMail.send
    # time.sleep(1000)
    n += 1
