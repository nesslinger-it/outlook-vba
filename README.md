# outlook-vba
Outlook VBA-Extensions

## AddDateToSubject
Simple Outlook VBA Script which adds the mail received date in the subject line of a MailItem.
Adds the date to all selected mails. If a slected item isn't a MailItem (e.g. MeetingItem / Calender-Entry) then a counter increase an displays a Message Box to the user.

### Example Integration in the Office Ribbon
![exampleIntegration](https://user-images.githubusercontent.com/9899606/147570384-b35c2e78-552e-49c9-8146-6f9c2d361140.JPG)

### Example output
![exampleOutput](https://user-images.githubusercontent.com/9899606/147731727-dad71a0f-1389-4c85-bede-1c593a5b6878.JPG)

## How to use VBA without lower the security settings
You can sign your VBA-Script with a digital certificate.
Office provides a easy way to create a user based certificate with the "SELFCERT.EXE" tool in the Office root directory e. g. C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE

After the creation of the certificate you can sign your code directly in Outlook "Tools --> Digital Signature..." 

You can find more under the following link:
https://support.microsoft.com/de-de/office/digitales-signieren-eines-makroprojekts-956e9cc8-bbf6-4365-8bfa-98505ecd1c01
