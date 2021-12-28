# outlook-vba
Outlook VBA-Extensions

## AddDateToSubject
Simple Outlook VBA Script which adds the mail received date in the subject.
Adds the date to all selected mails. If a slected item isn't a MailItem (e.g. MeetingItem / Calender-Entry) then a counter increase an displays a Message Box to the user.


## How to use VBA in a company network
You can sign your VBA-Script with a digital certificate.
Office provides a easy way to create a user based certificate with the "SELFCERT.EXE" tool in the Office root directory e. g. C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE

After the creation of the certificate you can sign your code directly in Outlook "Tools --> Digital Signature..." 

You can find more under the following link:
https://support.microsoft.com/de-de/office/digitales-signieren-eines-makroprojekts-956e9cc8-bbf6-4365-8bfa-98505ecd1c01
