# SMTP-Windows-Service
A windows service which reads an XML file after every 15 minutes from a fixed location and sends email to the receipts.

XML INPUT FORMAT:
<EmailMessage> <To></To> <Subject></Subject> <MessageBody></MessageBody> </EmailMessage>
