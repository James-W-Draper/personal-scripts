# Mail test
# Define the SMTP server and port
$smtpServer = "smtp.office365.com"
$smtpPort = 587

# Define the email details
$from = "lmcs_mail@enstargroup.com"
$to = "james.Draper@enstargroup.com"
$subject = "Test Email"
$body = "This is a test email sent via Office 365 SMTP server."

# Define the SMTP credentials
$password = "3n8TARlm"

# Create the SMTP client object
$smtp = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.EnableSsl = $true
$smtp.Credentials = New-Object System.Net.NetworkCredential($from, $password)

# Create the email message object
$message = New-Object Net.Mail.MailMessage($from, $to, $subject, $body)

# Show the time
Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Send the email
$smtp.Send($message)