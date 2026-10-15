import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import imaplib
import email
import re
from imapclient import IMAPClient
import sys

# Fill in your MPI email login username, password, and first name.
ID = ""
PASSWORD = ""
NAME = ""

def search_inbox_title(*subject_terms):
    # establish a connection
    mail = imaplib.IMAP4_SSL("mail.mpi-inf.mpg.de")
    # authenticate
    mail.login(ID, PASSWORD)
    # select the inbox to search
    mail.select("inbox")
    # search for mail whose subject contains every term; subjects look like
    # [Account: Request Finished] Account Created for "Last, First".
    criteria = " ".join('HEADER Subject "%s"' % term for term in subject_terms)
    result, data = mail.uid('search', None, '(%s)' % criteria)
    # if there is no specific email, inform the user
    if not data[0]:
        print("There are no emails with a subject containing", " + ".join(subject_terms))
        return
    # list of email IDs
    email_ids = data[0].split()
    # get the latest matching email
    latest_email_id = email_ids[-1]
    # fetch the email body (RFC822) for the given ID
    result, email_data = mail.uid('fetch', latest_email_id, '(BODY.PEEK[])')
    raw_email = email_data[0][1].decode()
    email_message = email.message_from_string(raw_email)
    # return the plain-text body of the latest email
    for part in email_message.walk():
        if part.get_content_type() == "text/plain":
            body = part.get_payload(decode=True)
            try:
                return body.decode('utf-8')
            except UnicodeDecodeError:
                return body.decode('iso-8859-1')
        else:
            continue

def send_email(to_address, subject, message, from_address= ID + "@mpi-inf.mpg.de", password=PASSWORD):
    msg = MIMEMultipart()
    msg["From"] = from_address
    msg["To"] = to_address
    msg["Subject"] = subject

    msg.attach(MIMEText(message, "plain"))
    
    ports = [25, 465, 587, 1025, 1465]
    for port in ports:
        try:
            # 465/1465 expect TLS from the first byte; a plain SMTP() there hangs waiting for a greeting
            if port == 465 or port == 1465:
                server = smtplib.SMTP_SSL("mail.mpi-inf.mpg.de", port, timeout=30)
            else:
                server = smtplib.SMTP("mail.mpi-inf.mpg.de", port, timeout=30)
                server.starttls()
            server.login(ID, password)
            text = msg.as_string()
            server.sendmail(from_address, to_address, text)
        except Exception as e:
            print(f"Failed to send email through port {port} due to {str(e)}")
            continue
        server.close()
        print(f"Email successfully sent to {to_address}")

        # the email is already out, so a failure here must not fall through to the next port and send it again
        try:
            with IMAPClient("mail.mpi-inf.mpg.de", port=993) as client:
                client.login(ID, password)
                client.select_folder("Sent")
                client.append("Sent", msg.as_bytes())
            print("Email saved in 'Sent' folder")
        except Exception as e:
            print(f"Email was sent, but saving it to the 'Sent' folder failed due to {str(e)}")
        return
    print(f"Could not send the email after trying all ports")
if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 send_email.py [name] [receiver]")
        sys.exit(1)

    name = sys.argv[1]
    receiver = sys.argv[2]

    # Example usage
    title = "Your MPI-INF Account Credentials and Resources"
    body = search_inbox_title('Account Created for', name)
    if body is None:
        sys.exit(f"Could not read the 'Account Created' email for {name}; nothing was sent.")
    # the email body has CRLF line endings, so strip the trailing \r the regex captures
    username = re.search(r"Login\.*: (.*)\n", body).group(1).strip()
    password = re.search(r"Initial \(one-time\) password\.*: (.*)\n", body).group(1).strip()
    messsage = f"""Dear {name},

Welcome to MPI-INF! Your account credentials have been created and listed below.

User name: {username}
Start password: {password}

Upon receipt of this email, please immediately visit password.mpi-inf.mpg.de to set your new password.

Afterward, you can access your account and details at domino.mpi-inf.mpg.de. Our email service is accessible at mail.mpi-inf.mpg.de.

For detailed instructions on using our services, such as connecting to our WiFi network, accessing printing services, or setting up a VPN, please visit our extensive documentation hub at https://plex.mpi-klsb.mpg.de/display/Documentation/IST+Documentation+Home. Simply use the search bar at the top right corner to find the information you're looking for. This resource is designed to address most questions you may have.

Please feel free to contact me if you have any questions. For hardware-related issues (getting a laptop or other equipment), please contact Zihang Wu (zihangwu@mpi-inf.mpg.de).

Best,
{NAME}

    """

    print(title)
    print(messsage)
    send_email(receiver, title, messsage)
