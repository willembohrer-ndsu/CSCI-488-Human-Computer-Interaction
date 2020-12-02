import psycopg2
import psycopg2.extras
import smtplib
import ssl
import csv
import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    ROOM_ID = 1
    now = datetime.datetime.now()
    # This connection information is for the User created within the database using:
    #sudo -u postgres createuser --interactive --pwprompt
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    # Create dictionary cursor in order for pulling data into.
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                SELECT
                    C.NAME,
                    C.SECTION,
                    P.EMAILADDRESS,
                    P.LASTNAME
                FROM
                    PROFESSOR P
                    INNER JOIN ROOM_SCHEDULE RS ON
                        RS.ROOM_ID = {}
                    INNER JOIN CLASS C ON
                        C.ROOM_SCHEDULE_ID = RS.ID AND
                        C.PROFESSOR_ID = P.ID;
                """.format(ROOM_ID))

    # Establish email port and server
    port = 465
    smtp_server = "smtp.gmail.com"

    # Email and password that the emails will be generated from
    sender = "hci.488.2020@gmail.com"
    email_password = "Human_Computer_Int488"

    #messages = []
    # Creates email
    for record in dict_cur:
        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = "{}".format(record[2])

        message['Subject'] = "Class Attendance {} - Section {}".format(record[0], record[1])

        body = "Attached is the attendance for {} - section {}, on {} {}, {}".format(record[0], record[1], now.strftime("%b"), now.day, now.year)
        
        # Add body to email
        message.attach(MIMEText(body, "plain"))

        filename = "Attendance_{}{}{}.xlsx".format(now.strftime("%b"), now.day, now.year)

        # Open file in binary mode
        with open(filename, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part)

        # Add header as key/value pair to attachment part
        part.add_header(
           "Content-Disposition", "attachment", filename = filename
         )
        # Add attachment to message and convert message to string
        message.attach(part)
        text = message.as_string()
        print(sender, email_password)

        # Create CSV file of last name and email address of PROFESSORS
        with open("emails.csv", mode='w') as csv_file:
            # Creates a dictionary of last name and professors
            fieldnames = ["last_name", "email_address"]
            csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
            #csv_writer.writeheader()
            csv_writer.writerow({"last_name": "{}".format(record[3]), "email_address": "{}".format(record[2])})
        # Reads the CSV file and sends an email to the PROFESSORS
        with open("emails.csv", mode='r') as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                # Login to server and send email
                context = ssl.create_default_context()
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                    server.ehlo()
                    server.login(sender, email_password)
                    print(row["email_address"])
                    server.sendmail(sender, row["email_address"], text)
                    
except(Exception, psycopg2.Error) as error:
    print(error)

finally:
    #closing database connection.
    if(db_conn):
        dict_cur.close()
        db_conn.close()
        print("Connection has been closed")
