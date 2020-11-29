import csv
import time
import smtplib
import asyncio
import schedule
import datetime
import psycopg2
import xlsxwriter
import psycopg2.extras
import RPi.GPIO as GPIO
from mfrc522 import SimpleMFRC522

reader = SimpleMFRC522()

# Constant values.
ROOM_ID = 1
COUNTER = 0
COL_COUNTER = 0
ROW_COUNTER = 1

async def exportExcel():
    now = datetime.datetime.now()
    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                    SELECT
                        S.ID AS {},
                        S.NAME AS {},
                        S.EMAILADDRESS AS {}
                    FROM
                        STUDENT S
                        INNER JOIN ATTENDANCE A ON
                            A.STUDENT_ID = S.ID
                        INNER JOIN ROOM_SCHEDULE RS ON
                            RS.ID = A.ROOM_SCHEDULE_ID
                        INNER JOIN CLASS C ON
                            C.ROOM_SCHEDULE_ID = RS.ID;
                    """.format('\"Student ID\"','\"Name\"','\"Email Address\"', ROOM_ID))

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('Attendance_{}{}{}.xlsx'.format(now.strftime("%b"), now.day, now.year))
    # TODO write logic for dynamically gathering the class and section for the exported attendance spreadsheet
    worksheet = workbook.add_worksheet('Class_Section_TODO')
    # Dynamically print all records in the dictionary cursor using their column name and the value
    column_names = [desc[0] for desc in dict_cur.description]
    # Parse through the dictionary cursor and write each cell of data
    for record in dict_cur:
        for column in record:
            if(COUNTER == 0):
                while COL_COUNTER < len(column_names):
                    worksheet.write(0, COL_COUNTER, column_names[COL_COUNTER])
                    COL_COUNTER += 1
                COUNTER += 1
        COL_COUNTER = 0
        for row in record:
            worksheet.write(ROW_COUNTER, COL_COUNTER, row)
            COL_COUNTER += 1
        ROW_COUNTER += 1
    workbook.close()

async def scanCard():
    # Scan card's data and remove any trailing spaces from the string.
    id, scanned_data = reader.read_no_block()
    scanned_data = scanned_data.strip()
    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                    CALL LOG_ATTENDANCE({}, {});
                    """.format(ROOM_ID, scanned_data))

async def emailReport():
    # Establish email port and server
    port = 465
    smtp_server = "smtp.gmail.com"

    # Email and password that the emails will be generated from
    sender = "hci.488.2020@gmail.com"
    email_password = "Human_Computer_Int488"

    # Email message
    message = """\
    Subject: Class Attendance
    To: {}
    From: {}

    Greetings Professor {}, this is your email test.
    Attached is the attendance for your class, {}, on {} {}, {}
    (insert list of students here)
    """.format('willem.bohrer@ndsu.edu', sender, 'Bohrer', 'Insert class here',now.strftime("%b"), now.day, now.year)
    # Create and read a CSV file with list of professors
    # Login to server with email and password
    with smtplib.SMTP("smtp.gmail.com", 465) as server:
        server.login(sender, email_password)
    # Create CSV file of first name, last name, and email address of PROFESSORS
    with open("emails.csv") as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow([i[0] for i in dict_cur.description])
        csv_writer.writerows(dict_cur)
    with open("emails") as csv_file:
        reader = csv.reader(csv_file)
        next(reader)
        for LASTNAME, EMAILADDRESS in reader:
            server.sendmail(
                sender,
                EMAILADDRESS,
                message.format(name = LASTNAME, recipient = EMAILADDRESS, sender = sender)
            )
            print("sent to {name}")

try:
    # This connection information is for the User created within the database using:
    # sudo -u postgres createuser --interactive --pwprompt
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    # Create dictionary cursor in order for pulling data into.
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

except(Exception, psycopg2.Error) as error:
    print("Error while connecting", error)

finally:
    # closing database connection.
    if(db_conn):
        dict_cur.close()
        db_conn.close()
        print("Connection has been closed.")
    print("Cleaning up.")
    GPIO.cleanup()
