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

# Declaration of Constant variables.
ROOM_ID = 1
COUNTER = 0
COL_COUNTER = 0
ROW_COUNTER = 1
reader = SimpleMFRC522()

async def getDBConnection():
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")
    return db_conn

async def exportExcel():
    now = datetime.datetime.now()
    db_conn = getDBConnection()

    section_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    section_cur.execute("""
                        SELECT
                            C.SECTION
                        FROM
                            STUDENT S
                            INNER JOIN ATTENDANCE A ON
                                A.STUDENT_ID = S.ID
                            INNER JOIN ROOM_SCHEDULE RS ON
                                RS.ID = A.ROOM_SCHEDULE_ID
                            INNER JOIN CLASS C ON
                                C.ROOM_SCHEDULE_ID = RS.ID
                        WHERE
                            RS.ROOM_ID = {} AND
                            A.CREATEDON > CURRENT_DATE - INTERVAL '1' DAY;
                        """.format(ROOM_ID))
    for record in section_cur:
        dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
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
                                C.ROOM_SCHEDULE_ID = RS.ID
                            WHERE
                                RS.ROOM_ID = {} AND
                                A.CREATEDON > CURRENT_DATE - INTERVAL '1' DAY;
                        """.format('\"Student ID\"','\"Name\"','\"Email Address\"', ROOM_ID))

        for record in dict_cur:
            column_names = [desc[0] for desc in dict_cur.description]
            COUNTER = 0
            for column in record:
                print('{}: {}'.format(column_names[COUNTER], column))
                COUNTER += 1
            print('\n')

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook('Attendance_{}{}{}.xlsx'.format(now.strftime("%b"), now.day, now.year))
        # TODO: write logic for dynamically gathering the class and section for the exported attendance spreadsheet
        worksheet = workbook.add_worksheet('Class_Section_TODO')
        # Dynamically print all records in the dictionary cursor using their column name and the value
        column_names = [desc[0] for desc in dict_cur.description]
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

async def emailReport():
    now = datetime.datetime.now()
    db_conn = getDBConnection()

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
                        C.PROFESSOR_ID = P.ID
                """.format(ROOM_ID))

    # Establish email port and server
    port = 465
    smtp_server = "smtp.gmail.com"

    # Email and password that the emails will be generated from
    sender = "hci.488.2020@gmail.com"
    email_password = "Human_Computer_Int488"

    # Creates email
    for record in dict_cur:
        column_names = [desc[0] for desc in dict_cur.description]
        COUNTER = 0
        for column in record:
            print('{}: {}'.format(column_names[COUNTER], column))
            COUNTER += 1
        print('\n')
        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = "{}".format(record[2])

        message['Subject'] = "Class Attendance {} - Section {}".format(record[0], record[1])

        body = "Attached is the attendance for {} - section {}, on {} {}, {}".format(record[0], record[1], now.strftime("%b"), now.day, now.year)

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        exportExcel()

        filename = 'Attendance_{}{}{}.xlsx'.format(now.strftime("%b"), now.day, now.year)

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

        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.ehlo()
            server.login(sender, email_password)
            server.sendmail(sender, record[2], text)

async def scanCard():
    # Scan card's data and remove any trailing spaces from the string.
    id, scanned_data = reader.read_no_block()
    scanned_data = scanned_data.strip()
    # Execute database procedure for inserting attendance records.
    dict_cur.execute("""
                    CALL LOG_ATTENDANCE({}, {});
                    """.format(ROOM_ID, scanned_data))

async def insertClass(email, starttime, endtime, room, days):
    # TODO: Create an SQL call to insert a class into the database utilizing data from the web interface
    return

async def emailToProfessor(email):
    # TODO: create code to call an email export for the specified professor to receive their attendance records
    return

while (True):
    await scanCard()

except(Exception, psycopg2.Error) as error:
    print("Error while connecting", error)

finally:
    # Closing database connection.
    if(db_conn):
        dict_cur.close()
        db_conn.close()
        print("Connection has been closed.")
    print("Cleaning up.")
    GPIO.cleanup()
