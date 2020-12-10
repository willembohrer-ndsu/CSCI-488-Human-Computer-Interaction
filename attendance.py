import csv
import ssl
import time
import pygame
import smtplib
import asyncio
import datetime
import schedule
import datetime
import psycopg2
import threading
import xlsxwriter
import psycopg2.extras
import RPi.GPIO as GPIO
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from mfrc522 import SimpleMFRC522
from flask import Flask, render_template

# Declaration of Constant variables.
ROOM_ID = 1
CONTINUE_SCANNING = True

app = Flask(__name__)

def getDBConnection():
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")
    return db_conn

@app.route("/")
def home():
    db_conn = getDBConnection()
    buildinglist = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    buildinglist.execute("""
                        SELECT DISTINCT
                            R.ID AS {},
                            R.BUILDING AS {}
                        FROM
                            ROOM R
                        WHERE
                            R.ACTIVE = TRUE;
                     """.format('\"Room_ID\"','\"Name\"'))

    return render_template("main.html", buildinglist=buildinglist)

@app.route("/setRoom", methods = ["POST"])
def setRoom():
    return 'Room Test'

def exportExcel(class_name, class_section):
    now = datetime.datetime.now()
    db_conn = getDBConnection()
    COUNTER = 0
    COL_COUNTER = 0
    ROW_COUNTER = 1
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    dict_cur.execute("""
                    SELECT
                        S.ID AS {},
                        S.NAME AS {},
                        S.EMAILADDRESS AS {},
                        C.SECTION AS {}
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
                            C.NAME = '{}' AND
                            C.SECTION = {} AND
                            A.CREATEDON > CURRENT_DATE - INTERVAL '1' DAY;
                    """.format('\"Student ID:\"','\"Name:\"','\"Email Address:\"','\"Section:\"', ROOM_ID, class_name, class_section))

    # Create an Excel Workbook and add a Worksheet.
    workbook = xlsxwriter.Workbook('Attendance_{}{}{}.xlsx'.format(now.strftime("%b"), now.day, now.year))
    worksheet = workbook.add_worksheet('Section_{}'.format(class_section))

    # Dynamically write all records in the dictionary cursor using their column name and the value
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

def emailReport():
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
        message = MIMEMultipart()
        message['From'] = sender
        message['To'] = "{}".format(record[2])

        message['Subject'] = "Class Attendance {} - Section {}".format(record[0], record[1])

        body = "Attached is the attendance for {} - section {}, on {} {}, {}".format(record[0], record[1], now.strftime("%b"), now.day, now.year)

        # Add body to email
        message.attach(MIMEText(body, "plain"))

        exportExcel(record[0], record[1])

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

def insertClass(email, name, number, section, starttime, endtime, room, days):
    # SQL call to insert a Class into the database utilizing data from the Web Interface.
    db_conn = getDBConnection()
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
    # Execute database procedure for inserting class records.
    dict_cur.execute("""
                         CALL CREATE_CLASS({}, {}, {}, {}, {}, {}, {}, {}); COMMIT;
                     """.format(email, name, number, section, starttime, endtime, room, days))
    return

@app.before_first_request
def activate_scanning():
    def run_scanning():
        try:
            # Create dictionary cursor in order for pulling data into.
            GPIO.setwarnings(False)
            reader = SimpleMFRC522()
            db_conn = getDBConnection()
            dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
            # Initialize sound to be played on a scan.
            pygame.init()
            pygame.mixer.music.set_volume(1.0)
            pygame.mixer.music.load("ding.mp3")
            # Schedule the daily export of attendance.
            #schedule.every().day.at("23:59").do(emailReport)
            schedule.every().minute.at(":15").do(emailReport)
            while CONTINUE_SCANNING:
                # Run any scheduled exports
                schedule.run_pending()
                # Scan card's data and remove any trailing spaces from the string.
                id, scanned_data = reader.read()
                scanned_data = scanned_data.strip()
                # Execute database procedure for inserting attendance records.
                dict_cur.execute("""
                                     CALL LOG_ATTENDANCE({}, {}); COMMIT;
                                 """.format(ROOM_ID, scanned_data))
                # Play sound after a successful scan
                pygame.mixer.music.play()
                time.sleep(2)
        except(Exception, psycopg2.Error) as error:
            print(error)
        finally:
            #closing database connection.
            if(db_conn):
                dict_cur.close()
                db_conn.close()
                GPIO.cleanup()
                print("Connection has been closed")

    thread = threading.Thread(target=run_scanning)
    thread.start()

app.run(debug = False)
