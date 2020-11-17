import psycopg2
import psycopg2.extras
import smtplib
import csv

try: 

    # This connection information is for the User created within the database using:
    #sudo -u postgres createuser --interactive --pwprompt
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    # Create dictionary cursor in order for pulling data into.
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                    SELECT
                        P.FIRSTNAME AS {},
                        P.LASTNAME AS {},
                        P.EMAILADDRESS AS {}
                    FROM
                        PROFESSOR P;
                    """.format('\"First Name\"','\"Last Name\"','\"Email Address\"'))

    # Establish email port and server
    port = 465
    smtp_server = "smtp.gmail.com"

    # Email and password that the emails will be generated from
    sender = "hci.488.2020@gmail.com"
    email_password = "Human_Computer_Int488"

    # Email message
    message = """\
    Subject: Class Attendance
    To: {recipient}
    From: {sender}

    Greetings Professor {LASTNAME}, this is your email test.
    These students attending class on (insert date here)
    (insert list of students here)
    """
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

except(Exception, psycopg2.Error) as error:
    print("Error while connecting", error)

finally:
    #closing database connection.
    if(db_conn):
        dict_cur.close()
        db_conn.close()
        print("Connection has been closed")
