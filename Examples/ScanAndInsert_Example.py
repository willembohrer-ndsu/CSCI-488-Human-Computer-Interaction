import psycopg2
import psycopg2.extras
import RPi.GPIO as GPIO
from mfrc522 import SimpleMFRC522

try:
    # Constant value for which room the Pi is located in.
    ROOM_ID = 1;

    reader = SimpleMFRC522()

    # This connection information is for the User created within the database using:
    # sudo -u postgres createuser --interactive --pwprompt
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    # Scan card's data and remove any trailing spaces from the string. In this example, I am using a Professor's name as the card's data.
    id, scanned_data = reader.read_no_block()
    scanned_data = scanned_data.strip()

    # Create dictionary cursor in order for pulling data into.
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                    CALL LOG_ATTENDANCE({}, {});
                    """.format(ROOM_ID, scanned_data))


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
