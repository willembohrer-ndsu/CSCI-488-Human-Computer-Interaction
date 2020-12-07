import psycopg2
import psycopg2.extras
import RPi.GPIO as GPIO
from mfrc522 import SimpleMFRC522

try:
    reader = SimpleMFRC522()

    # This connection information is for the User created within the database using:
    # sudo -u postgres createuser --interactive --pwprompt
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    # Scan card's data and remove any trailing spaces from the string. In this example, I am using a Professor's name as the card's data.
    id, scanned_data = reader.read()
    scanned_data = scanned_data.strip()

    # Create dictionary cursor in order for pulling data into.
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                    SELECT
                        P.NAME AS {},
                        R.NAME AS {},
                        A.NAME AS {},
                        A.DESCRIPTION AS {}
                    FROM
                        PROFESSOR P
                        INNER JOIN ROLE R ON
                            R.ID = P.ROLE_ID
                        INNER JOIN ACCESS A ON
                            A.ID = P.ACCESS_ID
                    WHERE
                        P.NAME = '{}';
                    """.format('\"Professor\"','\"Role\"','\"Access\"','\"Access Description\"', scanned_data))

    # Dynamically print all records in the dictionary cursor using their column name and the value
    column_names = [desc[0] for desc in dict_cur.description]
    for record in dict_cur:
        COUNTER = 0
        for column in record:
            print('{}: {}'.format(column_names[COUNTER], column))
            COUNTER += 1
        print('\n')

except(Exception, psycopg2.Error) as error:
    print("Error while connecting", error)

finally:
    # closing database connection.
    if(db_conn):
        dict_cur.close()
        db_conn.close()
        print("Connection has been closed")
    GPIO.cleanup()
