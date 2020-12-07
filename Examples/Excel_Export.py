import xlsxwriter
import psycopg2
import psycopg2.extras
import datetime

try:
    now = datetime.datetime.now()

    # This connection information is for the User created within the database using:
    # sudo -u postgres createuser --interactive --pwprompt
    db_conn = psycopg2.connect(user = "ApplicationUser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    # Create dictionary cursor in order for pulling data into.
    dict_cur = db_conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

    # Execute an SQL statement and store the results in the dictionary cursor
    dict_cur.execute("""
                    SELECT
                        S.ID AS {},
                        S.NAME AS {},
                        S.EMAILADDRESS AS {}
                    FROM
                        STUDENT S;
                    """.format('\"ID\"','\"Name\"','\"Email Address\"'))

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('Attendance_{}{}{}.xlsx'.format(now.strftime("%b"), now.day, now.year))
    # TODO write logic for dynamically gathering the class and section for the exported attendance spreadsheet
    worksheet = workbook.add_worksheet('Class_Section_TODO')
    # Dynamically print all records in the dictionary cursor using their column name and the value
    column_names = [desc[0] for desc in dict_cur.description]
    # Global variables
    COUNTER = 0
    COL_COUNTER = 0
    ROW_COUNTER = 1
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

except(Exception, psycopg2.Error) as error:
    print("Error while connecting", error)

finally:
    # Closing database connection.
    if(db_conn):
        dict_cur.close()
        db_conn.close()
        print("Connection has been closed")
