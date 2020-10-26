import psycopg2
try:
    # This connection information is for the User created within the database using:
    #sudo -u postgres createuser --interactive --pwprompt
    connection = psycopg2.connect(user = "applicationuser", password = "CoronaSux2020!", host = "localhost", port = "5432", database = "postgres")

    query_cursor = connection.cursor()

    # Print PostgreSQL Connection properties
    print(connection.get_dsn_parameters(),"\n")

    # Print PostgreSQL version
    query_cursor.execute("SELECT version();")
    record = query_cursor.fetchone()
    print("Successful connection to: ", record,"\n")

except(Exception, psycopg2.Error) as error:
    print("Error while connecting", error)

finally:
    #closing database connection.
    if(connection):
        query_cursor.close()
        connection.close()
        print("Connection has been closed")
