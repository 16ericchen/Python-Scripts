import sqlite3
from sqlite3 import Error

database = r"N:\Python\Yageo\yageoDB.db"

def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)
    return conn

def create_projects(conn, projects):
    sql = ''' INSERT or IGNORE INTO projects(APN,Project_Name)
              VALUES(?,?) '''
    cur = conn.cursor()
    cur.execute(sql, projects)
    conn.commit()
    return cur.lastrowid

def selectData(apn,conn):
    cur = conn.cursor()
    cur.execute("SELECT * FROM projects WHERE APN=?", (apn,))
    rows = cur.fetchall()
    if rows == []:
        return False
    return rows[0][1]

def populateData(apn,modified,conn):
    task_1 = apn,modified
    create_projects(conn,task_1)


def searchData(APN):
    global database
    conn = create_connection(database)
    if conn is not None:
        return selectData(APN,conn)
    else:
        print("Error! cannot create the database connection.")

def enterData(APN,modified):
    global database
    conn = create_connection(database)
    if conn is not None:
        populateData(APN,modified,conn)
    else:
        print("Error! cannot create the database connection.")

