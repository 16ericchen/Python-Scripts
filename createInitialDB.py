import sqlite3
from sqlite3 import Error
import os
from os import path

os.chdir(path.dirname(__file__))

def create_connection(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except Error as e:
        print(e)
    return conn

def create_table(conn, create_table_sql):
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
    except Error as e:
        print(e)

def createTable():
    database = r"C:\Users\EricChen\Desktop\Python Scripts\In Progress\web_Scraper\yageoDB.db"
    conn = create_connection(database)
    sql_create_projects_table = """ CREATE TABLE IF NOT EXISTS projects (
                                            APN text PRIMARY KEY,
                                            Project_Name text
                                        ); """
    create_table(conn, sql_create_projects_table)
                
def main():
    createTable()

if __name__ == '__main__':
    main()

