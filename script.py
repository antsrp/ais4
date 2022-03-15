import psycopg2
from config import config

from layout import *


def connect():
    conn = None
    try:
        params = config()

        conn = psycopg2.connect(**params)

    except (Exception, psycopg2.DatabaseError) as error:
        print(error)

    return conn


def end(cur, conn):
    cur.close()
    if conn is not None:
        conn.close()
        print('Database connection closed.')


def main():

    conn = connect()
    cur = conn.cursor()

    form = Form(cur)
    form.loop()

    end(cur, conn)


if __name__ == '__main__':
    main()
