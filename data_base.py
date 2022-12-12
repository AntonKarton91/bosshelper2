import sqlite3

import employerList

initial_database = '''CREATE TABLE IF NOT EXISTS employers (name TEXT, surname TEXT, stage TEXT)'''


def use_database(command=initial_database):
    with sqlite3.connect('employers.db') as con:
        cur = con.cursor()
        cur.execute(command)


for e in employerList.emloyers_list:
    with sqlite3.connect('employers.db') as con:
        cur = con.cursor()
        cur.execute('''INSERT INTO employers VALUES(?,?,?)''', (e["name"], e["surname"], e["stage"]))