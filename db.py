import sqlite3

class Database:

    def __init__(self):
        self.conn = sqlite3.connect("eml.db")
        self.cur = self.conn.cursor()
        self.cur.execute("CREATE TABLE IF NOT EXISTS email (id INTEGER PRIMARY KEY, mailbox TEXT, received INTEGER,"
                         " sent INTEGER, resolved INTEGER, year INTEGER, month INTEGER, day INTEGER, week INTEGER)")
        self.conn.commit()

    def insert(self, mailbox, received, sent, resolved, year, month, day, week):
        self.cur.execute("INSERT INTO email VALUES (NULL, ?, ?, ?, ?, ?, ?, ?, ?)", (mailbox, received, sent, resolved,
                                                                                     year, month, day, week))
        self.conn.commit()



    def search(self, mailbox):
        self.cur.execute("SELECT * FROM email WHERE mailbox=?", mailbox)

    def search_by_date(self, mailbox, year, month, day):
        self.cur.execute("SELECT * FROM email WHERE mailbox=? AND year=? AND month=? AND day=?", (mailbox, year, month,
                                                                                                  day))

    def view_all(self):
        self.cur.execute("SELECT * FROM email")
        rows = self.cur.fetchall()

        return rows
