import sqlite3

from PyQt5 import QtGui, QtWidgets, QtCore, QtSql
import uic

with sqlite3.connect("uavto\utrbase.db") as db:
    cur = db.cursor()
    db.commit()


def export(self):
    cur.execute('''SELECT fam, ima, Otch, Ud, klass FROM vod''')
    self.dialog_3.tableWidget.setRowCount(0)
    for row, form in enumerate(cur):
        self.dialog_3.tableWidget.insertRow(row)
        for column, item in enumerate(form):
            self.dialog_3.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))

