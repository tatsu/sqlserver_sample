#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pyodbc

# 環境に合わせて値を変更しておく。
DRIVER = 'ODBC Driver 17 for SQL Server'  # ODBCドライバを使う場合
# SQL Serverがインストールされた環境では、こちらで動くはず。
# driver = 'SQL Server'
SERVER = 'localhost'
DATABASE = 'TestDB'
USERNAME = 'SA'
PASSWORD = '<YourNewStrong@Passw0rd>'


class SqlServer:
    def __init__(self):
        self.conn = None

    def open(self, driver, server, database, username, password):
        """DBをオープンする"""
        conn_str = f'Driver={{{driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

        # Windows認証の場合
        # conn_str = f'Driver={{{driver}}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'

        self.conn = pyodbc.connect(conn_str, autocommit=True)

    def close(self):
        """DBをクローズする"""
        self.conn.close()

    def insert_record(self, id, name, filename):
        """Worklistテーブルにレコードを追加する。"""
        cursor = self.conn.cursor()
        params = (id, name, filename)
        sql_str = f"INSERT INTO Worklist(id, name, filename) VALUES(?, ?, ?)"
        cursor.execute(sql_str, params)

    def is_record_existed(self, id):
        """Worklistに同じidを持つレコードがあるか？"""
        cursor = self.conn.cursor()
        sql_str = f"SELECT COUNT(*) FROM Worklist WHERE id={id}"
        cursor.execute(sql_str)
        row = cursor.fetchone()
        return row[0] > 0

    def get_records(self, id=None):
        """
        Worklistの同じidを持つレコード全てを取得する。
        idがNone(=null)のときは、レコード全件を取得する。
        """
        cursor = self.conn.cursor()
        sql_str = "SELECT * FROM Worklist"
        if id is not None:
            sql_str += f" WHERE id={id}"
        cursor.execute(sql_str)
        rows = cursor.fetchall()
        return rows


# このファイルを
# $ python sqlserver.py
# のように実行時の主プログラムとして使う場合は
# 以下のif文が実行される。（このとき、__name__の値が'__main__'）
# import sqlserver
# のようにモジュールとして使う場合は
# 以下のif文は実行されない。

if __name__ == '__main__':
    # pythonではネスト（インデント）は必須で厳格。
    # ネストを起こす制御文の終端は:（コロン）が必須。
    # 終端の;（セミコロン）は不要。
    print('Connecting to SQL Server with ODBC driver')
    db = SqlServer()
    db.open(DRIVER, SERVER, DATABASE, USERNAME, PASSWORD)
    print('Connected\n')

    # レコード3を追加
    i = 3
    if db.is_record_existed(i):
        print('Record %d is existed' % i)
    else:
        db.insert_record(i, 'Ichiro Arai', '00000003.jpg')
        print('Record %d inserted' % i)

    # レコード4を追加
    if db.is_record_existed(4):
        print('Record %d is existed' % 4)
    else:
        db.insert_record(4, 'Jiro Arai', '00000004.jpg')
        print('Record %d inserted' % 4)

    # レコード3を取得
    rows = db.get_records(3)
    print('Record {0} is now {1} record(s)'.format(3, len(rows)))
    print()  # もう一つ改行

    # レコード全件を取得
    rows = db.get_records()
    for row in rows:
        print('id: %d, name: %s, filename: %s' % tuple(row))

    db.close()

    print()  # もう一つ改行
    print('Finish')
