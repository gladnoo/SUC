import sqlite3

conn = sqlite3.connect('kanban.db')
c = conn.cursor()

c.execute('CREATE TABLE IF NOT EXISTS comunicados (id INTEGER PRIMARY KEY AUTOINCREMENT, titulo TEXT, texto TEXT, autor TEXT, criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP)')
c.execute('CREATE TABLE IF NOT EXISTS eventos (id INTEGER PRIMARY KEY AUTOINCREMENT, titulo TEXT, data TEXT, tipo TEXT DEFAULT "evento", criado_por TEXT)')

conn.commit()
conn.close()
print('OK!')