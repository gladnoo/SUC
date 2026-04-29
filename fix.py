import sqlite3
conn = sqlite3.connect('kanban.db')
conn.execute("CREATE TABLE IF NOT EXISTS repositorio (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, categoria TEXT NOT NULL, drive_id TEXT NOT NULL, drive_link TEXT NOT NULL, mime_type TEXT, tamanho INTEGER, autor TEXT, criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")
conn.commit()
conn.close()
print('OK')

conn = sqlite3.connect('kanban.db')
conn.execute("""
    CREATE TABLE IF NOT EXISTS atas (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        titulo TEXT NOT NULL,
        data_reuniao TEXT,
        presentes TEXT,
        pauta TEXT,
        deliberacoes TEXT,
        link TEXT,
        autor TEXT,
        criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
""")
conn.commit()
conn.close()
print('OK')

conn = sqlite3.connect('kanban.db')
conn.execute("""
    CREATE TABLE IF NOT EXISTS ramais (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        ramal TEXT NOT NULL,
        cargo TEXT,
        setor TEXT,
        criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
""")
conn.commit()
conn.close()
print('OK')