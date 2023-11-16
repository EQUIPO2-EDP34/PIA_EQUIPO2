import sqlite3
import os

# Especificamos el nombre del archivo de la bd a eliminar
db_file = 'BD_NISSAN.db'

# Verificamos si el archivo de la bd existe y lo eliminamos
if os.path.exists(db_file):
    os.remove(db_file)
    print(f"La base de datos '{db_file}' ha sido eliminada.")
else:
    print(f"La base de datos '{db_file}' no existe.")
