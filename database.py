import sqlite3 as sq
import random
import string
from openpyxl import load_workbook


def generate_key(length):
    letters_and_digits = string.ascii_letters + string.digits
    rand_string = ''.join(random.sample(letters_and_digits, length))
    return rand_string


conn = sq.connect('Пациенты1.db')
cur = conn.cursor()

cur.executescript(f'''PRAGMA foreign_keys=on;
    CREATE TABLE IF NOT EXISTS ключи (фамилия TEXT PRIMARY KEY,
    ключ TEXT)''')
keys = [('Тарасова', 'ysQfR'), ('Дылдин', 'CGawR'), ('Сорокина', 'cb9KM'), ('Бородина', 'JIS7b'), ('ЗУБОВ', 'ztFh5'),
        ('Немцев', 'hsvWc'), ('Абрамов', 'mciRl'), ('Корзухина', 'b0Hvn'), ('Баринов', 'gPrys'), ('Тиккоева', 'AiY0p'),
        ('Чеснокова', 'uGwVC'), ('ЗУЕВА', 'DvNeg')]
cur.executemany("INSERT INTO ключи VALUES(?, ?)", keys)
conn.commit()


def is_new_patient(name, records):
    if name in records.keys():
        return records[name]


cur.executescript(f'''PRAGMA foreign_keys=on;
        CREATE TABLE IF NOT EXISTS анализы(
        num INT PRIMARY KEY,
        Ключ TEXT,
        ФИО TEXT,
        Пол TEXT,   
        Дата_Анализа TEXT,
        Возраст TEXT,
        Диагноз_основной TEXT,
        Диагноз_сопутствующий TEXT,
        Гены1 TEXT,
        Гены2 TEXT,
        Гены3 TEXT,  
        Сезон TEXT,
        Лейкоциты_WBC TEXT,
        Лимфоциты_LYMF TEXT,
        Моноциты_MON TEXT,
        Нейтрофилы_NEU TEXT,
        Гемоглобин_HGB TEXT,
        Тромбоциты_PLT TEXT,
        Эозинофилы_EOS TEXT,
        Базофилы_BAS TEXT,
        Общие_B_лимфоциты TEXT,
        Общие_T_лимфоциты TEXT,
        Т_хелперы TEXT,
        Т_цитотоксические лимфоциты TEXT,
        Соотношение_CD3_CD4_CD3_CD8 TEXT,
        NK_клетки_цитолитические TEXT,
        Общие_NK_клетки TEXT,
        Циркулирующие_имунные_комплексы TEXT,
        НСТ_тест_спонтанный TEXT,
        НСТ_тест_стимулированный TEXT,  
        CD3_IFNy_стимулированый TEXT,
        CD3_IFNy_спонтанный TEXT,
        CD3_TNFa_стимулированный TEXT,
        CD3_TNFa_спонтанный TEXT,
        CD3_IL2_стимулированный TEXT,
        CD3_IL2_спонтанный TEXT,
        ФНО TEXT)''')
conn.commit()


cur.executescript(f'''PRAGMA foreign_keys=on;
        CREATE TABLE IF NOT EXISTS ключи(
            Ключ TEXT PRIMARY KEY,
            ФИО TEXT)''')
conn.commit()

file = 'Результаты_Анализов.xlsx'

wb = load_workbook(file)
db_list = []
name_list = []
c = 0

conn = sq.connect('Пациенты1.db')
cursor = conn.cursor()
sqlite_select_query = """SELECT * from ключи"""
cursor.execute(sqlite_select_query)
records = dict(cursor.fetchall())

temp = []

for sheet in wb.worksheets[2:]:
    a = str(sheet)[12:-2]
    xl_list = wb[a]
    name = xl_list['C3'].value
    name_list.append(name)
for x in name_list:
    if x not in temp:
        temp.append(x)

for sheet in wb.worksheets[2:]:
    a = str(sheet)[12:-2]
    xl_list = wb[a]
    date1 = str(xl_list['C5'].value).replace('-', '.')
    date1 = date1[:10]
    name = xl_list['C3'].value
    if name in temp:
        key = is_new_patient(name, records)
    else:
        key = generate_key(5)
        cur.execute("INSERT OR IGNORE INTO ключи VALUES(?, ?)", (name, key))
        conn.commit()

    db_list.append((c, key, xl_list['C3'].value, 'NULL', date1, xl_list['C4'].value, 'NULL', 'NULL', 'NULL', 'NULL',
                    'NULL', xl_list['A6'].value,

                    xl_list['C9'].value, xl_list['C10'].value, xl_list['C12'].value, xl_list['C14'].value,
                    xl_list['C16'].value,
                    xl_list['C18'].value, 'NULL', 'NULL',

                    xl_list['C27'].value, xl_list['C29'].value, xl_list['C31'].value,
                    xl_list['C33'].value, 'NULL', xl_list['C37'].value, xl_list['C38'].value, xl_list['C24'].value,
                    xl_list['C51'].value, xl_list['C57'].value,

                    xl_list['C46'].value, xl_list['C47'].value, xl_list['C41'].value,
                    xl_list['C40'].value, xl_list['C42'].value, xl_list['C43'].value, 'NULL'))
    c += 1

conn = sq.connect('Пациенты1.db')
cursor = conn.cursor()

cursor.executemany("INSERT INTO анализы VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);",db_list)
conn.commit()
