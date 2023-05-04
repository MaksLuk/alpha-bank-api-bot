import aiosqlite
import asyncio
from dataclasses import dataclass
import excel


@dataclass
class RequestsData:
    all_requests: int
    success: int
    errors: int
    filename: str
    def __repr__(self):
        return f'▫️Всего: {self.all_requests}\n▫️Успешно: {self.success}\n▫️С ошибками: {self.errors}\n'


class Database:
    @classmethod
    async def init_db(cls):
        ''' создаёт таблицу в бд, если её нет, и возвращает идекс следующей проверки '''
        async with aiosqlite.connect("db.db") as db:
            await db.execute("""CREATE TABLE IF NOT EXISTS Requests(
                                    request_date TEXT DEFAULT CURRENT_TIMESTAMP, 
                                    status INT, 
                                    checking_number INT,
                                    INN TEXT, 
                                    FIO TEXT,
                                    phone TEXT,
                                    organization TEXT,
                                    script TEXT,
                                    city TEXT,
                                    comment TEXT);""")
            await db.commit()
            cursor = await db.execute("SELECT MAX(checking_number) FROM Requests;")
            max_number = await cursor.fetchone()
            if max_number[0] == None:
                return 1
            return max_number[0] + 1

    @classmethod
    async def insert_db(cls, data):
        async with aiosqlite.connect("db.db") as db:
            await db.execute("""INSERT INTO Requests(status, checking_number, INN, FIO, 
                                    phone, organization, script, city, comment) 
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);""", data)
            await db.commit()

    @classmethod
    async def get_statistics(cls):
        helper = {0: 'день', 7: 'неделя', 30: 'месяц'}
        result = []
        async with aiosqlite.connect("db.db") as db:
            for i in [0, 7, 30]:
                cursor = await db.execute(f"SELECT COUNT(*) FROM Requests WHERE date(request_date) >= date('now', '-{i} days') and status = 1;")
                success = await cursor.fetchone()
                cursor = await db.execute(f"SELECT COUNT(*) FROM Requests WHERE date(request_date) >= date('now', '-{i} days') and status != 1;")
                errors = await cursor.fetchone()
                
                cursor = await db.execute(f"""SELECT INN, FIO,  phone, organization, script, city, comment FROM Requests
                                              WHERE date(request_date) >= date('now', '-{i} days') and status = 1;""")
                success_list = await cursor.fetchall()
                cursor = await db.execute(f"""SELECT INN, FIO,  phone, organization, script, city, comment, status FROM Requests
                                              WHERE date(request_date) >= date('now', '-{i} days') and status != 1;""")
                errors_list = await cursor.fetchall()
                for j in range(len(errors_list)):
                    if errors_list[j][-1] == 777:
                        errors_list[j] = errors_list[j][:-1] + ('Указаны не все параметры',)
                    elif errors_list[j][-1] == 888:
                        errors_list[j] = errors_list[j][:-1] + ('Город не найден',)
                
                filename = excel.create_workbook([('ИНН', 'ФИО', 'Телефон', 'Организация', 'Сценарий', 'Город', 'Комментарий'), ] + success_list, 
                                                 [('ИНН', 'ФИО', 'Телефон', 'Организация', 'Сценарий', 'Город', 'Комментарий', 'Код ошибки'), ] + errors_list, 
                                                 helper[i])
                
                result.append(RequestsData(success[0] + errors[0], success[0], errors[0], filename))
        return result
            

