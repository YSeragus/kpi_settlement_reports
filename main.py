import requests
import pandas as pd

# ================= КОНФІГУРАЦІЯ =================

# Максимальное количество людей в одной пустой (унисекс) комнате
ROOM_CAPACITY = 3

# Режим заполнения:
# 1 - Сначала селим на уже существующие гендерные места, когда они закончатся — открываем пустые комнаты.
# 2 - Сначала селим в пустые комнаты (унисекс), когда они закончатся — доселяем на существующие гендерные места.
FILL_MODE = 1

# Структура: { номер_общежития: {'male': чол_мест, 'female': жін_мест, 'unisex': унісекс_мест} }
DORM_CAPACITIES = {
    3:  {'male':  10,  'female':   1,  'unisex':   13 },
    4:  {'male':  80,  'female':  17,  'unisex':   91 },
    6:  {'male':  25,  'female':   9,  'unisex':   21 },
    7:  {'male':  46,  'female':  50,  'unisex':   45 }, 
    8:  {'male': 158,  'female':  12,  'unisex':   12 },
    11: {'male':   7,  'female':   8,  'unisex':   50 },
    12: {'male':  15,  'female':  13,  'unisex':   52 },
    13: {'male':  25,  'female':  16,  'unisex':  114 },
    14: {'male':  17,  'female':  35,  'unisex':   66 },
    15: {'male':  65,  'female':  77,  'unisex':    0 },
    16: {'male':  122, 'female':  113, 'unisex':   88 },
    17: {'male':  2,   'female':  1,   'unisex':    2 },
    18: {'male':  153, 'female':  124, 'unisex':    9 },
    19: {'male':  101, 'female':  149, 'unisex':   45 },
    20: {'male':   67, 'female':   50, 'unisex':   90 }
}

# ================= КЛАС МЕНЕДЖЕРА ГУРТОЖИТКУ =================

class DormManager:
    def __init__(self, dorm_id, male_cap, female_cap, unisex_cap):
        self.dorm_id = dorm_id
        
        # Сохраняем изначальные квоты для статистики
        self.init_male = male_cap
        self.init_female = female_cap
        self.init_unisex = unisex_cap
        self.initial_capacity = male_cap + female_cap + unisex_cap
        
        # Места в уже "начатых" комнатах
        self.male_spots = male_cap
        self.female_spots = female_cap
        
        # Статистика реального заселения
        self.settled_male_existing = 0
        self.settled_female_existing = 0
        self.settled_male_unisex = 0
        self.settled_female_unisex = 0
        
        # Разбиваем унисекс квоту на массив пустых комнат 
        self.empty_rooms = []
        remaining_unisex = unisex_cap
        while remaining_unisex > 0:
            self.empty_rooms.append(min(remaining_unisex, ROOM_CAPACITY))
            remaining_unisex -= min(remaining_unisex, ROOM_CAPACITY)
            
        self.partial_rooms = {'Ч': [], 'Ж': []}
        self.students = [] 

    def try_existing(self, gender):
        if gender == 'Ч' and self.male_spots > 0:
            self.male_spots -= 1
            self.settled_male_existing += 1
            return True
        if gender == 'Ж' and self.female_spots > 0:
            self.female_spots -= 1
            self.settled_female_existing += 1
            return True
        return False

    def try_unisex(self, gender):
        if self.partial_rooms[gender]:
            self.partial_rooms[gender][0] -= 1
            if self.partial_rooms[gender][0] == 0:
                self.partial_rooms[gender].pop(0)
                
            if gender == 'Ч': self.settled_male_unisex += 1
            else: self.settled_female_unisex += 1
            return True
        
        if self.empty_rooms:
            room_cap = self.empty_rooms.pop(0)
            if room_cap > 1:
                self.partial_rooms[gender].append(room_cap - 1)
                
            if gender == 'Ч': self.settled_male_unisex += 1
            else: self.settled_female_unisex += 1
            return True
            
        return False

    def allocate(self, student_info):
        gender = student_info.get('gender')
        if gender not in ['Ч', 'Ж']:
            return False 
            
        assigned = False
        
        if FILL_MODE == 1:
            assigned = self.try_existing(gender) or self.try_unisex(gender)
        elif FILL_MODE == 2:
            assigned = self.try_unisex(gender) or self.try_existing(gender)
            
        if assigned:
            record = student_info.copy()
            del record['gender']
            self.students.append(record)
            
        return assigned

# ================= ОСНОВНІ ФУНКЦІЇ =================

def get_api_data(api_url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json'
    }
    response = requests.get(api_url, headers=headers)
    
    if response.status_code != 200:
        print(f"Помилка доступу: {response.status_code}")
        return None
        
    json_data = response.json()
    return pd.DataFrame(json_data['rows'])

def allocate_dorms(df, capacities):
    df = df.dropna(subset=['rating'])
    df = df.sort_values(by='rating', ascending=False).reset_index(drop=True)
    
    managers = {dorm: DormManager(dorm, caps['male'], caps['female'], caps['unisex']) 
                for dorm, caps in capacities.items()}
                
    unallocated = []     
    denied_students = [] 
    
    for index, student in df.iterrows():
        name = student['name']
        faculty = student.get('faculty', '')
        gender = student.get('gender', '') 
        global_rating = student['rating']
        status = student.get('status', '')   
        priorities = student.get('priority_ratings', [])
        
        if status == 'denied':
            denied_students.append({
                'ПІБ': name,
                'Стать': gender,
                'Факультет': faculty,
                'Балл': global_rating,
                'Статус': status
            })
            continue 
            
        assigned = False
        
        if isinstance(priorities, list) and len(priorities) > 0:
            priorities = sorted(priorities, key=lambda x: x['priority'])
            
            for p in priorities:
                dorm = p.get('dorm')
                
                if dorm in managers:
                    student_info = {
                        'ПІБ': name,
                        'Стать': gender,
                        'Факультет': faculty,
                        'Балл': p.get('rating', global_rating),
                        'Пріоритет': p['priority'],
                        'gender': gender
                    }
                    
                    if managers[dorm].allocate(student_info):
                        assigned = True
                        break 
                        
        if not assigned:
            unallocated.append({
                'ПІБ': name, 
                'Стать': gender,
                'Факультет': faculty,
                'Балл': global_rating
            })
            
    return managers, unallocated, denied_students

if __name__ == "__main__":
    
    API_URL = "https://settlement.kpi.ua/api/public/ratings"
    
    print("Завантажуємо дані по API...\n")
    data = get_api_data(API_URL)
    
    if data is not None:
        print(f"Використовуваний режим заповнення: {'Спочатку існуючі місця' if FILL_MODE == 1 else 'Спочатку порожні (унісекс) кімнати'}")
        print("Розподіляємо студентів по гуртожитках...\n")
        
        managers, unallocated, denied_students = allocate_dorms(data, DORM_CAPACITIES)
        
        # Используем xlsxwriter для управления шириной колонок
        with pd.ExcelWriter('settlement_results.xlsx', engine='xlsxwriter') as writer:
            total_capacity_all = 0
            
            for dorm, manager in managers.items():
                students_in_dorm = manager.students
                capacity = manager.initial_capacity
                total_capacity_all += capacity
                
                print(f"--- Гуртожиток №{dorm} ---")
                print(f"Виділено місць: {capacity}, Заселено: {len(students_in_dorm)}")
                
                sheet_name = f'Гурт.{dorm}, Місць {capacity}'
                
                if students_in_dorm:
                    df_dorm = pd.DataFrame(students_in_dorm)
                    
                    # --- ВЫВОД В ТЕРМИНАЛ КАК РАНЬШЕ ---
                    print(df_dorm[['ПІБ', 'Факультет', 'Балл', 'Пріоритет']].head(3).to_string(index=False))
                    
                    # Создаем 3 пустые колонки (F, G, H), чтобы статистика сдвинулась к I и J
                    for col_name in [' ', '  ', '   ']:
                        df_dorm[col_name] = ""
                    
                    # Названия для новых колонок
                    col_i_name = 'Місткість'
                    col_j_name = 'Заселено'
                    
                    # Добавляем колонки I и J
                    df_dorm[col_i_name] = ""
                    df_dorm[col_j_name] = ""
                    
                    # Если студентов меньше 3, добавляем пустые строки, чтобы вместилась статистика
                    while len(df_dorm) < 3:
                        df_dorm.loc[len(df_dorm)] = ""
                    
                    df_dorm = df_dorm.fillna("")
                    
                    # Записываем статистику в столбики I и J
                    df_dorm.loc[0, col_i_name] = f"Ч: {manager.init_male}"
                    df_dorm.loc[1, col_i_name] = f"Ж: {manager.init_female}"
                    df_dorm.loc[2, col_i_name] = f"Унісекс: {manager.init_unisex}"
                    
                    # Теперь записываем поселенных Ч и Ж только на их основные места, а унисекс выводим с разделением
                    df_dorm.loc[0, col_j_name] = f"Ч: {manager.settled_male_existing}"
                    df_dorm.loc[1, col_j_name] = f"Ж: {manager.settled_female_existing}"
                    df_dorm.loc[2, col_j_name] = f"Ч:{manager.settled_male_unisex}, Ж:{manager.settled_female_unisex}"
                    
                    df_dorm.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # --- НАСТРОЙКА ШИРИНЫ СТОЛБЦОВ ---
                    worksheet = writer.sheets[sheet_name]
                    worksheet.set_column('A:A', 15) # ПІБ
                    worksheet.set_column('B:B', 6)  # Стать
                    worksheet.set_column('C:C', 10) # Факультет
                    worksheet.set_column('D:D', 12) # Балл
                    worksheet.set_column('E:E', 10) # Пріоритет
                    worksheet.set_column('F:H', 5)  # Пустые разделители
                    worksheet.set_column('I:I', 10) # Статистика 1
                    worksheet.set_column('J:J', 20) # Статистика 2
                else:
                    print("Нікого не заселено.")
                print("-" * 25 + "\n")
            
            # --- ЛИСТ ДЛЯ ТЕХ, КТО НЕ ПРОШЕЛ ---
            df_unallocated = pd.DataFrame(unallocated, columns=['ПІБ', 'Стать', 'Факультет', 'Балл'])
            df_unallocated.to_excel(writer, sheet_name='Не пройшли', index=False)
            
            if len(df_unallocated) > 0:
                ws_unallocated = writer.sheets['Не пройшли']
                ws_unallocated.set_column('A:A', 15)
                ws_unallocated.set_column('B:B', 6)
                ws_unallocated.set_column('C:C', 10)
                ws_unallocated.set_column('D:D', 12)
            
            # --- ЛИСТ ДЛЯ ОТКАЗОВ (DENIED) ---
            df_denied = pd.DataFrame(denied_students, columns=['ПІБ', 'Стать', 'Факультет', 'Балл', 'Статус'])
            df_denied.to_excel(writer, sheet_name='Відмови', index=False)
            
            if len(df_denied) > 0:
                ws_denied = writer.sheets['Відмови']
                ws_denied.set_column('A:A', 15)
                ws_denied.set_column('B:B', 6)
                ws_denied.set_column('C:C', 10)
                ws_denied.set_column('D:D', 12)
                ws_denied.set_column('E:E', 12)
            
            
            # --- ВЫВОД ФИНАЛЬНОЙ СТАТИСТИКИ ---
            print(f"Кому не вистачило місць: {len(df_unallocated)} осіб")
            print(f"Всього місць: {total_capacity_all}, Подано: {len(data)-len(df_denied)}, Відмов: {len(df_denied)}")
                
        print("\n Готово! Повні списки по кожному гуртожитку збережено у файл 'settlement_results.xlsx'.")