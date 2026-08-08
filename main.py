import requests
import pandas as pd

def get_api_data(api_url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json'
    }
    response = requests.get(api_url, headers=headers)
    
    if response.status_code != 200:
        print(f"Ошибка доступа: {response.status_code}")
        return None
        
    json_data = response.json()
    df = pd.DataFrame(json_data['rows'])
    return df

def allocate_dorms(df, capacities):
    # Очищаем от пустых значений и сортируем студентов по основному баллу (по убыванию)
    df = df.dropna(subset=['rating'])
    df = df.sort_values(by='rating', ascending=False).reset_index(drop=True)
    
    # Создаем словари и списки
    allocations = {dorm: [] for dorm in capacities.keys()}
    unallocated = []     # Кому не хватило мест
    denied_students = [] # Те, кто отказался или чья заявка отклонена
    
    # Идем по каждому студенту
    for index, student in df.iterrows():
        name = student['name']
        faculty = student.get('faculty', '') # Получаем факультет
        global_rating = student['rating']
        status = student.get('status', '')   # Получаем статус
        priorities = student.get('priority_ratings', [])
        
        # --- НОВАЯ ЛОГИКА: ОБРАБОТКА ОТКАЗОВ ---
        if status == 'denied':
            denied_students.append({
                'ПІБ': name,
                'Факультет': faculty,
                'Балл': global_rating,
                'Статус': status
            })
            continue # Пропускаем этап заселения для этого человека
            
        assigned = False
        
        # Если приоритеты вообще указаны (это список)
        if isinstance(priorities, list) and len(priorities) > 0:
            
            # Сортируем его выборы по приоритету (1, 2, 3...)
            priorities = sorted(priorities, key=lambda x: x['priority'])
            
            for p in priorities:
                dorm = p.get('dorm')
                
                # Если такая общага есть в нашем списке мест
                if dorm in capacities:
                    # Проверяем, остались ли там еще места
                    if len(allocations[dorm]) < capacities[dorm]:
                        # Заселяем, добавляя колонку факультета
                        allocations[dorm].append({
                            'ПІБ': name,
                            'Факультет': faculty,
                            'Балл': p.get('rating', global_rating),
                            'Пройшов/ла за пріоритетом': p['priority']
                        })
                        assigned = True
                        break # Студент заселен, берем следующего
                        
        # Если прошелся по всем приоритетам и мест нигде не было
        if not assigned:
            unallocated.append({
                'ПІБ': name, 
                'Факультет': faculty,
                'Балл': global_rating
            })
            
    return allocations, unallocated, denied_students

if __name__ == "__main__":
    
    API_URL = "https://settlement.kpi.ua/api/public/ratings"
    
    DORM_CAPACITIES = {
        3:  24,
        4:  188,
        6:  55,
        7:  141,
        8:  182,
        11: 65,  
        12: 80,
        13: 155,
        14: 118,
        15: 142,
        16: 323,
        17: 5,
        18: 286,
        19: 295,
        20: 207
    }
    
    print("Скачиваем данные по API...\n")
    data = get_api_data(API_URL)
    
    if data is not None:
        print("Распределяем студентов по общежитиям...\n")
        
        allocations, unallocated, denied_students = allocate_dorms(data, DORM_CAPACITIES)
        
        with pd.ExcelWriter('settlement_results.xlsx') as writer:
            
            # Выводим статистику и сохраняем списки прошедших
            for dorm, capacity in DORM_CAPACITIES.items():
                students_in_dorm = allocations[dorm]
                print(f"--- Общежитие №{dorm} ---")
                print(f"Выделено мест: {capacity}, Заселено: {len(students_in_dorm)}")
                
                if students_in_dorm:
                    df_dorm = pd.DataFrame(students_in_dorm)
                    print(df_dorm.head(3).to_string(index=False)) 
                    df_dorm.to_excel(writer, sheet_name=f'Общ. {dorm}, Мест {capacity}', index=False)
                else:
                    print("Никто не заселен.")
                print("-" * 25 + "\n")
            
            # --- ЛИСТ ДЛЯ ТЕХ, КТО НЕ ПРОШЕЛ ---
            df_unallocated = pd.DataFrame(unallocated, columns=['ПІБ', 'Факультет', 'Балл'])
            print(f"Кому не хватило мест (или нет приоритетов): {len(df_unallocated)} человек")
            df_unallocated.to_excel(writer, sheet_name='Не прошли', index=False)
            
            # --- ЛИСТ ДЛЯ ОТКАЗОВ (DENIED) ---
            df_denied = pd.DataFrame(denied_students, columns=['ПІБ', 'Факультет', 'Балл', 'Статус'])
            df_denied.to_excel(writer, sheet_name='Отказы', index=False)
            
            # --- ВЫВОД ФИНАЛЬНОЙ СТАТИСТИКИ ---
            valid_applications = len(data) - len(df_denied)
            print(f"Всего мест: {sum(DORM_CAPACITIES.values())}, Подано: {valid_applications}, Отказ: {len(df_denied)}")
                
        print("\n✅ Готово! Полные списки по каждому общежитию сохранены в файл 'settlement_results.xlsx'.")