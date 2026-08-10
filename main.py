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

# ================= КЛАСИ ДЛЯ РОЗПОДІЛУ =================

class DormManager:
    def __init__(self, dorm_id, male_cap, female_cap, unisex_cap):
        self.dorm_id = dorm_id
        
        self.init_male = male_cap
        self.init_female = female_cap
        self.init_unisex = unisex_cap
        self.initial_capacity = male_cap + female_cap + unisex_cap
        
        self.male_spots = male_cap
        self.female_spots = female_cap
        
        self.settled_male_existing = 0
        self.settled_female_existing = 0
        self.settled_male_unisex = 0
        self.settled_female_unisex = 0
        
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


# ================= КЛАС ДЛЯ ГЕНЕРАЦІЇ ЗВІТІВ (EXCEL & TERMINAL) =================

class ExcelReportGenerator:
    def __init__(self, filename, data, managers, unallocated, denied_students):
        self.filename = filename
        self.data = data
        self.managers = managers
        self.unallocated = unallocated
        self.denied_students = denied_students
        self.total_capacity_all = sum(m.initial_capacity for m in managers.values())

    def _get_score_bins(self, step):
        """Генерує кошики балів чітко від 110 до 250 із заданим кроком"""
        edges = list(range(110, 251, step))
        if edges[-1] < 250:
            edges.append(250)
            
        # Щоб включити бал рівно 250.0 (при right=False), трохи збільшуємо останню межу
        bins = edges.copy()
        bins[-1] = 250.1 
        
        labels = []
        for i in range(len(edges) - 1):
            labels.append(f"{edges[i]}-{edges[i+1]}")
            
        return bins, labels

    def generate(self):
        with pd.ExcelWriter(self.filename, engine='xlsxwriter') as writer:
            self.workbook = writer.book
            
            # Створення стилів
            self.fmt_sec = self.workbook.add_format({'font_color': '#808080', 'font_size': 9, 'italic': True})
            self.fmt_sec_hdr = self.workbook.add_format({'font_color': '#808080', 'font_size': 9, 'italic': True, 'bold': True, 'bottom': 1})
            
            # Більш темні кольори без жирного шрифту (bold=False за замовчуванням)
            self.fmt_orange = self.workbook.add_format({'font_color': '#C0504D'}) # Темно-помаранчевий/теракотовий для бонусу
            self.fmt_blue = self.workbook.add_format({'font_color': "#4D5EFA"})   # Темно-синій для пільг
            
            # Генерація сторінок
            self._write_general_statistics(writer)
            self._write_dorm_sheets(writer)
            self._write_other_sheets(writer)
            
            print(f"Кому не вистачило місць (або немає пріоритетів): {len(self.unallocated)} осіб")
            print(f"Всього місць: {self.total_capacity_all}, Подано(заявок): {len(self.data)-len(self.denied_students)}, Відмов: {len(self.denied_students)}")
                
        print(f"\n✅ Готово! Повні списки по кожному гуртожитку та листи зі статистикою збережено у файл '{self.filename}'.")

    def _write_general_statistics(self, writer):
        all_settled = []
        for manager in self.managers.values():
            for s in manager.students:
                s_copy = s.copy()
                s_copy['Гуртожиток'] = f"№{manager.dorm_id}"
                all_settled.append(s_copy)
        df_all_settled = pd.DataFrame(all_settled)

        valid_data = self.data[(self.data['status'] != 'denied') & self.data['rating'].notna()]
        unique_students = valid_data.drop_duplicates(subset=['name'])
        
        # 1. Загальна інформація
        general_info = pd.DataFrame({
            'Метрика': ['Всього місць', 'Унікальних заявок', 'Чоловіків', 'Жінок'],
            'Значення': [self.total_capacity_all, len(unique_students), unique_students['gender'].value_counts().get('Ч', 0), unique_students['gender'].value_counts().get('Ж', 0)]
        })
        
        # 2. Факультети
        faculty_counts = unique_students['faculty'].value_counts().reset_index()
        faculty_counts.columns = ['Факультет', 'Кількість']
        
        # Записуємо таблиці у стовпці A та B один під одним
        current_row = 0
        general_info.to_excel(writer, sheet_name='Статистика', startrow=current_row, startcol=0, index=False)
        row_gen = current_row
        current_row += len(general_info) + 2
        
        faculty_counts.to_excel(writer, sheet_name='Статистика', startrow=current_row, startcol=0, index=False)
        row_fac = current_row
        current_row += len(faculty_counts) + 2
        
        if not df_all_settled.empty:
            # 3. Пріоритети 
            gen_priority = df_all_settled['Пріоритет'].value_counts().reset_index()
            gen_priority.columns = ['Пріоритет', 'Кількість']
            gen_priority['Пріоритет'] = gen_priority['Пріоритет'].astype(str) + ' пріоритет'
            
            gen_priority.to_excel(writer, sheet_name='Статистика', startrow=current_row, startcol=0, index=False)
            row_pri = current_row
            current_row += len(gen_priority) + 2

            # 4. Розподіл балів (Крок 10 балів)
            df_all_settled['Балл'] = pd.to_numeric(df_all_settled['Балл'], errors='coerce')
            score_bins, score_labels = self._get_score_bins(step=10)
            
            df_all_settled['Група балів'] = pd.cut(df_all_settled['Балл'], bins=score_bins, labels=score_labels, right=False)
            gen_scores_dist = df_all_settled['Група балів'].value_counts().reindex(score_labels).fillna(0).reset_index()
            gen_scores_dist.columns = ['Діапазон балів', 'Кількість']
            
            gen_scores_dist.to_excel(writer, sheet_name='Статистика', startrow=current_row, startcol=0, index=False)
            row_scores = current_row
            current_row += len(gen_scores_dist) + 2

            # 5. Пільговики vs Звичайні (Загальне)
            df_all_settled['Категорія пільги'] = df_all_settled['Пільга'].apply(lambda x: 'Пільговик' if pd.notna(x) and str(x).strip() != '' else 'Без пільг')
            gen_benefit_ratio = df_all_settled['Категорія пільги'].value_counts().reset_index()
            gen_benefit_ratio.columns = ['Категорія', 'Кількість']
            
            gen_benefit_ratio.to_excel(writer, sheet_name='Статистика', startrow=current_row, startcol=0, index=False)
            row_ben_ratio = current_row
            current_row += len(gen_benefit_ratio) + 2

            # 6. Мінімальний бал по гуртожитках (Поріг входу)
            min_scores = df_all_settled.groupby('Гуртожиток')['Балл'].min().reset_index()
            min_scores.columns = ['Гуртожиток', 'Мін. бал']
            min_scores = min_scores.sort_values('Мін. бал', ascending=False)
            
            min_scores.to_excel(writer, sheet_name='Статистика', startrow=current_row, startcol=0, index=False)
            row_min_scores = current_row
            current_row += len(min_scores) + 2

        worksheet_stat = writer.sheets['Статистика']
        worksheet_stat.set_column('A:A', 30)
        worksheet_stat.set_column('B:B', 15)

        worksheet_stat.set_tab_color('#FFFF99') # Світло-жовтий колір
        
        # Графіки для загальної сторінки
        if not df_all_settled.empty:
            self._draw_pie_chart(worksheet_stat, 'Співвідношення статей', 'Статистика', row_gen+3, 0, 2, 'D2')
            self._draw_pie_chart(worksheet_stat, 'Заселення за пріоритетами', 'Статистика', row_pri+1, 0, len(gen_priority), 'L2')
            
            self._draw_column_chart(worksheet_stat, 'Кількість заявок за факультетами', 'Статистика', row_fac+1, 0, len(faculty_counts), 'D18', width=600, color='#4F81BD')
            self._draw_column_chart(worksheet_stat, 'Розподіл балів (Загальне)', 'Статистика', row_scores+1, 0, len(score_labels), 'L18', width=600, color='#8064A2')
            
            self._draw_doughnut_chart(worksheet_stat, 'Пільговики / Без пільг', 'Статистика', row_ben_ratio+1, 0, len(gen_benefit_ratio), 'D34')
            
            # Мінімальний бал (Поріг входу)
            self._draw_column_chart(worksheet_stat, 'Мінімальний бал (поріг входу)', 'Статистика', row_min_scores+1, 0, len(min_scores), 'L34', width=600, color='#9BBB59')

    def _write_dorm_sheets(self, writer):
        for dorm, manager in self.managers.items():
            students_in_dorm = manager.students
            capacity = manager.initial_capacity
            sheet_name = f'Гурт.{dorm}, Місць {capacity}'
            
            print(f"--- Гуртожиток №{dorm} ---")
            print(f"Виділено місць: {capacity}, Заселено: {len(students_in_dorm)}")
            
            if students_in_dorm:
                df_analysis = pd.DataFrame(students_in_dorm)
                df_analysis['Балл'] = pd.to_numeric(df_analysis['Балл'], errors='coerce')
                
                # Не відображаємо колонку Пільга у самій таблиці студентів
                df_display = df_analysis[['ПІБ', 'Стать', 'Факультет', 'Балл', 'Пріоритет']].copy()
                print(df_display.head(3).to_string(index=False))
                
                # Додавання порожніх стовпців для вирівнювання
                for col_name in [' ', '  ', '   ']:
                    df_display[col_name] = ""
                
                df_display['Місткість'] = ""
                df_display['Заселено'] = ""
                
                while len(df_display) < 3:
                    df_display.loc[len(df_display)] = ""
                df_display = df_display.fillna("")
                
                # Запис даних місткості
                df_display.loc[0, 'Місткість'] = f"Ч: {manager.init_male}"
                df_display.loc[1, 'Місткість'] = f"Ж: {manager.init_female}"
                df_display.loc[2, 'Місткість'] = f"Унісекс: {manager.init_unisex}"
                df_display.loc[0, 'Заселено'] = f"Ч: {manager.settled_male_existing}"
                df_display.loc[1, 'Заселено'] = f"Ж: {manager.settled_female_existing}"
                df_display.loc[2, 'Заселено'] = f"Ч:{manager.settled_male_unisex}, Ж:{manager.settled_female_unisex}"
                
                df_display.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]

                if len(students_in_dorm) >= capacity:
                    worksheet.set_tab_color('#FF9999') # Светло-красный (заполнено)
                else:
                    worksheet.set_tab_color('#99FF99') # Светло-зеленый (есть места)
                
                # Фарбуємо бали тих, хто має бонус, і ПІБ тих, хто має пільгу
                for r_idx, row in df_analysis.iterrows():
                    if row.get('Бонус', False):
                        worksheet.write(r_idx + 1, 3, row['Балл'], self.fmt_orange)
                    
                    benefit = row.get('Пільга', '')
                    if pd.notna(benefit) and str(benefit).strip() != '':
                        worksheet.write(r_idx + 1, 0, row['ПІБ'], self.fmt_blue)
                
                # Формування додаткової статистики (Крок 30 балів для гуртожитків)
                score_bins, score_labels = self._get_score_bins(step=30)
                df_analysis['Група балів'] = pd.cut(df_analysis['Балл'], bins=score_bins, labels=score_labels, right=False)
                
                df_gender = df_analysis['Стать'].value_counts().reset_index()
                df_fac = df_analysis['Факультет'].value_counts().reset_index()
                df_pri = df_analysis['Пріоритет'].value_counts().reset_index()
                df_pri['Пріоритет'] = df_pri['Пріоритет'].astype(str) + ' пріоритет'
                df_score = df_analysis['Група балів'].value_counts().reindex(score_labels).fillna(0).reset_index()
                df_bonus = df_analysis['Бонус'].map({True: 'З бонусом', False: 'Без бонуса'}).value_counts().reset_index()
                
                df_analysis['Категорія пільги'] = df_analysis['Пільга'].apply(lambda x: 'Пільговик' if pd.notna(x) and str(x).strip() != '' else 'Без пільг')
                df_benefit = df_analysis['Категорія пільги'].value_counts().reset_index()
                
                tables = {
                    'gender': ('Стать', 'Кількість', df_gender),
                    'fac': ('Факультет', 'Кількість', df_fac),
                    'pri': ('Пріоритет', 'Кількість', df_pri),
                    'score': ('Діапазон балів', 'Кількість', df_score),
                    'bonus': ('Бонус', 'Кількість', df_bonus),
                    'benefit': ('Пільга', 'Кількість', df_benefit)
                }
                
                chart_refs = {}
                start_row = 8  
                for key, (t1, t2, df_t) in tables.items():
                    chart_refs[key] = {'start': start_row, 'len': len(df_t)}
                    worksheet.write(start_row, 8, t1, self.fmt_sec_hdr)
                    worksheet.write(start_row, 9, t2, self.fmt_sec_hdr)
                    for row_num, row_data in enumerate(df_t.itertuples(index=False)):
                        worksheet.write(start_row + 1 + row_num, 8, row_data[0], self.fmt_sec)
                        worksheet.write(start_row + 1 + row_num, 9, row_data[1], self.fmt_sec)
                    start_row += len(df_t) + 2
                
                worksheet.set_column('A:A', 25); worksheet.set_column('B:B', 6); worksheet.set_column('C:C', 12)
                worksheet.set_column('D:D', 12); worksheet.set_column('E:E', 10); worksheet.set_column('F:H', 3)
                worksheet.set_column('I:I', 15); worksheet.set_column('J:J', 20); worksheet.set_column('K:K', 3)
                
                # Малювання графіків
                if chart_refs['gender']['len'] > 0:
                    r_st, r_ln = chart_refs['gender']['start'], chart_refs['gender']['len']
                    self._draw_pie_chart(worksheet, 'Співвідношення статей', sheet_name, r_st+1, 8, r_ln, 'L2', 350, 250)
                
                if chart_refs['pri']['len'] > 0:
                    r_st, r_ln = chart_refs['pri']['start'], chart_refs['pri']['len']
                    self._draw_pie_chart(worksheet, 'Пріоритети заселення', sheet_name, r_st+1, 8, r_ln, 'R2', 350, 250)
                
                if chart_refs['fac']['len'] > 0:
                    r_st, r_ln = chart_refs['fac']['start'], chart_refs['fac']['len']
                    self._draw_column_chart(worksheet, 'Заселено за факультетами', sheet_name, r_st+1, 8, r_ln, 'L16', width=350, height=280, color='#9BBB59')
                    
                if chart_refs['score']['len'] > 0:
                    r_st, r_ln = chart_refs['score']['start'], chart_refs['score']['len']
                    self._draw_column_chart(worksheet, 'Розподіл балів', sheet_name, r_st+1, 8, r_ln, 'R16', width=350, height=280, color='#8064A2')
                    
                if chart_refs['bonus']['len'] > 0:
                    r_st, r_ln = chart_refs['bonus']['start'], chart_refs['bonus']['len']
                    self._draw_doughnut_chart(worksheet, 'Наявність бонусу', sheet_name, r_st+1, 8, r_ln, 'L31', 350, 250)

                if chart_refs['benefit']['len'] > 0:
                    r_st, r_ln = chart_refs['benefit']['start'], chart_refs['benefit']['len']
                    self._draw_doughnut_chart(worksheet, 'Пільговики / Без пільг', sheet_name, r_st+1, 8, r_ln, 'R31', 350, 250)
            else:
                print("Нікого не заселено.")
            print("-" * 25 + "\n")

    def _write_other_sheets(self, writer):
        # Задаем цвета для листов: Светло-синий и Светло-фиолетовый
        sheet_colors = {
            'Не пройшли': '#99CCFF', 
            'Відмови': '#CC99FF'     
        }
        
        for df_data, sheet_n in [(self.unallocated, 'Не пройшли'), (self.denied_students, 'Відмови')]:
            df_out = pd.DataFrame(df_data)
            if not df_out.empty:
                df_out.to_excel(writer, sheet_name=sheet_n, index=False)
                ws = writer.sheets[sheet_n]
                ws.set_column('A:A', 25); ws.set_column('B:B', 6); ws.set_column('C:C', 12); ws.set_column('D:D', 12); ws.set_column('E:E', 12)
                
                # Устанавливаем цвет из словаря
                if sheet_n in sheet_colors:
                    ws.set_tab_color(sheet_colors[sheet_n])

    # --- ХЕЛПЕРИ ДЛЯ МАЛЮВАННЯ ГРАФІКІВ ---
    def _draw_pie_chart(self, worksheet, title, sheet_name, start_row, cat_col, length, position, width=450, height=300):
        chart = self.workbook.add_chart({'type': 'pie'})
        chart.add_series({
            'name': title,
            'categories': [sheet_name, start_row, cat_col, start_row + length - 1, cat_col],
            'values':     [sheet_name, start_row, cat_col + 1, start_row + length - 1, cat_col + 1],
            'data_labels': {'percentage': True, 'show_name': True}
        })
        chart.set_title({'name': title})
        chart.set_size({'width': width, 'height': height})
        worksheet.insert_chart(position, chart)
        
    def _draw_doughnut_chart(self, worksheet, title, sheet_name, start_row, cat_col, length, position, width=450, height=300):
        chart = self.workbook.add_chart({'type': 'doughnut'})
        chart.add_series({
            'name': title,
            'categories': [sheet_name, start_row, cat_col, start_row + length - 1, cat_col],
            'values':     [sheet_name, start_row, cat_col + 1, start_row + length - 1, cat_col + 1],
            'data_labels': {'percentage': True, 'show_name': True}
        })
        chart.set_title({'name': title})
        chart.set_size({'width': width, 'height': height})
        worksheet.insert_chart(position, chart)

    def _draw_column_chart(self, worksheet, title, sheet_name, start_row, cat_col, length, position, width=600, height=300, color='#4F81BD'):
        chart = self.workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name': 'Значення',
            'categories': [sheet_name, start_row, cat_col, start_row + length - 1, cat_col],
            'values':     [sheet_name, start_row, cat_col + 1, start_row + length - 1, cat_col + 1],
            'fill': {'color': color}
        })
        chart.set_title({'name': title})
        chart.set_legend({'none': True})
        chart.set_size({'width': width, 'height': height})
        worksheet.insert_chart(position, chart)


# ================= ОСНОВНИЙ СКРИПТ =================

def get_api_data(api_url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json'
    }
    response = requests.get(api_url, headers=headers)
    if response.status_code != 200:
        print(f"Помилка доступу: {response.status_code}")
        return None
    return pd.DataFrame(response.json()['rows'])

def allocate_dorms(df, capacities):
    # 1. Отбрасываем тех, у кого нет рейтинга
    valid_df = df.dropna(subset=['rating']).copy()
    
    # Сортируем глобально по рейтингу в самом начале. 
    # Это нужно для тай-брейка: если у двух студентов ОДИНАКОВЫЙ локальный балл в одно общежитие, 
    # победит тот, у кого был выше глобальный балл.
    valid_df = valid_df.sort_values(by='rating', ascending=False).reset_index(drop=True)
    
    unassigned_queue = []
    denied_students = []
    
    # Формируем очередь студентов
    for index, student in valid_df.iterrows():
        name = student['name']
        faculty = student.get('faculty', '')
        gender = student.get('gender', '') 
        benefit = student.get('benefit', '')
        global_rating = student['rating']
        status = student.get('status', '')   
        
        if status == 'denied':
            denied_students.append({
                'ПІБ': name, 'Стать': gender, 'Факультет': faculty,
                'Балл': global_rating, 'Статус': status
            })
            continue 
            
        priorities = student.get('priority_ratings', [])
        if not isinstance(priorities, list):
            priorities = []
            
        priorities = sorted(priorities, key=lambda x: x['priority'])
        
        # Добавляем студента в очередь
        unassigned_queue.append({
            'id': index,
            'name': name,
            'faculty': faculty,
            'gender': gender,
            'benefit': benefit,
            'global_rating': global_rating,
            'has_bonus': student.get('has_priority_bonus', False),
            'priorities': priorities,
            'current_prio_idx': 0 # Начинаем с 0 (то есть с 1-го приоритета)
        })
        
    # Словари для хранения "кандидатов", предварительно зачисленных в каждое общежитие
    dorm_candidates = {dorm: [] for dorm in capacities.keys()}
    unallocated_final = []
    
    # 2. Алгоритм Гейла-Шепли (Отложенное согласие)
    while unassigned_queue:
        s = unassigned_queue.pop(0) # Берем первого из очереди
        
        # Если у студента закончились приоритеты — он не прошел никуда
        if s['current_prio_idx'] >= len(s['priorities']):
            unallocated_final.append({
                'ПІБ': s['name'], 'Стать': s['gender'],
                'Факультет': s['faculty'], 'Балл': s['global_rating']
            })
            continue
            
        # Студент "стучится" в общежитие по своему текущему приоритету
        prio = s['priorities'][s['current_prio_idx']]
        dorm_id = prio.get('dorm')
        
        if dorm_id in dorm_candidates:
            local_rating = prio.get('rating', s['global_rating'])
            bonus = prio.get('has_priority_bonus', s['has_bonus'])
            
            candidate = {
                'student': s,
                'rating': local_rating,
                'priority_num': prio['priority'],
                'bonus': bonus
            }
            
            # Добавляем его к списку претендентов на эту общагу
            dorm_candidates[dorm_id].append(candidate)
            
            # --- ЧЕСТНАЯ СОРТИРОВКА ---
            # Сортируем всех претендентов на это общежитие по ИХ ЛОКАЛЬНОМУ баллу (по убыванию)
            dorm_candidates[dorm_id].sort(key=lambda x: x['rating'], reverse=True)
            
            # Создаем временный менеджер, чтобы проверить, кто влезет по лимитам гендера/унисекс
            caps = capacities[dorm_id]
            temp_manager = DormManager(dorm_id, caps['male'], caps['female'], caps['unisex'])
            
            accepted = []
            for cand in dorm_candidates[dorm_id]:
                stu = cand['student']
                student_info = {
                    'ПІБ': stu['name'],
                    'Стать': stu['gender'],
                    'Факультет': stu['faculty'],
                    'Балл': cand['rating'],
                    'Пріоритет': cand['priority_num'],
                    'Бонус': cand['bonus'],
                    'Пільга': stu['benefit'],
                    'gender': stu['gender']
                }
                
                if temp_manager.allocate(student_info):
                    accepted.append(cand) # Места хватило, оставляем в общаге
                else:
                    # Места не хватило (вытеснили конкуренты с баллами выше)
                    stu['current_prio_idx'] += 1 # Переходим к следующему приоритету
                    unassigned_queue.append(stu) # Возвращаем студента обратно в очередь "неприкаянных"
                    
            # Оставляем в общежитии только тех, кто прошел жесткий отбор
            dorm_candidates[dorm_id] = accepted
            
        else:
            # Если общежития нет в словаре лимитов, сразу шлем на следующий приоритет
            s['current_prio_idx'] += 1
            unassigned_queue.append(s)

    # 3. Финализируем результаты для записи в Excel
    final_managers = {}
    for dorm_id, caps in capacities.items():
        manager = DormManager(dorm_id, caps['male'], caps['female'], caps['unisex'])
        # Кандидаты уже отсортированы по убыванию локального балла, просто заливаем их в менеджер
        for cand in dorm_candidates[dorm_id]:
            stu = cand['student']
            student_info = {
                'ПІБ': stu['name'],
                'Стать': stu['gender'],
                'Факультет': stu['faculty'],
                'Балл': cand['rating'],
                'Пріоритет': cand['priority_num'],
                'Бонус': cand['bonus'],
                'Пільга': stu['benefit'],
                'gender': stu['gender']
            }
            manager.allocate(student_info)
        final_managers[dorm_id] = manager
        
    return final_managers, unallocated_final, denied_students

if __name__ == "__main__":
    API_URL = "https://settlement.kpi.ua/api/public/ratings"
    print("Завантажуємо дані по API...\n")
    data = get_api_data(API_URL)
    
    if data is not None:
        print(f"Використовуваний режим заповнення: {'Спочатку існуючі місця' if FILL_MODE == 1 else 'Спочатку порожні (унісекс) кімнати'}")
        print("Розподіляємо студентів по гуртожитках...\n")
        
        managers, unallocated, denied_students = allocate_dorms(data, DORM_CAPACITIES)
        
        report_generator = ExcelReportGenerator('settlement_results.xlsx', data, managers, unallocated, denied_students)
        report_generator.generate()