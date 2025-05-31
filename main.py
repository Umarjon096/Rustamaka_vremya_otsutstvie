import pandas as pd
from datetime import datetime, time, timedelta

# --- 1. Загрузка и подготовка данных ---

# Путь к вашему Excel-файлу из отчёта СКУД
file_path = "file.xls"

# Считываем Excel
df = pd.read_excel(file_path)

# Переименуем колонки для удобства
df = df.rename(columns={
    'Точка СКУД': 'СКУД',
    'Дата': 'Дата',
    'Время': 'Время',
    'Name': 'Сотрудник'
})

# Объединяем дату и время в один столбец datetime
# Предполагаем, что колонки 'Дата' и 'Время' имеют формат, который pandas корректно преобразует
df['Дата и время'] = pd.to_datetime(
    df['Дата'].astype(str) + ' ' + df['Время'].astype(str),
    dayfirst=True,     # если даты в формате ДД.ММ.ГГГГ
    errors='coerce'    # чтобы не падало при некорректных значениях
)

# Обрезаем лишние пробелы в URL
df['СКУД'] = df['СКУД'].astype(str).str.strip()

# Задаём списки точек входа и выхода
entry_points = ['http://10.100.6.65', 'http://10.100.6.79']
exit_points  = ['http://10.100.6.25', 'http://10.100.6.38']

# Определяем действие — Вход, Выход или Неизвестно
df['Действие'] = df['СКУД'].apply(
    lambda url: 'Вход' if url in entry_points
                else ('Выход' if url in exit_points else 'Неизвестно')
)

# Разделяем обратно на отдельные столбцы Дата и Время (они понадобятся для группировки)
df['Дата'] = df['Дата и время'].dt.date
df['Время'] = df['Дата и время'].dt.time

# Сортируем по сотруднику, дате и времени
df = df.sort_values(by=['Сотрудник', 'Дата и время']).reset_index(drop=True)


# --- 2. Находим время прихода и ухода (первые/последние события) ---

# Первый вход (минимум среди строк с Действие == 'Вход')
enter = (
    df[df['Действие'] == 'Вход']
    .groupby(['Сотрудник', 'Дата'], as_index=False)['Дата и время']
    .min()
)
enter = enter.rename(columns={'Дата и время': 'Время прихода'})
enter['Время прихода'] = enter['Время прихода'].dt.time

# Последний выход (максимум среди строк с Действие == 'Выход')
exit = (
    df[df['Действие'] == 'Выход']
    .groupby(['Сотрудник', 'Дата'], as_index=False)['Дата и время']
    .max()
)
exit = exit.rename(columns={'Дата и время': 'Время ухода'})
exit['Время ухода'] = exit['Время ухода'].dt.time

# Объединяем приход/уход обратно в "итоговый" DataFrame (по сотруднику и дате)
summary = (
    pd.merge(enter, exit, on=['Сотрудник', 'Дата'], how='outer')
    .sort_values(by=['Сотрудник', 'Дата'])
    .reset_index(drop=True)
)


# --- 3. Вспомогательные функции для вычисления отсутствия ---

def adjust_for_lunch(start_abs: datetime, end_abs: datetime) -> timedelta:
    """
    Вернёт длительность отсутствия в полном виде (end_abs - start_abs),
    но вычтет ту часть, что попадает в обеденный интервал 12:00–14:00 того же дня.
    """
    if end_abs <= start_abs:
        return timedelta(0)

    lunch_start = datetime.combine(start_abs.date(), time(12, 0))
    lunch_end   = datetime.combine(start_abs.date(), time(14, 0))

    full_interval = end_abs - start_abs

    # Находим перекрытие с обедом
    overlap_start = max(start_abs, lunch_start)
    overlap_end   = min(end_abs, lunch_end)
    lunch_overlap = timedelta(0)
    if overlap_end > overlap_start:
        lunch_overlap = overlap_end - overlap_start

    return full_interval - lunch_overlap


def compute_absence_within_workday(start_abs: datetime, end_abs: datetime, work_date: datetime.date) -> timedelta:
    """
    Из промежутка отсутствия [start_abs, end_abs] сначала берём
    пересечение с рабочим временем 09:00–18:00, а затем вычитаем
    обед (12:00–14:00). Возвращаем скорректированную длительность.
    """
    if end_abs <= start_abs:
        return timedelta(0)

    # Задаём границы рабочего дня
    work_start = datetime.combine(work_date, time(9, 0))
    work_end   = datetime.combine(work_date, time(18, 0))

    # Вычисляем пересечение с рабочим временем
    interval_start = max(start_abs, work_start)
    interval_end   = min(end_abs, work_end)

    if interval_end <= interval_start:
        # Отсутствие целиком вне рабочих часов
        return timedelta(0)

    # Теперь из этого пересечения вычтем обеденный отрезок
    return adjust_for_lunch(interval_start, interval_end)


# --- 4. Считаем общее время отсутствия (не на работе) в рабочие часы, исключая обед ---

absences = []

# Пробегаем по группам (Сотрудник, Дата)
for (employee, work_date), group in df.groupby(['Сотрудник', 'Дата']):
    # Берём только события «Вход» и «Выход», уже отсортированные по времени
    events = (
        group[group['Действие'].isin(['Вход', 'Выход'])]
        .sort_values('Дата и время')
    )
    total_absence = timedelta(0)

    # Пробегаем по каждой строке с событием «Выход»
    for idx, row in events.iterrows():
        if row['Действие'] == 'Выход':
            # Ищем следующую запись «Вход» в той же группе
            subsequent = events[
                (events['Дата и время'] > row['Дата и время']) &
                (events['Действие'] == 'Вход')
            ]
            if not subsequent.empty:
                next_entry = subsequent.iloc[0]
                start_abs = row['Дата и время']
                end_abs   = next_entry['Дата и время']

                # Вычисляем отсутствие внутри рабочей смены (9–18) и без обеда (12–14)
                adjusted = compute_absence_within_workday(start_abs, end_abs, work_date)
                total_absence += adjusted

    absences.append({
        'Сотрудник': employee,
        'Дата': work_date,
        'Время отсутствия': total_absence
    })

df_absence = pd.DataFrame(absences)


# --- 5. Сводим всё вместе для финального вывода ---

result = pd.merge(
    summary,
    df_absence,
    on=['Сотрудник', 'Дата'],
    how='left'
).sort_values(by=['Сотрудник', 'Дата']).reset_index(drop=True)

# Переводим timedelta в часы для наглядности
def format_timedelta(td: timedelta) -> str:
    """
    Преобразует timedelta в строку формата 'ЧЧ:ММ'.
    Например, 1 час 30 минут -> '01:30'
    """
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}"

result['Часы отсутствия'] = result[
    'Время отсутствия'
].apply(format_timedelta)

# Оставим только необходимые столбцы
final = result[[
    'Сотрудник',
    'Дата',
    'Время прихода',
    'Время ухода',
    'Часы отсутствия'
]]

# Настройка для полного вывода в консоль
pd.set_option('display.max_rows', None)
pd.set_option('display.width', None)

print(final)
