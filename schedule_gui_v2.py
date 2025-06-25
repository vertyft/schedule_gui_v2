import pandas as pd
from datetime import datetime, timedelta
import re, os
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext

DAYS = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
WEEK_RE = re.compile(r'(\d{2}\.\d{2}\.\d{2})\s*-\s*(\d{2}\.\d{2}\.\d{2})')

def base_subject(name: str) -> str:
    return re.split(r'[,(]', name, 1)[0].strip().lower()

def parse_schedule(file_path):
    df = pd.read_excel(file_path, header=None)
    schedule = []
    current_week = None
    current_day = None

    for _, row in df.iterrows():
        cell_6 = str(row[6])
        m = WEEK_RE.search(cell_6)
        if m:
            current_week = {
                'range': f"{m.group(1)}-{m.group(2)}",
                'start': datetime.strptime(m.group(1), '%d.%m.%y').date(),
                'end'  : datetime.strptime(m.group(2), '%d.%m.%y').date()
            }
            continue

        day_candidate = str(row[0]).strip()
        if day_candidate in DAYS:
            current_day = day_candidate

        if not (current_week and current_day):
            continue

        subject_full = str(row[5]).strip()
        if not subject_full or subject_full.lower() == 'название дисциплины':
            continue

        schedule.append({
            'date': current_week['start'] + timedelta(days=DAYS.index(current_day)),
            'time': str(row[2]).strip(),
            'room': str(row[4]).strip(),
            'subject': subject_full,
            'subject_id': base_subject(subject_full),
            'teacher': cell_6.strip(),
            'week': current_week['range']
        })

    return pd.DataFrame(schedule).drop_duplicates(
        subset=['date', 'time', 'room', 'subject', 'teacher']
    )

def search():
    file_path = file_path_var.get()
    mode = search_mode.get()
    keyword = keyword_entry.get().strip().lower()
    result_box.delete('1.0', tk.END)

    if not file_path or not keyword:
        messagebox.showerror("Ошибка", "Выберите файл и введите предмет или преподавателя.")
        return

    if not os.path.exists(file_path):
        messagebox.showerror("Ошибка", f"Файл {file_path} не найден.")
        return

    try:
        df_sched = parse_schedule(file_path)
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{e}")
        return

    if mode == 'subject':
        out = df_sched[df_sched['subject_id'] == keyword]
    else:
        mask = df_sched['teacher'].str.lower().str.contains(rf'\b{keyword}\b')
        out = df_sched[mask]

    if out.empty:
        result_box.insert(tk.END, "Ничего не найдено.")
        return

    output_lines = []
    group_name = os.path.splitext(os.path.basename(file_path))[0]

    output_lines.append(f"👥 Группа: {group_name}")
    if mode == 'subject':
        output_lines.append(f"📚 Предмет: {out.iloc[0]['subject']}")
    else:
        output_lines.append(f"👨‍🏫 Преподаватель: {out.iloc[0]['teacher']}")
    output_lines.append("📅 Даты занятий:")

    for i, row in enumerate(out.sort_values(['date', 'time']).itertuples(), 1):
        output_lines.append(f"{i:2}. {row.date:%d.%m.%Y} ({row.time}, {row.room}, {row.teacher})")

    output_lines.append(f"\n🔢 Всего занятий     : {len(out)}")
    output_lines.append(f"🗓️  Задействовано недель: {out['week'].nunique()}")

    result_text = "\n".join(output_lines)
    result_box.insert(tk.END, result_text)

    txt_name = os.path.splitext(file_path)[0] + '_результат.txt'
    with open(txt_name, 'w', encoding='utf-8') as f:
        f.write(result_text)
    messagebox.showinfo("Готово", f"Результат сохранен в:\n{txt_name}")

def choose_file():
    path = filedialog.askopenfilename(
        title="Выберите Excel-файл",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        file_path_var.set(path)

# ---------------- GUI ----------------
root = tk.Tk()
root.title("Расписание занятий")
root.geometry("700x600")

messagebox.showinfo("Расписание занятий", "Автор программы за результат ответственности не несет,\nперепроверяйте даты самостоятельно.")

tk.Button(root, text="Выбрать файл", command=choose_file).pack(pady=5)
file_path_var = tk.StringVar()
tk.Entry(root, textvariable=file_path_var, width=70).pack(pady=5)

search_mode = tk.StringVar(value='subject')
tk.Radiobutton(root, text="По предмету", variable=search_mode, value='subject').pack(anchor='w')
tk.Radiobutton(root, text="По преподавателю", variable=search_mode, value='teacher').pack(anchor='w')

tk.Label(root, text="Введите название предмета или фамилию преподавателя:").pack()
keyword_entry = tk.Entry(root, width=50)
keyword_entry.pack(pady=5)

tk.Button(root, text="Найти", command=search).pack(pady=10)

result_box = scrolledtext.ScrolledText(root, width=80, height=25)
result_box.pack()

root.mainloop()
