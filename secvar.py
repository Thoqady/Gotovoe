import tkinter as tk
from tkinter import filedialog
import xlsxwriter


def create_excel_A6(entries):
    # Получаем тексты из полей ввода
    row, place, type_subsystem, mnemonic_ne, project, department, responsible, leader, shift_contact, shift_contact2, shift_contact3, opl_group = [entry.get() for entry in entries]

    # Открываем диалоговое окно для выбора пути сохранения файла
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if save_path:
        # Создаем новый Excel-документ с xlsxwriter
        workbook = xlsxwriter.Workbook(save_path)
        worksheet = workbook.add_worksheet()
        worksheet.set_landscape()

        # Устанавливаем ширину столбцов
        worksheet.set_column('A:A', 17.33)
        worksheet.set_column('B:B', 33.78)

        # Устанавливаем высоту столбцов
        worksheet.set_row(0, 27)
        worksheet.set_row(1, 22.2)
        worksheet.set_row(2, 21.6)
        worksheet.set_row(3, 22.8)
        worksheet.set_row(4, 25.8)
        worksheet.set_row(5, 19.8)
        worksheet.set_row(6, 19.8)
        worksheet.set_row(7, 19.8)
        worksheet.set_row(8, 20.4)

        # Получаем объект format для настройки стилей ячеек
        cell_format_text = workbook.add_format({
            'font_size': 18,   # размер шрифта
            'align': 'center',  # выравнивание по центру
            'valign': 'vcenter',
            'border': 1,
            'font_name': 'Times New Roman',
        })

        cell_format_text.set_bold()

        cell_format_A = workbook.add_format({
            'font_size': 11,   # размер шрифта
            'valign': 'vcenter',  # выравнивание по вертикали сверху
            'align': 'left',
            'font_name': 'Times New Roman',
            'border': 1,
        })

        cell_format_B = workbook.add_format({
            'font_size': 11,   # размер шрифта
            'valign': 'vcenter',  # выравнивание по вертикали сверху
            'align': 'center',
            'font_name': 'Times New Roman',
            'border': 1,
        })

        cell_format_B .set_text_wrap()

        # Вставляем текст
        worksheet.write('A1', f"Ряд {row}", cell_format_text)
        worksheet.write('B1', f"Место {place}", cell_format_text)
        worksheet.write('A2', f"Тип/подсистема", cell_format_A)
        worksheet.write('B2', f"{type_subsystem}", cell_format_B)
        worksheet.write('A3', f"Мнемоника / NE", cell_format_A)
        worksheet.write('B3', f"{mnemonic_ne}", cell_format_B)
        worksheet.write('A4', f"Проект", cell_format_A)
        worksheet.write('B4', f"{project}", cell_format_B)
        worksheet.write('A5', f"Подразделение", cell_format_A)
        worksheet.write('B5', f"{department}", cell_format_B)
        worksheet.write('A6', f"Отв.лицо", cell_format_A)
        worksheet.write('B6', f"{responsible} тел. {leader}", cell_format_B)
        worksheet.write('A7', f"Руководитель", cell_format_A)
        worksheet.write('B7', f"{shift_contact} тел. {shift_contact2}", cell_format_B)
        worksheet.write('A8', f"Контакт деж. Смены", cell_format_A)
        worksheet.write('B8', f"{shift_contact3}", cell_format_B)
        worksheet.write('A9', f"Группа OPL", cell_format_A)
        worksheet.write('B9', f"{opl_group}", cell_format_B)

        # Закрываем Excel-документ
        workbook.close()


# Создаем основное окно Tkinter
root = tk.Tk()
root.title("Создание Qr-code")

# Устанавливаем начальные размеры окна
root.geometry("400x550")

# Создаем Frame для размещения виджетов
frame = tk.Frame(root)
frame.pack(expand=True, fill="both")

# Создаем метки и поля ввода для каждого поля
labels = ["Ряд - ", "Место - ", "Тип/подсистема - ", "Мнемоника / NE - ", "Проект - ",
          "Подразделение - ", "Отв.лицо - ", 'тел. ', "Руководитель - ", 'тел.',"Контакт деж. Смены - ", "Группа OPL - "]

entries = []

for i, label_text in enumerate(labels):
    entry_label = tk.Label(frame, text=label_text)
    entry_label.grid(row=i, column=0, sticky="e", padx=5, pady=5)

    entry = tk.Entry(frame)
    entry.grid(row=i, column=1, sticky="w", padx=5, pady=5)
    entries.append(entry)

button_A7 = tk.Button(frame, text="Формат 100 на 70 мм", command=lambda: create_excel_A6(entries))
button_A7.grid(row=len(labels) + 3, column=0, columnspan=2, pady=20)
button_A7.place(relx=0.32, rely=0.8)
# Запускаем главный цикл Tkinter
root.mainloop()