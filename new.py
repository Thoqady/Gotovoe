import tkinter as tk
from tkinter import filedialog
import xlsxwriter
import qrcode

def create_excel_A7(entries, save_path):
    # Создаем новый Excel-документ с xlsxwriter
    workbook = xlsxwriter.Workbook(save_path)
    worksheet = workbook.add_worksheet()

    # Устанавливаем размеры ячеек внутри Excel
    for i, entry in enumerate(entries):
        # Получаем тексты из полей ввода
        info, text_A2, text_A1 = entry

        # Вычисляем номера строк в зависимости от индекса
        row_A = i * 2
        row_B = i * 2 + 1

        # Устанавливаем размеры ячеек
        worksheet.set_row(row_A, 153)
        worksheet.set_row(row_B, 45.3)

        # Получаем объект format для настройки стилей ячеек
        cell_format_text = workbook.add_format({
            'font_size': 12,   # размер шрифта
            'align': 'center',  # выравнивание по центру
            'valign': 'vcenter',  # выравнивание по вертикали по центру
        })
        # Получаем объект format для настройки стилей ячеек
        cell_format_text2 = workbook.add_format({
            'font_size': 76,  # размер шрифта
            'align': 'center',  # выравнивание по центру
            'valign': 'vcenter',  # выравнивание по вертикали по центру
        })

        # Устанавливаем жирный шрифт
        cell_format_text2.set_bold()

        # Объединяем ячейки A2 и B2
        worksheet.merge_range(f'A{row_A + 2}:B{row_A + 2}', text_A2, cell_format_text)

        # Вставляем текст из ячеек A1 и A2
        worksheet.write(f'A{row_A + 1}', text_A2, cell_format_text2)
        worksheet.write(f'A{row_B + 1}', text_A1, cell_format_text)


        # Создаем объект QR-кода
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,  # Устанавливаем размер каждого квадрата в QR-коде
            border=5,
        )
        qr.add_data(info)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")

        # Переводим размеры QR-кода в пиксели (примерно 1 см = 37.795276 пикселя)
        qr_width_pixels = int(9 * 37.795276)
        qr_height_pixels = int(9 * 37.795276)

        # Сохраняем QR-код как изображение
        qr_img = qr_img.resize((qr_width_pixels, qr_height_pixels))
        qr_img.save(f'qr_code_{i+1}.png')

        # Вставляем изображение QR-кода в ячейку B1 посередине
        worksheet.insert_image(f'B{row_A + 1}', f'qr_code_{i+1}.png', {'x_offset': 5, 'y_offset': 15, 'x_scale': 0.5, 'y_scale': 0.5})

    # Устанавливаем ширину столбцов
    worksheet.set_column('A:B', 24.57)

    # Закрываем Excel-документ
    workbook.close()

def add_entry(entries, canvas, interior):
    # Получаем текущее количество строк
    current_row = len(entries)

    # Создаем метки и поля ввода для каждого поля
    entry_label = tk.Label(interior, text="Инфо - ")
    entry_label.grid(row=current_row, column=0, sticky="e", padx=5, pady=5)

    entry = tk.Entry(interior)
    entry.grid(row=current_row, column=1, sticky="w", padx=5, pady=5)

    entry_label_A1 = tk.Label(interior, text="Номер стойки - ")
    entry_label_A1.grid(row=current_row, column=2, sticky="e", padx=5, pady=5)

    entry_A1 = tk.Entry(interior)
    entry_A1.grid(row=current_row, column=3, sticky="w", padx=5, pady=5)

    entry_label_A2 = tk.Label(interior, text="Отв. - ")
    entry_label_A2.grid(row=current_row, column=4, sticky="e", padx=5, pady=5)

    entry_A2 = tk.Entry(interior)
    entry_A2.grid(row=current_row, column=5, sticky="w", padx=5, pady=5)

    entries.append((entry, entry_A1, entry_A2))

    # Обновляем область прокрутки
    canvas.update_idletasks()
    canvas.config(scrollregion=canvas.bbox("all"))

def save_file(entries):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        create_excel_A7([(e[0].get(), e[1].get(), e[2].get()) for e in entries], save_path)

# Создаем основное окно Tkinter
root = tk.Tk()
root.title("Создание Qr-code")

# Устанавливаем начальные размеры окна
root.geometry("630x500")

# Создаем Frame для размещения виджетов
frame = tk.Frame(root)
frame.pack(expand=True, fill="both")

# Создаем Canvas для обеспечения прокрутки
canvas = tk.Canvas(frame)
canvas.pack(side="left", fill="both", expand=True)

# Добавляем полосу прокрутки
scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")

# Привязываем прокрутку к Canvas
canvas.configure(yscrollcommand=scrollbar.set)

# Создаем внутренний фрейм для размещения меток и полей ввода
interior = tk.Frame(canvas)
interior_id = canvas.create_window((0, 0), window=interior, anchor="nw")

# Добавляем функцию прокрутки к Canvas
def _on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")
canvas.bind_all("<MouseWheel>", _on_mousewheel)

# Создаем кнопку для добавления новых строк
button_add = tk.Button(root, text="Добавить строку", command=lambda: add_entry(entries, canvas, interior))
button_add.pack(side="bottom", pady=10)

# Создаем кнопку для сохранения
button_save = tk.Button(root, text="Сохранить", command=lambda: save_file(entries))
button_save.pack(side="bottom", pady=10)

# Создаем список для хранения объектов меток и полей ввода
entries = []

# Запускаем главный цикл Tkinter
root.mainloop()
