from main2 import *

# libraries Import
from tkinter import Tk, filedialog, messagebox
import customtkinter
import subprocess
import os

# Main Window Properties

window = Tk()
window.title("Время отсутствие сотрудника на рабочее время")
window.geometry("720x480")
window.configure(bg="#a2a7b3")


# Функция выбора папки для Entry_id1
def browse_input_file():
    file_selected = filedialog.askopenfile(filetypes =[('SKUD otchet', '*.xls *.xlsx')])
    if file_selected:
        Entry_id1.delete(0, 'end')
        Entry_id1.insert(0, file_selected.name)
        print(file_selected.name)

# Функция выбора папки для Entry_id6
def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        Entry_id6.delete(0, 'end')
        Entry_id6.insert(0, folder_selected)


# Обработка изменения чекбокса
def toggle_checkbox():
    if Checkbox_id8.get() == 1:
        Entry_id6.configure(state="disabled")
        Button_id7.configure(state="disabled")
    else:
        Entry_id6.configure(state="normal")
        Button_id7.configure(state="normal")

def toggle_checkbox_in_file():
    if Checkbox_in_file.get() == 1:
        Entry_id6.configure(state="disabled")
        Button_id7.configure(state="disabled")
        Checkbox_id8.configure(state="disabled")
    else:
        Entry_id6.configure(state="normal")
        Button_id7.configure(state="normal")
        Checkbox_id8.configure(state="normal")




def start():
    imput = Entry_id1.get()
    input_folder = '/'.join(imput.split("/")[:-1])
    input_file = '/'.join(imput.split("/")[-1:])
    output_folder = Entry_id6.get()
    in_same_folder = Checkbox_id8.get() == 1
    in_same_file = Checkbox_in_file.get() == 1
    exlporer = Checkbox_explorer.get() == 1
    xlsx = Checkbox_xlsx.get() == 1
    output_file = 'SKUD_otchet.xlsx'
    output_path = output_folder + '/' + output_file

    if not imput:
        messagebox.showerror("Xatolik", "Birinchi popka ko'rsatilmagan.")
        return
    if not os.path.isfile(imput):
        messagebox.showerror("Xatolik", "Birinchi popka noto'g'ri yoki mavjud emas.")
        return

    if not in_same_folder and not in_same_file:
        if not output_path:
            messagebox.showerror("Xatolik", "Ikkinchi popka ko'rsatilmagan. Yoki galochka qo'ying.")
            return
        if not os.path.isdir(output_path):
            messagebox.showerror("Xatolik", "Ikkinchi popka noto'g'ri yoki mavjud emas.")
            return
    else:
        output_path = input_folder + '/' + output_file
    if in_same_file:
        if input_file[-3:].lower() == 'xls':
            messagebox.showerror("Xatolik", "Hozircha xls formatga list qoshish imkonsiz.")
            return
        else:
            output_path = imput
    try:
        if not output_path == imput:
            os.remove(output_path)
    except FileNotFoundError:
        print("❗ Файл не найден — удалять нечего. Продолжим")
    except PermissionError:
        print("❗ Файл занят (возможно, открыт в Excel). Закрой файл и попробуй снова.")
        messagebox.showerror("Error", "❗ Файл занят (возможно, открыт в Excel). Закрой файл и попробуй снова.")
        exit()
    except Exception as e:
        print(f"❗ Неизвестная ошибка при удалении: {e}")
        messagebox.showerror("Error", f"❗ Неизвестная ошибка при удалении выходного файла:\n {e}")
        exit()
    work_start, work_end, lunch_start, lunch_end = start_hour.get(), end_hour.get(), start_lunch_hour.get(),end_lunch_hour.get()
    try:
        working_time(imput, output_path, in_same_file, int(work_start), int(work_end), int(lunch_start), int(lunch_end),
                     Entry_enter_ips.get(), Entry_exit_ips.get())
    except Exception as e:
        messagebox.showerror("Error", f"❗ Ошибка !!! \n{e}")
        print(e)
        return
    messagebox.showinfo("Success", "Fayl saqlandi")
    # Открыть проводник и выделить файл
    if exlporer: subprocess.run(f'explorer /select,"{output_path.replace('/', '\\')}"')
    if xlsx: os.startfile(output_path)



Entry_id1 = customtkinter.CTkEntry(
    master=window,
    placeholder_text="fayl",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=550,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    )
Entry_id1.place(x=10, y=50)
Button_id3 = customtkinter.CTkButton(
    master=window,
    text="Obzor",
    font=("undefined", 16),
    text_color="#000000",
    hover=True,
    hover_color="#b7b3b3",
    height=30,
    width=95,
    border_width=2,
    corner_radius=10,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    command=browse_input_file,
    )
Button_id3.place(x=580, y=50)
Label_id5 = customtkinter.CTkLabel(
    master=window,
    text="Fayl saqlash uchun popkani ko'rsating",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=250,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    )
Label_id5.place(x=10, y=100)
Button_id7 = customtkinter.CTkButton(
    master=window,
    text="Obzor",
    font=("undefined", 16),
    text_color="#000000",
    hover=True,
    hover_color="#a8a4a4",
    height=30,
    width=95,
    border_width=2,
    corner_radius=10,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    command=browse_output_folder,
    )
Button_id7.place(x=580, y=140)
Label_id2 = customtkinter.CTkLabel(
    master=window,
    text="Faylni ko'rsating",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=100,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    )
Label_id2.place(x=10, y=10)
Button_id4 = customtkinter.CTkButton(
    master=window,
    text="Start",
    font=("undefined", 26),
    text_color="#000000",
    hover=True,
    hover_color="#b2aeae",
    height=50,
    width=150,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    command=start,
    )
Button_id4.place(x=550, y=410)

Checkbox_id8 = customtkinter.CTkCheckBox(
    master=window,
    text="Fayllar turgan popkaga",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
command=toggle_checkbox,
    )
Checkbox_id8.place(x=300, y=100)


Checkbox_in_file = customtkinter.CTkCheckBox(
    master=window,
    text="Kirish fayliga list holida",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
command=toggle_checkbox_in_file,
    )
Checkbox_in_file.place(x=470, y=100)





Label_enter_ip = customtkinter.CTkLabel(
    master=window,
    text="Kirish turniketlarini IP adreslarini yozing:",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=100,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    )
Label_enter_ip.place(x=10, y=190)

Entry_enter_ips = customtkinter.CTkEntry(
    master=window,
    placeholder_text="kirish ip lari (10.10.10.10, 10.10.10.20)",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=550,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    )
Entry_enter_ips.place(x=10, y=220)
Entry_enter_ips.insert(0, '10.100.6.65, 10.100.6.79')

Label_exit_ip = customtkinter.CTkLabel(
    master=window,
    text="Chiqish turniketlarini IP adreslarini yozing:",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=100,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    )

Label_exit_ip.place(x=10, y=260)
Entry_exit_ips = customtkinter.CTkEntry(
    master=window,
    placeholder_text="chiqish ip lari (10.10.10.10, 10.10.10.20)",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=550,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    )
Entry_exit_ips.place(x=10, y=290)
Entry_exit_ips.insert(0,'10.100.6.25, 10.100.6.38')

def on_validate_input(p):
    # Проверяем, является ли введённое значение целым числом
    if p == "" or p.isdigit() and int(p)<24:
        return True
    return False


validate_input = window.register(on_validate_input)


customtkinter.CTkLabel(
    master=window,
    text="Ish vaqti boshlanishi",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=50,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    ).place(x=10, y=330)



# Создаем поле ввода для начала времени
start_hour = customtkinter.CTkEntry(
    master=window,
    placeholder_text="",
    font=("Arial", 14),
    width=40,
    height=30,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    validate="key",
    validatecommand=(validate_input, "%P")
)
start_hour.place(x=160, y=330)
start_hour.insert(0, 9)


customtkinter.CTkLabel(
    master=window,
    text="tugashi",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=50,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    ).place(x=230, y=330)


# Создаем поле ввода для конца времени
end_hour = customtkinter.CTkEntry(
    master=window,
    placeholder_text="",
    font=("Arial", 14),
    width=40,
    height=30,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    validate="key",
    validatecommand=(validate_input, "%P")
)
end_hour.place(x=300, y=330)
end_hour.insert(0, 18)




customtkinter.CTkLabel(
    master=window,
    text="Abet vaqti boshlanishi",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=50,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    ).place(x=10, y=370)


# Создаем поле ввода для начала времени
start_lunch_hour = customtkinter.CTkEntry(
    master=window,
    placeholder_text="",
    font=("Arial", 14),
    width=40,
    height=30,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    validate="key",
    validatecommand=(validate_input, "%P")
)
start_lunch_hour.place(x=160, y=370)
start_lunch_hour.insert(0, 12)


customtkinter.CTkLabel(
    master=window,
    text="tugashi",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=50,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    ).place(x=230, y=370)


# Создаем поле ввода для конца времени
end_lunch_hour = customtkinter.CTkEntry(
    master=window,
    placeholder_text="",
    font=("Arial", 14),
    width=40,
    height=30,
    corner_radius=10,
    bg_color="#a2a7b3",
    fg_color="#a2a7b3",
    validate="key",
    validatecommand=(validate_input, "%P")
)
end_lunch_hour.place(x=300, y=370)
end_lunch_hour.insert(0, 14)









Checkbox_explorer = customtkinter.CTkCheckBox(
    master=window,
    text="Faylni provodnikda ko'rsatish",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
    )
Checkbox_explorer.place(x=110, y=430)




Checkbox_xlsx = customtkinter.CTkCheckBox(
    master=window,
    text="Faylni ochish (Excel)",
    text_color="#000000",
    border_color="#000000",
    fg_color="#808080",
    hover_color="#808080",
    corner_radius=4,
    border_width=2,
    )
Checkbox_xlsx.place(x=310, y=430)


Entry_id6 = customtkinter.CTkEntry(
    master=window,
    placeholder_text="popka",
    placeholder_text_color="#454545",
    font=("Arial", 14),
    text_color="#000000",
    height=30,
    width=550,
    border_width=2,
    corner_radius=6,
    border_color="#000000",
    bg_color="#a2a7b3",
    fg_color="#F0F0F0",
    )
Entry_id6.place(x=10, y=140)

#run the main loop
window.mainloop()

