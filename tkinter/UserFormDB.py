import tkinter
from tkinter import ttk
from tkinter import messagebox
import sqlite3
import os
import openpyxl


def entrar_dados():
    aceitado = accept_var.get()

    if aceitado == "Aceitado":

        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            # Registration courses
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()

            print("Primeiro nome: ", firstname, "Último nome: ", lastname)
            print("Título: ", title, "Idade: ", age, "Nacionalidade: ", nationality)
            print("# Cursos: ", numcourses, "# Semestres: ", numsemesters)
            print("Status do registro", registration_status)
            print("---------------------------------------------------------------")

            # Local store file Excel/CSV
            filepath = r"C:\Users\fgsjd\OneDrive\Área de Trabalho\data.xlsx"

            # Verificar se o arquivo existe
            if os.path.exists(filepath):
                # Criar uma nova planilha
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["Primeiro nome", "Último nome", "Título", "Idade", "Nacionalidade", "# Cursos",
                           "# Semestres", "Status do registro"]
                sheet.append(heading)
                # Salvar o arquivo
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, nationality, numcourses, numsemesters, registration_status])
            workbook.save(filepath)
            # Create table
            conn = sqlite3.connect('data2.db')

            # Insert data
            data_insert_query = '''INSERT INTO Student_data (firstname, lastname, title, age, nationality,registration_status, 
                num_courses, num_semesters) VALUES (?,?,?,?,?,?,?,?)'''
            data_insert_tuple = (
                firstname, lastname, title, age, nationality, registration_status, numcourses, numsemesters)

            cursor = conn.cursor()
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()
            conn.close()
        else:
            tkinter.messagebox.showwarning(title="Erro", message="Primeiro e último nome são obrigatórios")
    else:
        tkinter.messagebox.showwarning(title="Error", message="Você precisa aceitar os termos e condições")


window = tkinter.Tk()
window.title("Formulário de entrada de dados")

frame = tkinter.Frame(window)
frame.pack()

# Saving user info
user_info_frame = tkinter.LabelFrame(frame, text="Informações do usuário")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="Primeiro nome")
first_name_label.grid(row=0, column=0)

last_name_label = tkinter.Label(user_info_frame, text="Último nome")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)

first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Título")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Idade")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nacionalidade")
nationality_combobox = ttk.Combobox(user_info_frame,
                                    values=["Africa", "Antártica", "Asia", "Europa", "Oceania", "América do Norte",
                                            "América do Sul"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Saving course info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Status do registro")
reg_status_var = tkinter.StringVar(value="Não registrado")

registered_check = tkinter.Checkbutton(courses_frame, text="Atualmente registrado",
                                       variable=reg_status_var, onvalue="Registrado", offvalue="Não registrado")
registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(courses_frame, text="# Cursos completados")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tkinter.Label(courses_frame, text="# Semestres")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Termos & condições")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Não aceito")
terms_check = tkinter.Checkbutton(terms_frame, text="Eu aceito os termos e condições",
                                  variable=accept_var, onvalue="Aceitado", offvalue="Não aceitado")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Entre com os dados", command=entrar_dados)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
