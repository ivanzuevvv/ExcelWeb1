import tkinter as tk
import random

def change_bg_color():
    random_color = "#{:06x}".format(random.randint(0, 0xFFFFFF))  # Генерируем случайный цвет
    root.configure(bg=random_color)  # Применяем новый цвет к фону окна

root = tk.Tk()
root.title("Изменение цвета фона")

change_color_button = tk.Button(root, text="Изменить цвет фона", command=change_bg_color)
change_color_button.pack(pady=20)

root.mainloop()

