#!/usr/bin/env python3
import os
import logging
import sys
import random

from tkinter import messagebox
from tkinter import *
import tkinter as tk
from tkinter import ttk
from pptx import Presentation


class CertificateCreator:
    def __init__(self, windows):
        self.window = windows
        self.window.title('Создание сертификата')
        self.window.geometry('450x300')

        self.name = tk.StringVar()
        self.price = tk.StringVar()
        self.who_buy = tk.StringVar()

        self.setup_ui()
        self.setup_logging()

    @staticmethod
    def setup_logging():
        logging.basicConfig(
            level=logging.INFO,
            filename='program.txt',
            filemode='a',
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%d.%m.%Y - %H:%M'
        )
        logger = logging.getLogger()
        stream_handler = logging.StreamHandler(sys.stdout)
        stream_handler.setLevel(logging.INFO)
        stream_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(stream_handler)

    def setup_ui(self):
        frame = Frame(self.window)
        frame.pack()

        ttk.Label(frame, text='ФИО').pack(padx=100, pady=5)
        ttk.Entry(frame, textvariable=self.name, width=30).pack(padx=100, pady=5)

        ttk.Label(frame, text='СУММА').pack(padx=100, pady=5)
        ttk.Entry(frame, textvariable=self.price, width=30).pack(padx=100, pady=5)

        ttk.Label(frame, text='КОМУ ПРОДАНО').pack(padx=100, pady=5)
        ttk.Entry(frame, textvariable=self.who_buy, width=30).pack(padx=100, pady=5)

        btn_create = tk.Button(
            self.window, text='Создать сертификат',
            command=self.replace_text_in_presentation)
        btn_create.pack(pady=5)

        exit_button = tk.Button(self.window, text="Выход", command=self.window.destroy)
        exit_button.pack(pady=5)

    @staticmethod
    def get_random_number():
        """Получить случайное число."""
        return random.randint(100000, 999999)

    @staticmethod
    def check_input_file(input_file_path):
        """Проверяет существование файлов в папке с программой."""
        if not os.path.exists(input_file_path):
            messagebox.showinfo(
                message=f'Файл «{input_file_path}» не найден в папке с программой.'
            )
            return False
        return True

    @staticmethod
    def check_output_file(output_file_path):
        return os.path.exists(output_file_path)

    def replace_text_in_presentation(self):
        input_file_path = 'Сертификат.pptx'
        output_file_path = 'Сертификат на печать.pptx'

        if not self.check_input_file(input_file_path):
            return

        price_value = self.price.get()
        name_value = self.name.get()
        who_buy_value = self.who_buy.get()

        try:
            price_value = float(price_value)
        except ValueError:
            messagebox.showerror("Ошибка", "Цена должна быть числом")
            logging.error("Цена должна быть числом")
            return

        if price_value > 100000:
            messagebox.showerror("Ошибка", "Цена слишком большая (максимум 100000)")
            logging.error("Цена слишком большая (максимум 100000)")
            return

        if price_value <= 0:
            messagebox.showerror("Ошибка", "Цена должна быть больше 0")
            logging.error("Цена должна быть больше 0")
            return

        prs = Presentation(input_file_path)

        replacements = {
            'price': str(price_value),
            'name': name_value,
            'serial': str(self.get_random_number()),
        }

        slide = prs.slides[0]
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for key, value in replacements.items():
                        run.text = run.text.replace(key, value)

        prs.save(output_file_path)

        message = (
            f'Сертификат №: {replacements["serial"]}; '
            f'Именной: {replacements["name"]}; '
            f'Кому: {who_buy_value}; '
            f'Цена: {replacements["price"]} ₽'
        )

        if self.check_output_file(output_file_path):
            messagebox.showinfo(message=message)
            logging.info(message)
        self.window.destroy()

    def main(self):
        self.window.mainloop()


if __name__ == '__main__':
    window = Tk()
    app = CertificateCreator(window)
    app.main()
