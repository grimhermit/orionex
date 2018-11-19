# -- coding: utf-8 --
from __future__ import unicode_literals
import csv
import tkinter
import tkinter.filedialog as tf
import tkinter.messagebox as tm
import tkinter.scrolledtext as ts
import pickle
import urllib.request as ur
import urllib.error as ue
import os
import xlrd


class CsvParser:

    def __init__(self):
        # Список ключей словаря из прочитаного файла, предназначенные для удаления как лишние
        self.keys_to_kill = ['export_version', 'current_version', 'timestamp', 'product_class', 'product_name_DE',
                             'title_DE', 'title_EN', 'title_IT', 'product_keyword', 'special_price_flag',
                             'recommended_selling_price', 'full_text_DE', 'full_text_IT',
                             'category_path', 'category', 'group_product', 'single_image', 'tax_id', 'ean_code',
                             'battery_supply', 'battery_combination', 'product_image_2', 'product_image_3', 'product_image_4',
                             'material_description', 'material_thickness', 'total_length', 'diameter', 'weight', 'length',
                             'width', 'height', 'contents', 'availability', 'delivery_week', 'dress_size', 'shoe_size',
                             'main_color', 'color_description', 'battery_type', 'packaging_size', 'aroma', 'isbn', 'number_of_pages',
                             'total_duration', 'hardcore_flag', 'age_rating', 'product_languages', 'country_of_origin_code',
                             'series_name', 'reference_quantity', 'reference_price_factor', 'promotional_packing', 'product_icons',
                             'selling_unit', 'voc_percentage', 'barcode', 'barcode_type', 'food_information_flag',
                             'detailed_full_text_de', 'detailed_full_text_en', 'group_id', 'tariff_number']
        # Эталонные ключи
        self.etha_keys = ['name : Название',	'vendor : Производитель', 'supplier : Поставщик', 'image : Иллюстрация',
                          'pre_order : Предзаказ', 'article : Артикул', 'code_1c : 1C', 'folder : Категория', 'tags : Теги',
                          'hidden : Скрыто', 'kind_id : ID', 'is_kind : Модификация', 'discounted : Товар уже со скидкой',
                          'note : Анонс товара', 'body : Описание', 'amount : Количество', 'unit : Единица измерения',
                          'weight : Вес', 'weight_unit : Единица веса', 'dimensions : Габариты',	'new : Новинка',
                          'special : Спецпредложение', 'yml : Yandex.Market', 'price : Цена', 'price_old : Старая цена',
                          'price2 : Цена 2', 'price3 : Цена 3', 'currency : Валюта', 'seo_noindex : Индексация',
                          'seo_h1 : Заголовок (H1)', 'seo_title : Title', 'seo_description : Description',
                          'seo_keywords : Keywords', 'sef_url : ЧПУ', 'uuid : UUID Товара', 'uuid_mod : UUID Модификации']
        # Список названий ключей, значения которых будут записаны .
        self.replacements = {'product_name_EN': 'name : Название', 'label_name': 'vendor : Производитель',
                             'product_image_1': 'image : Иллюстрация', 'product_id': 'article : Артикул',
                             'detailed_full_text_en': 'body : Описание', 'novelty_flag': 'new : Новинка',
                             'product_price': 'price : Цена'}

        # Количество хранимых в бинарниках данных увеличилось, было бы здорово хранить их в отдельном каталоге
        self.osdir = os.getcwd()
        self.data_dir_name = "Data"
        self.input_dir_name = "Input"
        # Логический атрибут, значение которого будет менятся в зависимости от наличия файла соответсвия кодов групп
        # Это наш триггер
        self.xlsx_bool = False
        self.xlsx_list = []
        self.xlsx_id = set()
        # пользователь составил список категорий, с которыми обычно работает, в виду данного обстоятельства
        # реорганизуем процесс отбора категорий
        # атрибут хранит ссылку
        self.hyperlink = str('')
        self.csv_path = str('')
        self.csvfile_name = str('result.csv')
        self.convertation_cource = float(0)
        self.output = str(self.osdir + '/')
        self.out = ""
        self.cats = []
        self.sorted_cats = []
        self.cats_id = []
        self.sorted_cats_id = []
        self.cats_to_die = []
        self.rows = []
        self.buffer = []
        self.buffer2 = []
        self.buffer3 = []
        self.outlines = []
        self.root = tkinter.Tk()
        self.root.title("Orion export v. 2.0")
        self.root.geometry = "600x600"
        # Интерфейс выглядит плохо, применим несколько фреймов для организации виджетов внутри окна

        self.alpha_frame = tkinter.Frame(self.root)
        self.omega_frame = tkinter.Frame(self.root)
        self.frame1 = tkinter.LabelFrame(self.alpha_frame, text="1. Загрузка файла с сайта orion.de:")
        self.frame2 = tkinter.LabelFrame(self.alpha_frame, text="3. Отбор лишних категорий и курс конвертации цен:")
        self.frame3 = tkinter.LabelFrame(self.alpha_frame, text="4. Работа с наценками:")
        self.frame4 = tkinter.LabelFrame(self.alpha_frame, text="5. Завершение работы:")
        # frame 5 резервируем для вывода справки
        self.frame5 = tkinter.LabelFrame(self.root, text="Инструкция:")
        # frame 6 эксперимент с адаптацией интерфейса (не удался, реализую позже, возможно)
        self.frame6 = tkinter.LabelFrame(self.alpha_frame, text="2. Обработка загруженного ранее файла:")
        # Атрибуты ГУИ
        self.label_download_stat = tkinter.Label(self.frame1, text="Статус загрузки файла.")
        self.button_change_hyperlink = tkinter.Button(self.frame1, text="Обновить ссылку.",
                                                      command=self.set_hyperlink)
        self.button_download = tkinter.Button(self.frame1, text="Скачать.", command=self.download_file)

        self.button1 = tkinter.Button(self.frame4, text="Обработать", command=self.buffer_handler)
        self.button2 = tkinter.Button(self.frame2, text="Изменить курс рубля", command=self.change_con_course)
        self.button3 = tkinter.Button(self.frame6, text="Прочитать файл", command=lambda: self.reader(self.csv_path))
        self.label_input = tkinter.Label(self.frame6, text="Путь к исходному файлу выгрузки:")
        self.input_entry = tkinter.Label(self.frame6, text=self.output)
        self.label_output = tkinter.Label(self.frame6, text="Указать путь к каталогу, куда будет помещен результат:")
        self.output_entry = tkinter.Label(self.frame6, text=self.output)
        self.label_cats_to_die = tkinter.Label(self.frame2, text="Категории товаров для удаления:")
        self.cats_to_die_dialog = tkinter.Button(self.frame2, text="Выбрать.", command=self.cats_catcher)
        self.quit_button = tkinter.Button(self.frame4, text="Выход", command=self.root.destroy)
        self.set_input_path = tkinter.Button(self.frame6, text="...", command=self.open_file)
        self.set_output_path = tkinter.Button(self.frame6, text="...", command=self.path_to_save)
        self.label_con_course = tkinter.Label(self.frame2, text=("Курс RUB к EUR:" + str(self.convertation_cource)))
        self.final_msg = None
        # добавляем атрибуты класса
        # Атрибут для харенения значения наценки 1
        self.price_markup1 = float(0)
        # Атрибут для хранения значения наценки 2
        self.price_markup2 = float(0)
        self.price_markup3 = float(0)
        self.price_markup4 = float(0)
        self.price_markup5 = float(0)
        self.price_markup6 = float(0)
        self.price_markup7 = float(0)
        self.price_markup8 = float(0)
        self.price_markup9 = float(0)
        self.price_markup10 = float(0)
        # Общий список категорий, к которым будет применяться наценка.
        self.chosen_cats_markup = []
        # Атрибуты...
        self.chosen_cats_markup1_group = []
        self.chosen_cats_markup2_group = []
        self.chosen_cats_markup3_group = []
        self.chosen_cats_markup4_group = []
        self.chosen_cats_markup5_group = []
        self.chosen_cats_markup6_group = []
        self.chosen_cats_markup7_group = []
        self.chosen_cats_markup8_group = []
        self.chosen_cats_markup9_group = []
        self.chosen_cats_markup10_group = []
        self.chosen_cats_markup_general_group = set()
        # Список для первоначальной фильтрации выводимых категорий товаров для наценки
        # Это все категории из sorted_cats минус cats_to_die и минус категории, к которым уже
        # Будет применятся наценка, для чего и нужен отдельный список, общий для всех наценяемых категорий
        # Изменил на множество со списка
        self.invited_cats_on_markup = set()
        # добавляем элементы графического интерфейса
        # self.label_price_markup = tkinter.Label(self.frame3, text="Работа с наценками:")
        self.label_set_price_markup1 = tkinter.Label(self.frame3, text=("Наценка 1 = " + str(self.price_markup1) + "%"))
        self.label_choose_cats_to_markup1 = tkinter.Label(self.frame3, text="Категории товаров для наценки 1:")
        self.button_price_markup1 = tkinter.Button(self.frame3,
                                                   text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup1_group))
        self.button_set_price_markup1 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup1,
                                                                                                  self.label_set_price_markup1))

        self.label_set_price_markup2 = tkinter.Label(self.frame3, text=("Наценка 2 = " + str(self.price_markup2) + "%"))
        self.label_choose_cats_to_markup2 = tkinter.Label(self.frame3, text="Категории товаров для наценки 2:")
        self.button_price_markup2 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup2_group))
        self.button_set_price_markup2 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup2,
                                                                                                  self.label_set_price_markup2))

        self.label_set_price_markup3 = tkinter.Label(self.frame3, text=("Наценка 3 = " + str(self.price_markup3) + "%"))
        self.label_choose_cats_to_markup3 = tkinter.Label(self.frame3, text="Категории товаров для наценки 3:")
        self.button_price_markup3 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup3_group))
        self.button_set_price_markup3 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup3,
                                                                                                  self.label_set_price_markup3))

        self.label_set_price_markup4 = tkinter.Label(self.frame3, text=("Наценка 4 = " + str(self.price_markup4) + "%"))
        self.label_choose_cats_to_markup4 = tkinter.Label(self.frame3, text="Категории товаров для наценки 4:")
        self.button_price_markup4 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup4_group))
        self.button_set_price_markup4 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup4,
                                                                                                  self.label_set_price_markup4))

        self.label_set_price_markup5 = tkinter.Label(self.frame3, text=("Наценка 5 = " + str(self.price_markup5) + "%"))
        self.label_choose_cats_to_markup5 = tkinter.Label(self.frame3, text="Категории товаров для наценки 5:")
        self.button_price_markup5 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup5_group))
        self.button_set_price_markup5 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup5,
                                                                                                  self.label_set_price_markup5))

        self.label_set_price_markup6 = tkinter.Label(self.frame3, text=("Наценка 6 = " + str(self.price_markup6) + "%"))
        self.label_choose_cats_to_markup6 = tkinter.Label(self.frame3, text="Категории товаров для наценки 6:")
        self.button_price_markup6 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup6_group))
        self.button_set_price_markup6 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup6,
                                                                                                  self.label_set_price_markup6))

        self.label_set_price_markup7 = tkinter.Label(self.frame3, text=("Наценка 7 = " + str(self.price_markup7) + "%"))
        self.label_choose_cats_to_markup7 = tkinter.Label(self.frame3, text="Категории товаров для наценки 7:")
        self.button_price_markup7 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup7_group))
        self.button_set_price_markup7 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup7,
                                                                                                  self.label_set_price_markup7))

        self.label_set_price_markup8 = tkinter.Label(self.frame3, text=("Наценка 8 = " + str(self.price_markup8) + "%"))
        self.label_choose_cats_to_markup8 = tkinter.Label(self.frame3, text="Категории товаров для наценки 8:")
        self.button_price_markup8 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup8_group))
        self.button_set_price_markup8 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup8,
                                                                                                  self.label_set_price_markup8))

        self.label_set_price_markup9 = tkinter.Label(self.frame3, text=("Наценка 9 = " + str(self.price_markup9) + "%"))
        self.label_choose_cats_to_markup9 = tkinter.Label(self.frame3, text="Категории товаров для наценки 9:")
        self.button_price_markup9 = tkinter.Button(self.frame3, text="Выбрать.",
                                                   command=lambda: self.cats_markup_util(self.chosen_cats_markup9_group))
        self.button_set_price_markup9 = tkinter.Button(self.frame3, text="Изменить.",
                                                       command=lambda: self.set_price_markup_util(self.price_markup9,
                                                                                                  self.label_set_price_markup9))

        self.label_set_price_markup10 = tkinter.Label(self.frame3, text=("Наценка 10 = " + str(self.price_markup10) + "%"))
        self.label_choose_cats_to_markup10 = tkinter.Label(self.frame3, text="Категории товаров для наценки 10:")
        self.button_price_markup10 = tkinter.Button(self.frame3,
                                                    text="Выбрать.",
                                                    command=lambda: self.cats_markup_util(self.chosen_cats_markup10_group))
        self.button_set_price_markup10 = tkinter.Button(self.frame3, text="Изменить.",
                                                        command=lambda: self.set_price_markup_util(self.price_markup10,
                                                                                                   self.label_set_price_markup10))
        # Текст справки перемещен внутрь основного окна по желанию пользователя
        self.help = ts.ScrolledText(self.frame5)


    def download_file(self):

        try:
            preparation = list(self.hyperlink.split("/"))
            preparation2 = list(preparation[6].split("?"))
            result_filename = str(self.osdir + '/' + self.input_dir_name + '/' + preparation2[0])
            self.csv_path = result_filename
        except IndexError:
            error = tm.showerror(title="Ошибка!",
                                 message="Введена не корректная ссылка.")

        self.label_download_stat.config(text="Идет загрузка файла...")

        with open(str(self.data_dir_name + '/' + "input.dat"), 'wb') as file:
            pickle.dump(self.csv_path, file)
        self.input_entry.config(text=self.csv_path)
        # аутпут будет в каталоге исполняемого файла
        self.output = str(self.osdir + '/' + self.csvfile_name)
        self.output_entry.config(text=self.output)

        with open(str(self.data_dir_name + '/' + "output.dat"), 'wb') as file2:
            pickle.dump(self.output, file2)

        try:

            ur.urlretrieve(self.hyperlink, self.csv_path)
            self.label_download_stat.config(text="Файл загружен, попытка его прочесть...")
        except ue.URLError:
            self.label_download_stat.config(text="Не удалось скачать файл, проверьте сетевое подключение.")
        except IndexError:
            error = tm.showerror(title="Ошибка!",
                                 message="Введена не корректная ссылка.")
        except ValueError:
            error = tm.showerror(title="Ошибка!",
                                 message="Введена не корректная ссылка.")
        try:
            self.reader(self.csv_path)
            self.label_download_stat.config(text="Файл прочитан, можно удалять категории.")
        except FileNotFoundError:
            error = tm.showerror(title="Ошибка!",
                                 message="Скачанный файл отсутсвует, "
                                         "либо загрузился в нечитаемом виде, повторите загрузку.")

    def set_hyperlink(self):
        val = tkinter.StringVar()
        new_hyperlink = tkinter.Toplevel(self.root, bd=1)
        val.set(self.hyperlink)
        new_hyperlink.title("Изменить ссылку на скачивание файла")
        new_hyperlink.minsize(width=20, height=20)
        hyperlink_entry = tkinter.Entry(new_hyperlink, width=30, bd=1, exportselection=0, textvariable=val)
        hyperlink_labler = tkinter.Label(new_hyperlink, text="Введите новую ссылку:")
        ok_button = tkinter.Button(new_hyperlink, text="Ok",
                                       command=lambda: self.set_hyperlink_ok(val, new_hyperlink))
        abort_button = tkinter.Button(new_hyperlink, text="Отмена", command=new_hyperlink.destroy)
        hyperlink_labler.grid(row=0, column=1)
        hyperlink_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_hyperlink_ok(self, val, new_hyperlink):
        """Сохраняем ссылку"""
        self.hyperlink = val.get()
        with open(str(self.data_dir_name + '/' + "hyperlink.dat"), 'wb') as file:
            pickle.dump(self.hyperlink, file)
        new_hyperlink.destroy()

    def draw_me(self):
        """размещение и отображение элементов графического интерфейса главного окна"""
        # alpha_frame:
        self.alpha_frame.pack(side="left")
        # omega_frame
        self.omega_frame.pack(side="left", expand=1, fill="both")
        # frame 1
        self.frame1.pack()
        self.label_download_stat.grid(row=0, column=0)
        self.button_change_hyperlink.grid(row=0, column=2)
        self.button_download.grid(row=0, column=1)
        # frame 6
        self.frame6.pack()
        self.label_input.grid(row=0, column=0)
        self.input_entry.grid(row=1, column=0)
        self.set_input_path.grid(row=1, column=1)
        self.label_output.grid(row=3, column=0)
        self.output_entry.grid(row=4, column=0)
        self.set_output_path.grid(row=4, column=1)
        self.button3.grid(row=5, column=0)
        # frame 2
        self.frame2.pack()
        self.label_cats_to_die.grid(row=6, column=0)
        self.cats_to_die_dialog.grid(row=7, column=0)
        self.label_con_course.grid(row=6, column=1)
        self.button2.grid(row=7, column=1)
        # frame 4
        self.frame4.pack(side="bottom")
        self.button1.grid(row=29, column=0)
        self.quit_button.grid(row=29, column=4)
        # новые элементы ГУИ
        # Frame 3
        self.frame3.pack()
        self.label_set_price_markup1.grid(row=9, column=3)
        self.label_choose_cats_to_markup1.grid(row=9, column=0)
        self.button_price_markup1.grid(row=9, column=1)
        self.button_set_price_markup1.grid(row=9, column=4)
        self.label_set_price_markup2.grid(row=10, column=3)
        self.label_choose_cats_to_markup2.grid(row=10, column=0)
        self.button_price_markup2.grid(row=10, column=1)
        self.button_set_price_markup2.grid(row=10, column=4)
        # 3-4
        self.label_set_price_markup3.grid(row=11, column=3)
        self.label_choose_cats_to_markup3.grid(row=11, column=0)
        self.button_price_markup3.grid(row=11, column=1)
        self.button_set_price_markup3.grid(row=11, column=4)
        self.label_set_price_markup4.grid(row=12, column=3)
        self.label_choose_cats_to_markup4.grid(row=12, column=0)
        self.button_price_markup4.grid(row=12, column=1)
        self.button_set_price_markup4.grid(row=12, column=4)
        # 5-6
        self.label_set_price_markup5.grid(row=13, column=3)
        self.label_choose_cats_to_markup5.grid(row=13, column=0)
        self.button_price_markup5.grid(row=13, column=1)
        self.button_set_price_markup5.grid(row=13, column=4)
        self.label_set_price_markup6.grid(row=14, column=3)
        self.label_choose_cats_to_markup6.grid(row=14, column=0)
        self.button_price_markup6.grid(row=14, column=1)
        self.button_set_price_markup6.grid(row=14, column=4)
        # 7 -8
        self.label_set_price_markup7.grid(row=15, column=3)
        self.label_choose_cats_to_markup7.grid(row=15, column=0)
        self.button_price_markup7.grid(row=15, column=1)
        self.button_set_price_markup7.grid(row=15, column=4)
        self.label_set_price_markup8.grid(row=16, column=3)
        self.label_choose_cats_to_markup8.grid(row=16, column=0)
        self.button_price_markup8.grid(row=16, column=1)
        self.button_set_price_markup8.grid(row=16, column=4)
        # 9 -10
        self.label_set_price_markup9.grid(row=17, column=3)
        self.label_choose_cats_to_markup9.grid(row=17, column=0)
        self.button_price_markup9.grid(row=17, column=1)
        self.button_set_price_markup9.grid(row=17, column=4)
        self.label_set_price_markup10.grid(row=18, column=3)
        self.label_choose_cats_to_markup10.grid(row=18, column=0)
        self.button_price_markup10.grid(row=18, column=1)
        self.button_set_price_markup10.grid(row=18, column=4)
        # Frame 5
        self.frame5.pack(expand=1, fill="both")

        self.help.pack(expand=1, fill="both")
        help_text = ("""    
        Инструкция по работе с программой:
	1. Загрузка файла с orion.de:
При нажатии кнопки "Скачать", будет выполнена автоматическая загрузка 
файла (по умолчанию ссылка уже занесена в программу). При необходимости 
изменить ссылку, нужно воспользоваться кнопкой "Обновить файл"
	2. Обработка загруженного файла:
	2.1 Путь к исходному файлу выгрузки - подставляется автоматически, 
если была произведена автоматическая загрузка.Если исходный прайс сохранен 
локально, для его обработки нажать на кнопку "..." и указать путь к файлу.
	2.2 Указать путь к каталогу, куда будет помещен результат - подставляется 
автоматически, если была произведена автоматическая загрузка.
Если необходимо сохранять в другую директорию, нажать на кнопку "..." и 
указать путь для сохранения результата.
	2.3 Прочитать файл - после получения исходного файла и каталога 
для результата, нажать на кнопку "Прочитать файл".
	3. Отбор лишних категорий и курс конвертации цен:
Категории для удаления, кнопка "Выбрать" позволяет в отдельно окне выбрать 
категории для удаления, которые не будут отображены в файле результата.
Курс RUB к EUR: 00.0 - содержит сохраненный в программе курс, для 
изменения нажать кнопку "Изменить курс рубля".
	4. Работа с наценками: Доступно максимум 10 наценок. 
Для установки наценки, необходимо:
	1) выбрать категории для наценки: кнопка "Выбрать"; 
Для изминения уже выбранных категории, нужно очистить текущее значение 
выбранных категории, для этого нажать "Выбрать", в открывшемся окне 
"Отмена" (или закрыть окно), после выбрать категории снова. 
	2) установить размер наценки в процентах: кнопка "Изменить";
5. Завершение работы:
Для выполнения нажать кнопку "Обработать". При успешной обработке появится 
окно сообщение "Результат обработки сохранен в <путь к файлу>".
Кнопка "Выход" произведет закрытие программы без обработки файла""")
        self.help.insert(1.0, help_text)
        self.help.configure(state="disabled")
        # Loop
        self.root.mainloop()

    def set_price_markup_util(self, markup, label):
        """Дочернее окно изменения наценки"""
        value = tkinter.DoubleVar()
        value.set(markup)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok",
                                   command=lambda: self.set_price_markup_util_ok(value, new_price_markup, markup, label))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup_util_ok(self, value, new_price_markup, markup, label):
        """Применение введенного пользователем значения наценки"""
        try:
            markup = value.get()
            new_price_markup.destroy()
            if label == self.label_set_price_markup1:
                label.config(text="Наценка 1 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup1.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup2:
                label.config(text="Наценка 2 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup2.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup3:
                label.config(text="Наценка 3 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup3.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup4:
                label.config(text="Наценка 4 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup4.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup5:
                label.config(text="Наценка 5 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup5.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup6:
                label.config(text="Наценка 6 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup6.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup7:
                label.config(text="Наценка 7 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup7.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup8:
                label.config(text="Наценка 8 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup8.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup9:
                label.config(text="Наценка 9 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup9.dat"), 'wb') as file:
                    pickle.dump(markup, file)
            elif label == self.label_set_price_markup10:
                label.config(text="Наценка 10 = " + str(markup) + " %")
                with open(str(self.data_dir_name + '/' + "markup10.dat"), 'wb') as file:
                    pickle.dump(markup, file)

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть числом с плавающей точкой "
                                                          "(десятичной дробью), "
                                                          "с точкой, как разделительным знаком.")

    def cats_markup_util(self, any_list):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки"""
        value = tkinter.StringVar()
        value.set(any_list)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if any_list:
            for i in any_list:
                self.chosen_cats_markup_general_group.remove(i)

            any_list = []
        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup_util_ok(choose_cats, available_cats, any_list))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup_util_ok(self, choose_cats, available_cats, any_list):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки"""
        y = choose_cats.curselection()
        if any_list:
            any_list = []

        for i in y:
            print(i)
            any_list.append(choose_cats.get(i))
        print(any_list)

        for i in any_list:
            self.chosen_cats_markup_general_group.add(i)

        available_cats.destroy()

    def cats_catcher(self):
        """Дочернее окно выбора категорий товаров, подлежащих удалению"""
        if self.xlsx_bool:
            angry = tm.showerror(title="Внимание!",
                                 message="В папке с исполняемым файлом обнаружен файл соответствия групп товаров, "
                                         "возможность удалять категори вручную отключена.")
        else:
            value = tkinter.DoubleVar()
            value.set(self.convertation_cource)
            available_cats = tkinter.Toplevel(self.root, bd=1)
            available_cats.title("Выберите лишние категории товаров:")
            available_cats.minsize(width=50, height=50)
            choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
            xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
            ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
            choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
            if self.cats_to_die:
                self.cats_to_die = []
            if self.invited_cats_on_markup:
                self.invited_cats_on_markup = set()
            for i in self.sorted_cats:
                if i not in self.cats_to_die:
                    choose_cats.insert("end", i)

            ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_hanging(choose_cats, available_cats))
            abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
            choose_cats.grid(row=0, column=1, )
            ok_button.grid(row=2, column=0)
            abort_button.grid(row=2, column=4)
            xbar.grid(row=1, column=1, rowspan=1, sticky="we")
            ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_hanging(self, choose_cats, available_cats):
        """Функция формирует список индексов выбранных для удаления категорий товаров"""

        y = choose_cats.curselection()
        self.cats_to_die = []

        for i in y:
            self.cats_to_die.append(choose_cats.get(i))

        for i in self.sorted_cats:
            if i not in self.cats_to_die:
                # print(i)
                self.invited_cats_on_markup.add(i)
        available_cats.destroy()

    def set_last_params(self):
       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "hyperlink.dat"), 'rb') as file15:
               self.hyperlink = pickle.load(file15)

       except FileNotFoundError:
           pass

       try:
            with open(str(self.osdir + '/' + self.data_dir_name + '/' + "concourse.dat"), 'rb') as file:
                self.convertation_cource = pickle.load(file)
                self.label_con_course.config(text="Курс RUB к EUR:" + str(self.convertation_cource))

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup1.dat"), 'rb') as file5:
               self.price_markup1 = pickle.load(file5)
               self.label_set_price_markup1.config(text="Наценка 1 = " + str(self.price_markup1) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup2.dat"), 'rb') as file6:
               self.price_markup2 = pickle.load(file6)
               self.label_set_price_markup2.config(text="Наценка 2 = " + str(self.price_markup2) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup3.dat"), 'rb') as file7:
               self.price_markup3 = pickle.load(file7)
               self.label_set_price_markup3.config(text="Наценка 3 = " + str(self.price_markup3) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup4.dat"), 'rb') as file8:
               self.price_markup4 = pickle.load(file8)
               self.label_set_price_markup4.config(text="Наценка 4 = " + str(self.price_markup4) + "%")

       except FileNotFoundError:
           pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup5.dat"), 'rb') as file9:
               self.price_markup5 = pickle.load(file9)
               self.label_set_price_markup5.config(text="Наценка 5 = " + str(self.price_markup5) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup6.dat"), 'rb') as file10:
               self.price_markup6 = pickle.load(file10)
               self.label_set_price_markup6.config(text="Наценка 6 = " + str(self.price_markup6) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup7.dat"), 'rb') as file11:
               self.price_markup7 = pickle.load(file11)
               self.label_set_price_markup7.config(text="Наценка 7 = " + str(self.price_markup7) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup8.dat"), 'rb') as file12:
               self.price_markup8 = pickle.load(file12)
               self.label_set_price_markup8.config(text="Наценка 8 = " + str(self.price_markup8) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup9.dat"), 'rb') as file13:
               self.price_markup9 = pickle.load(file13)
               self.label_set_price_markup9.config(text="Наценка 9 = " + str(self.price_markup9) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "markup10.dat"), 'rb') as file14:
               self.price_markup10 = pickle.load(file14)
               self.label_set_price_markup10.config(text="Наценка 10 = " + str(self.price_markup10) + "%")

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "input.dat"), 'rb') as file2:
               self.csv_path = pickle.load(file2)

               self.input_entry.config(text=self.csv_path)

       except FileNotFoundError:
            pass

       try:
           with open(str(self.osdir + '/' + self.data_dir_name + '/' + "output.dat"), 'rb') as file3:
               self.output = pickle.load(file3)
               self.output_entry.config(text=self.output)

       except FileNotFoundError:
            pass

    def open_file(self):
        x = tf.askopenfile()
        try:
            self.csv_path = x.name
        except AttributeError:
            pass

        if self.csv_path:

            self.csvfile_name = str("result.csv")
            self.input_entry.config(text=self.csv_path)
            with open(str(self.data_dir_name + '/' + "input.dat"), 'wb') as file:
                pickle.dump(self.csv_path, file)

    def path_to_save(self):
        self.out = tf.askdirectory()
        if self.csv_path:
            try:
                self.output = self.out + '/' + self.csvfile_name
            except TypeError:
                self.output = self.out

        self.output_entry.config(text=self.output)
        with open(str(self.data_dir_name + '/' + "output.dat"), 'wb') as file:
            pickle.dump(self.output, file)

    def change_con_course(self):
        value = tkinter.DoubleVar()
        value.set(self.convertation_cource)
        new_course = tkinter.Toplevel(self.root, bd=1)
        new_course.title("Изменить курс рубля")
        new_course.minsize(width=20, height=20)
        course_entry = tkinter.Entry(new_course, width=30, bd=1, exportselection=0, textvariable=value)
        course_labler = tkinter.Label(new_course, text="Ввести курс рубля к евро:")
        ok_button = tkinter.Button(new_course, text="Ok", command=lambda: self.change_con_course_ok(value, new_course))
        abort_button = tkinter.Button(new_course, text="Отмена", command=new_course.destroy)
        course_labler.grid(row=0, column=1)
        course_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def change_con_course_ok(self, value, new_course):
        try:
            self.convertation_cource = value.get()
            with open(str(self.data_dir_name + '/' + "concourse.dat"), 'wb') as file:
                pickle.dump(self.convertation_cource, file)
            new_course.destroy()
            self.label_con_course.config(text="Курс RUB к EUR:" + str(self.convertation_cource))

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!",
                                 message="Вводимое значение должно содержать арабские цифры и точку в качестве разделителя.")

    def reader(self, path):
        csv.register_dialect('orion', delimiter=';', quoting=csv.QUOTE_ALL)
        with open(path, 'r', encoding='utf-8', newline='') as file:
            reader = csv.DictReader(file, dialect='orion', quotechar='"', doublequote=True, skipinitialspace=True)
            try:
                for line in reader:
                    self.buffer.append(line)
                    self.cats.append(line['category_path'])
                    self.cats_id.append(line['category'])
            except KeyError:
                anger = tm.showerror(title="Ошибка!",
                                     message="Прочитанный файл возможно ошибочен и не соответствует требуемой структуре данных."
                                                              "\n Удостоверьтесь, что входной файл корректен.")
        self.sorted_cats = list(set(self.cats))
        self.sorted_cats_id = list(set(self.cats_id))
        if self.xlsx_bool:
            csv.register_dialect('g_c_c', delimiter=",", quoting=csv.QUOTE_NONE)
            xlsx_csv_path = str(self.osdir + "/" + self.data_dir_name + "/" + 'group_codes_correspondence.csv')
            with open(xlsx_csv_path, 'r', encoding='utf-8', newline='') as file2:
                reader2 = csv.DictReader(file2, dialect='g_c_c', skipinitialspace=True)

                for line in reader2:
                    if line['name']:
                        self.xlsx_list.append(line)
                        self.xlsx_id.add(line['orn'])
            other = set()
            another_other = set()

            for line in self.buffer:
                if line['category'] not in self.xlsx_id:

                    other.add(line['category_path'])
                else:
                    another_other.add(line['category_path'])
            self.cats_to_die = list(other)
            self.invited_cats_on_markup = another_other

    def buffer_handler(self):
        if self.cats_to_die:
            for line in self.buffer:
                if line['category_path'] not in self.cats_to_die:

                    self.buffer2.append(line)
        else:
            for line in self.buffer:
                self.buffer2.append(line)
        self.buffer = []
        self.serpentis()

    def serpentis(self):
        if self.buffer2:

            for row in self.buffer2:

                if row['special_price_flag'] == 'X':

                    product_price = float(row['product_price'])
                    row['product_price'] = float(round(product_price * self.convertation_cource))

                else:
                    product_price = float(row['product_price'])
                    row['product_price'] = float(round(product_price / 2 * self.convertation_cource))
                self.buffer3.append(row)
            self.buffer2 = []
        self.serpentis_iter2()

    def serpentis_iter2(self):
            for row in self.buffer3:
                if self.price_markup1:
                    print(self.chosen_cats_markup1_group)
                    if row['category_path'] in self.chosen_cats_markup1_group:
                        markuped_price1 = float(row['product_price'])
                        print("markuped_price1", markuped_price1)
                        row['product_price'] = float(round(markuped_price1 + self.price_markup1 * markuped_price1))
                        print('product_price', row['product_price'])
                    else:
                        print("Это баг!!!")
                if self.price_markup2:
                    if row['category_path'] in self.chosen_cats_markup2_group:
                        markuped_price2 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price2 + self.price_markup2 * markuped_price2))
                if self.price_markup3:
                    if row['category_path'] in self.chosen_cats_markup3_group:
                        markuped_price3 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price3 + self.price_markup3 * markuped_price3))
                if self.price_markup4:
                    if row['category_path'] in self.chosen_cats_markup4_group:
                        markuped_price4 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price4 + self.price_markup4 * markuped_price4))
                if self.price_markup5:
                    if row['category_path'] in self.chosen_cats_markup5_group:
                        markuped_price5 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price5 + self.price_markup5 * markuped_price5))
                if self.price_markup6:
                    if row['category_path'] in self.chosen_cats_markup6_group:
                        markuped_price6 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price6 + self.price_markup6 * markuped_price6))
                if self.price_markup7:
                    if row['category_path'] in self.chosen_cats_markup7_group:
                        markuped_price7 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price7 + self.price_markup7 * markuped_price7))
                if self.price_markup8:
                    if row['category_path'] in self.chosen_cats_markup8_group:
                        markuped_price8 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price8 + self.price_markup8 * markuped_price8))
                if self.price_markup9:
                    if row['category_path'] in self.chosen_cats_markup9_group:
                        markuped_price9 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price9 + self.price_markup9 * markuped_price9))
                if self.price_markup10:
                    if row['category_path'] in self.chosen_cats_markup10_group:
                        markuped_price10 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price10 + self.price_markup10 * markuped_price10))
                self.rows.append(row)
            self.buffer3 = []
            self.serpentis_iter3()

    def serpentis_iter3(self):
            if self.xlsx_bool:
                self.keys_to_kill = ['export_version', 'current_version', 'timestamp', 'product_class',
                                     'product_name_DE',
                                     'title_DE', 'title_EN', 'title_IT', 'product_keyword', 'special_price_flag',
                                     'recommended_selling_price', 'full_text_DE', 'full_text_IT',
                                     'category_path', 'group_product', 'single_image', 'tax_id', 'ean_code',
                                     'battery_supply', 'battery_combination', 'product_image_2', 'product_image_3',
                                     'product_image_4',
                                     'material_description', 'material_thickness', 'total_length', 'diameter', 'weight',
                                     'length',
                                     'width', 'height', 'contents', 'availability', 'delivery_week', 'dress_size',
                                     'shoe_size',
                                     'main_color', 'color_description', 'battery_type', 'packaging_size', 'aroma',
                                     'isbn', 'number_of_pages',
                                     'total_duration', 'hardcore_flag', 'age_rating', 'product_languages',
                                     'country_of_origin_code',
                                     'series_name', 'reference_quantity', 'reference_price_factor',
                                     'promotional_packing', 'product_icons',
                                     'selling_unit', 'voc_percentage', 'barcode', 'barcode_type',
                                     'food_information_flag',
                                     'detailed_full_text_de', 'detailed_full_text_en', 'group_id', 'tariff_number']
            for row in self.rows:
                for i in self.keys_to_kill:
                    if i in row.keys():
                        row.pop(i)

                row['name : Название'] = row.pop('product_name_EN')
                row['vendor : Производитель'] = row.pop('label_name')
                row['image : Иллюстрация'] = row.pop('product_image_1')
                row['supplier : Поставщик'] = 'orion.de'
                row['pre_order : Предзаказ'] = 0
                row['article : Артикул'] = str(row.pop('product_id'))
                row['code_1c : 1C'] = ''
                if self.xlsx_bool:
                    if row['category'] in self.xlsx_id:
                        for n in self.xlsx_list:
                            if n['orn'] == row['category']:
                                row['folder : Категория'] = n['descr']
                        row.pop('category')
                else:
                    row['folder : Категория'] = ''
                row['tags : Теги'] = ''
                row['hidden : Скрыто'] = 0
                row['kind_id : ID'] = ''
                row['is_kind : Модификация'] = ''
                row['discounted : Товар уже со скидкой'] = 0
                row['note : Анонс товара'] = ''
                row['body : Описание'] = row.pop('full_text_EN')
                row['amount : Количество'] = float(1000.00)
                row['unit : Единица измерения'] = ''
                row['weight : Вес'] = 0
                row['weight_unit : Единица веса'] = 'kg'
                row['dimensions : Габариты'] = "0,0,0"
                row['new : Новинка'] = row.pop('novelty_flag')
                row['special : Спецпредложение'] = 0
                row['yml : Yandex.Market'] = 0

                row['price_old : Старая цена'] = row.pop('product_price')
                if str(row['article : Артикул'][0]) == '2' or str(row['article : Артикул'][0]) == '0' \
                        and str(row['article : Артикул'][1]) == '2':
                    row['price : Цена'] = round(float(row['price_old : Старая цена'] - (
                            25 * (row['price_old : Старая цена']) / 100)))
                else:
                    row['price : Цена'] = round(float(row['price_old : Старая цена'] - (
                            10 * (row['price_old : Старая цена']) / 100)))
                row['price2 : Цена 2'] = row['price_old : Старая цена']
                row['price3 : Цена 3'] = float(0.00)
                row['currency : Валюта'] = 'RUB'
                row['seo_noindex : Индексация'] = 0
                row['seo_h1 : Заголовок (H1)'] = ''
                row['seo_title : Title'] = ''
                row['seo_description : Description'] = ''
                row['seo_keywords : Keywords'] = ''
                row['sef_url : ЧПУ'] = ''
                row['uuid : UUID Товара'] = ''
                row['uuid_mod : UUID Модификации'] = ''
                self.outlines.append(row)
            self.rows = []
            self.writer()

    def writer(self):
            try:
                with open(self.output, 'w', encoding='utf-8', newline='') as output_file:
                    writer = csv.DictWriter(output_file, dialect='orion', fieldnames=self.etha_keys, delimiter=';',
                                quotechar='"', doublequote=True, skipinitialspace=True)

                    writer.writeheader()

                    for line in self.outlines:
                        writer.writerow(line)
            except IsADirectoryError:
                with open((self.output + "result.csv"), 'w', encoding='utf-8', newline='') as output_file:
                    writer = csv.DictWriter(output_file, dialect='orion', fieldnames=self.etha_keys, delimiter=';',
                                quotechar='"', doublequote=True, skipinitialspace=True)

                    writer.writeheader()

                    for line in self.outlines:
                        writer.writerow(line)

            self.final_msg = tm.showinfo(title="Готово.", message=("Результат обработки сохранен в  " + self.output))

    def check_dirs(self):
        """метод призван создавать список каталогов, для хранения служебных данных, относительно текущего каталога,
        в котором выполняется"""
        if not os.path.isdir(str(self.osdir) + self.data_dir_name):
            try:
                os.mkdir(self.data_dir_name)
            except FileExistsError:
                pass
        if not os.path.isdir(str(self.osdir) + self.input_dir_name):
            try:
                os.mkdir(self.input_dir_name)
            except FileExistsError:
                pass

    def xlsx_to_csv(self):
        if os.path.isfile(str(self.osdir) + "/" + "соответствие кодов групп.xlsx"):
            self.xlsx_bool = True
            print(self.xlsx_bool)
            xlsx_workbook = xlrd.open_workbook('соответствие кодов групп.xlsx')
            xlsx_file = xlsx_workbook.sheet_by_name('Лист1')
            g_c_c_path = str(self.osdir + "/" + self.data_dir_name + "/" + 'group_codes_correspondence.csv')
            with open(g_c_c_path, 'w', encoding='utf8', newline='') as file:
                writer = csv.writer(file, quoting=csv.QUOTE_NONE, escapechar='\\')
                for rownum in range(xlsx_file.nrows):
                    writer.writerow(xlsx_file.row_values(rownum))


snake = CsvParser()
snake.check_dirs()
snake.xlsx_to_csv()
snake.set_last_params()
snake.draw_me()