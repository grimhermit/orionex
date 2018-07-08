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

        # атрибут хранит ссылку
        self.hyperlink = str('')
        self.csv_path = str('')
        self.csvfile_name = str('')
        self.convertation_cource = float(0)
        self.output = "output/"
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
        # self.label_hyperlink = tkinter.Label(self.frame1, text=str(self.hyperlink))
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
        self.price_markup1 = int(0)
        # Атрибут для хранения значения наценки 2
        self.price_markup2 = int(0)
        self.price_markup3 = int(0)
        self.price_markup4 = int(0)
        self.price_markup5 = int(0)
        self.price_markup6 = int(0)
        self.price_markup7 = int(0)
        self.price_markup8 = int(0)
        self.price_markup9 = int(0)
        self.price_markup10 = int(0)
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
        self.invited_cats_on_markup = []
        # добавляем элементы графического интерфейса
        # self.label_price_markup = tkinter.Label(self.frame3, text="Работа с наценками:")
        self.label_set_price_markup1 = tkinter.Label(self.frame3, text=("Наценка 1 = " + str(self.price_markup1) + "%"))
        self.label_choose_cats_to_markup1 = tkinter.Label(self.frame3, text="Категории товаров для наценки 1:")
        self.button_price_markup1 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup1)
        self.button_set_price_markup1 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup1)

        self.label_set_price_markup2 = tkinter.Label(self.frame3, text=("Наценка 2 = " + str(self.price_markup2) + "%"))
        self.label_choose_cats_to_markup2 = tkinter.Label(self.frame3, text="Категории товаров для наценки 2:")
        self.button_price_markup2 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup2)
        self.button_set_price_markup2 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup2)

        self.label_set_price_markup3 = tkinter.Label(self.frame3, text=("Наценка 3 = " + str(self.price_markup3) + "%"))
        self.label_choose_cats_to_markup3 = tkinter.Label(self.frame3, text="Категории товаров для наценки 3:")
        self.button_price_markup3 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup3)
        self.button_set_price_markup3 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup3)

        self.label_set_price_markup4 = tkinter.Label(self.frame3, text=("Наценка 4 = " + str(self.price_markup4) + "%"))
        self.label_choose_cats_to_markup4 = tkinter.Label(self.frame3, text="Категории товаров для наценки 4:")
        self.button_price_markup4 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup4)
        self.button_set_price_markup4 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup4)

        self.label_set_price_markup5 = tkinter.Label(self.frame3, text=("Наценка 5 = " + str(self.price_markup5) + "%"))
        self.label_choose_cats_to_markup5 = tkinter.Label(self.frame3, text="Категории товаров для наценки 5:")
        self.button_price_markup5 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup5)
        self.button_set_price_markup5 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup5)

        self.label_set_price_markup6 = tkinter.Label(self.frame3, text=("Наценка 6 = " + str(self.price_markup6) + "%"))
        self.label_choose_cats_to_markup6 = tkinter.Label(self.frame3, text="Категории товаров для наценки 6:")
        self.button_price_markup6 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup6)
        self.button_set_price_markup6 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup6)

        self.label_set_price_markup7 = tkinter.Label(self.frame3, text=("Наценка 7 = " + str(self.price_markup7) + "%"))
        self.label_choose_cats_to_markup7 = tkinter.Label(self.frame3, text="Категории товаров для наценки 7:")
        self.button_price_markup7 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup7)
        self.button_set_price_markup7 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup7)

        self.label_set_price_markup8 = tkinter.Label(self.frame3, text=("Наценка 8 = " + str(self.price_markup8) + "%"))
        self.label_choose_cats_to_markup8 = tkinter.Label(self.frame3, text="Категории товаров для наценки 8:")
        self.button_price_markup8 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup8)
        self.button_set_price_markup8 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup8)

        self.label_set_price_markup9 = tkinter.Label(self.frame3, text=("Наценка 9 = " + str(self.price_markup9) + "%"))
        self.label_choose_cats_to_markup9 = tkinter.Label(self.frame3, text="Категории товаров для наценки 9:")
        self.button_price_markup9 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup9)
        self.button_set_price_markup9 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup9)

        self.label_set_price_markup10 = tkinter.Label(self.frame3, text=("Наценка 10 = " + str(self.price_markup10) + "%"))
        self.label_choose_cats_to_markup10 = tkinter.Label(self.frame3, text="Категории товаров для наценки 10:")
        self.button_price_markup10 = tkinter.Button(self.frame3, text="Выбрать.", command=self.cats_markup10)
        self.button_set_price_markup10 = tkinter.Button(self.frame3, text="Изменить.", command=self.set_price_markup10)
        # Текст справки перемещен внутрь основного окна по желанию пользователя
        self.help = ts.ScrolledText(self.frame5)

    def download_file(self):
        preparation = list(self.hyperlink.split("/"))
        preparation2 = list(preparation[6].split("?"))
        result_filename = preparation2[0]
        try:
            self.label_download_stat.config(text="Идет загрузка файла...")
            ur.urlretrieve(self.hyperlink, result_filename)
            self.label_download_stat.config(text="Файл загружен, попытка его прочесть...")
        except ue.URLError:
            self.label_download_stat.config(text="Не удалось скачать файл, проверьте сетевое подключение.")

        try:
            self.reader(result_filename)
            self.label_download_stat.config(text="Файл прочитан, можно удалять категории.")
        except FileNotFoundError:
            error = tm.showerror(title="Ошибка!", message="Скачанный файл отсутсвует, либо загрузился в нечитаемом виде, повторите загрузку.")

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
        with open("hyperlink.dat", 'wb') as file:
            pickle.dump(self.hyperlink, file)
        new_hyperlink.destroy()


    def draw_me(self):
        """размещение и отображение элементов графического интерфейса главного окна"""
        #alpha_frame:
        self.alpha_frame.pack(side="left")
        #omega_frame
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
        # self.label_price_markup.grid(row=8, column=1)
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
        Уважаемый пользователь, программа Orion export приветствует Вас!
    Мы надеемся, что данная инструкция поможет Вам эффективно использовать все 
возможности программы с момента первого знакомства с ней. 

    Слева Вы видите элементы интерфейса, разделенные на группы по тематической 
принадлежности и пронумерованные в порядке работы с ними:
1) группа "1. Загрузка файла с сайта orion.de" позволит скачать файл выгрузки 
товарных остатков прямо с сайта orion.de по ссылке, которую можно ввести в 
программу по нажатию на кнопку "Обновить ссылку", программа запомнит введенную 
ссылку и ее не нужно будет вводить каждый раз, нажатие кнопки "Скачать" даст 
программе сигнал на загрузку файла из сети, после чего файл будет сразу 
прочитан (по умолчанию файл скачивается в директорию, в которой программа 
находится в данный момент, одноименный файл будет перезаписан) и можно будет 
приступить к дальнейшим действиям, а в поле "Статус загрузки файла" 
отобразится результат скачивания файла, если что-то пошло не так (например 
пропало подключение к сети или введенная ссылка оказалась неправильной) 
в данном поле отобразится соответствующее сообщение;
2) группа "2. Обработка загруженного ранее файла" пригодится, 
если Вы хотите обработать уже загруженный файл, который нет нужды скачивать, 
а также она полезна тем, что позволяет указать программе, куда нужно сохранить 
обработанный файл, нажатие на верхнюю кнопку "..." позволяет Вам указать путь 
к файлу вручную, а на нижнюю - куда программа сохранит результат Вашей 
совместной работы, кнопка "Прочитать файл" даст программе команду считать 
данные из файла, путь к которому был указан вручную, программа запоминает 
введенные пути, что избавляет от необходимости вводить их каждый раз 
при ее запуске;
3) группа "3. Отбор лишних категорий и курс конвертации цен" позволит выбрать 
из исходного файла те категории товаров, которые Вы не хотите видеть в конечном 
файле по нажатию на кнопку "Выбрать", а кнопка "Изменить курс рубля" позволит 
откорректировать курс рубля к евро (курс указывается в виде десятичной дроби, 
разделитель - точка) после подтверждения ввода курс будет записан программой 
для того, чтобы не вводить его каждый раз при ее запуске;
4) группа "4. Работа с наценками" позволит Вам выбрать до 10 разных групп 
категорий товаров, к каждой из которых будет применена соответствующая наценка 
по нажатию  на кнопку "Выбрать" (к каждой категории товаров может быть 
применена только одна наценка и если вы ошиблись, просто нажмите по кнопкe 
"Выбрать" еще раз, чтобы повторить выбор, кнопка "Изменить" выведет 
небольшое окно, в котором можно указать размер наценки, вводить нужно 
только целые числа, наценка рассчитывается в процентах;
5) группа "5. Завершение работы" - это самая важная и последняя группа, 
конечный файл будет сформирован только после нажатия на кнопку "Обработать", 
во всех предыдущих пунктах Вы лишь указывали программе как и какие данные 
будут сохранены в конечном файле, по окончанию записи конечного файла 
программа выведет сообщение о том, что результат сохранен, после чего, 
можно ее закрыть нажатием на кнопку "Выход". """)
        self.help.insert(1.0, help_text)
        self.help.configure(state="disabled")
        # Loop
        self.root.mainloop()

# 10
    def set_price_markup10(self):
        """Дочернее окно изменения наценки 10"""
        value10 = tkinter.IntVar()
        value10.set(self.price_markup10)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 10")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value10)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup10_ok(value10, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup10_ok(self, value10, new_price_markup):
        """Применение введенного пользователем значения наценки 10"""
        try:
            self.price_markup10 = value10.get()
            with open("markup10.dat", 'wb') as file:
                pickle.dump(self.price_markup10, file)
            new_price_markup.destroy()
            self.label_set_price_markup10.config(text="Наценка 10 = " + str(self.price_markup10) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup10(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 10"""
        value = tkinter.IntVar()
        value.set(self.price_markup10)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup10_group:
            for i in self.chosen_cats_markup10_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup10_group = []
        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup10_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup10_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 10"""
        y = choose_cats.curselection()
        self.chosen_cats_markup10_group = []

        for i in y:
            self.chosen_cats_markup10_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup10_group)

        for i in self.chosen_cats_markup10_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()

# 9
    def set_price_markup9(self):
        """Дочернее окно изменения наценки 9"""
        value9 = tkinter.IntVar()
        value9.set(self.price_markup9)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 9")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value9)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup9_ok(value9, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup9_ok(self, value9, new_price_markup):
        """Применение введенного пользователем значения наценки 9"""
        try:
            self.price_markup9 = value9.get()
            with open("markup9.dat", 'wb') as file:
                pickle.dump(self.price_markup9, file)
            new_price_markup.destroy()
            self.label_set_price_markup9.config(text="Наценка 9 = " + str(self.price_markup9) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup9(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 9"""
        value = tkinter.IntVar()
        value.set(self.price_markup9)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup9_group:
            for i in self.chosen_cats_markup9_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup9_group = []

        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup9_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup9_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 1"""
        y = choose_cats.curselection()
        self.chosen_cats_markup9_group = []

        for i in y:
            self.chosen_cats_markup9_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup9_group)

        for i in self.chosen_cats_markup9_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()

# 8
    def set_price_markup8(self):
        """Дочернее окно изменения наценки 8"""
        value8 = tkinter.IntVar()
        value8.set(self.price_markup8)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 8")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value8)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup8_ok(value8, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup8_ok(self, value8, new_price_markup):
        """Применение введенного пользователем значения наценки 8"""
        try:
            self.price_markup8 = value8.get()
            with open("markup8.dat", 'wb') as file:
                pickle.dump(self.price_markup8, file)
            new_price_markup.destroy()
            self.label_set_price_markup8.config(text="Наценка 8 = " + str(self.price_markup8) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup8(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 8"""
        value = tkinter.IntVar()
        value.set(self.price_markup8)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup8_group:
            for i in self.chosen_cats_markup8_group:
                self.chosen_cats_markup_general_group.remove(i)
            self.chosen_cats_markup8_group = []
        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup8_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup8_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 1"""
        y = choose_cats.curselection()
        self.chosen_cats_markup8_group = []

        for i in y:
            self.chosen_cats_markup8_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup8_group)

        for i in self.chosen_cats_markup8_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()

# 7

    def set_price_markup7(self):
        """Дочернее окно изменения наценки 7"""
        value7 = tkinter.IntVar()
        value7.set(self.price_markup7)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 7")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value7)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup7_ok(value7, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup7_ok(self, value7, new_price_markup):
        """Применение введенного пользователем значения наценки 7"""
        try:
            self.price_markup7 = value7.get()
            with open("markup7.dat", 'wb') as file:
                pickle.dump(self.price_markup7, file)
            new_price_markup.destroy()
            self.label_set_price_markup7.config(text="Наценка 7 = " + str(self.price_markup7) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup7(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 7"""
        value = tkinter.IntVar()
        value.set(self.price_markup7)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup7_group:
            for i in self.chosen_cats_markup7_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i, len(self.chosen_cats_markup_general_group))
            self.chosen_cats_markup7_group = []
        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup7_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup7_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 7"""
        y = choose_cats.curselection()
        self.chosen_cats_markup7_group = []

        for i in y:
            self.chosen_cats_markup7_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup7_group)

        for i in self.chosen_cats_markup7_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()

# 6

    def set_price_markup6(self):
        """Дочернее окно изменения наценки 6"""
        value6 = tkinter.IntVar()
        value6.set(self.price_markup6)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 6")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value6)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup6_ok(value6, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup6_ok(self, value6, new_price_markup):
        """Применение введенного пользователем значения наценки 6"""
        try:
            self.price_markup6 = value6.get()
            with open("markup6.dat", 'wb') as file:
                pickle.dump(self.price_markup6, file)
            new_price_markup.destroy()
            self.label_set_price_markup6.config(text="Наценка 6 = " + str(self.price_markup6) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup6(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 6"""
        value = tkinter.IntVar()
        value.set(self.price_markup6)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup6_group:
            for i in self.chosen_cats_markup6_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup6_group = []

        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup6_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup6_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 6"""
        y = choose_cats.curselection()
        self.chosen_cats_markup6_group = []

        for i in y:
            self.chosen_cats_markup6_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup6_group)

        for i in self.chosen_cats_markup6_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()

# 5
    def set_price_markup5(self):
        """Дочернее окно изменения наценки 5"""
        value5 = tkinter.IntVar()
        value5.set(self.price_markup5)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 5")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value5)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup5_ok(value5, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup5_ok(self, value5, new_price_markup):
        """Применение введенного пользователем значения наценки 5"""
        try:
            self.price_markup5 = value5.get()
            with open("markup5.dat", 'wb') as file:
                pickle.dump(self.price_markup5, file)
            new_price_markup.destroy()
            self.label_set_price_markup5.config(text="Наценка 5 = " + str(self.price_markup5) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup5(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 5"""
        value = tkinter.IntVar()
        value.set(self.price_markup5)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup5_group:
            for i in self.chosen_cats_markup5_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup5_group = []

        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup5_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup5_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 5"""
        y = choose_cats.curselection()
        self.chosen_cats_markup5_group = []

        for i in y:
            self.chosen_cats_markup5_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup5_group)

        for i in self.chosen_cats_markup5_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()

# 4
    def set_price_markup4(self):
        """Дочернее окно изменения наценки 4"""
        value4 = tkinter.IntVar()
        value4.set(self.price_markup4)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 4")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value4)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup4_ok(value4, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup4_ok(self, value4, new_price_markup):
        """Применение введенного пользователем значения наценки 4"""
        try:
            self.price_markup4 = value4.get()
            with open("markup4.dat", 'wb') as file:
                pickle.dump(self.price_markup4, file)
            new_price_markup.destroy()
            self.label_set_price_markup4.config(text="Наценка 4 = " + str(self.price_markup4) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup4(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 4"""
        value = tkinter.IntVar()
        value.set(self.price_markup4)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup4_group:
            for i in self.chosen_cats_markup4_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup4_group = []

        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup4_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup4_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 4"""
        y = choose_cats.curselection()
        self.chosen_cats_markup4_group = []

        for i in y:
            self.chosen_cats_markup4_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup4_group)

        for i in self.chosen_cats_markup4_group:
            self.chosen_cats_markup_general_group.add(i)
            print("добалено в общее множество " + i)
        available_cats.destroy()
# 3

    def set_price_markup3(self):
        """Дочернее окно изменения наценки 3"""
        value3 = tkinter.IntVar()
        value3.set(self.price_markup3)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 3")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value3)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup3_ok(value3, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup3_ok(self, value3, new_price_markup):
        """Применение введенного пользователем значения наценки 3"""
        try:
            self.price_markup3 = value3.get()
            with open("markup3.dat", 'wb') as file:
                pickle.dump(self.price_markup3, file)
            new_price_markup.destroy()
            self.label_set_price_markup3.config(text="Наценка 3 = " + str(self.price_markup3) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup3(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 3"""
        value = tkinter.IntVar()
        value.set(self.price_markup3)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup3_group:
            for i in self.chosen_cats_markup3_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup3_group = []

        if self.cats_to_die:
            print("cats to die is TRUE")
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup3_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup3_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 1"""
        y = choose_cats.curselection()
        self.chosen_cats_markup3_group = []

        for i in y:
            self.chosen_cats_markup3_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup3_group)

        for i in self.chosen_cats_markup3_group:
            print("добалено в общее множество " + i)
            self.chosen_cats_markup_general_group.add(i)
        available_cats.destroy()

# 2

    def set_price_markup2(self):
        """Дочернее окно изменения наценки 2"""
        value2 = tkinter.IntVar()
        value2.set(self.price_markup2)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 2")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value2)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup2_ok(value2, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup2_ok(self, value2, new_price_markup):
        """Применение введенного пользователем значения наценки 2"""
        try:
            self.price_markup2 = value2.get()
            with open("markup2.dat", 'wb') as file:
                pickle.dump(self.price_markup2, file)
            new_price_markup.destroy()
            self.label_set_price_markup2.config(text="Наценка 2 = " + str(self.price_markup2) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup2(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 2"""
        value = tkinter.IntVar()
        value.set(self.price_markup2)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup2_group:
            for i in self.chosen_cats_markup2_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup2_group = []

        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup2_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup2_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 1"""
        y = choose_cats.curselection()
        self.chosen_cats_markup2_group = []

        for i in y:
            self.chosen_cats_markup2_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup2_group)

        for i in self.chosen_cats_markup2_group:
            print("добалено в общее множество " + i)
            self.chosen_cats_markup_general_group.add(i)
        available_cats.destroy()
# 1

    def set_price_markup1(self):
        """Дочернее окно изменения наценки 1"""
        value = tkinter.IntVar()
        value.set(self.price_markup1)
        new_price_markup = tkinter.Toplevel(self.root, bd=1)
        new_price_markup.title("Изменить наценку 1")
        new_price_markup.minsize(width=20, height=20)
        markup_entry = tkinter.Entry(new_price_markup, width=30, bd=1, exportselection=0, textvariable=value)
        markup_labler = tkinter.Label(new_price_markup, text="Ввести процент наценки:")
        ok_button = tkinter.Button(new_price_markup, text="Ok", command=lambda: self.set_price_markup1_ok(value, new_price_markup))
        abort_button = tkinter.Button(new_price_markup, text="Отмена", command=new_price_markup.destroy)
        markup_labler.grid(row=0, column=1)
        markup_entry.grid(row=1, column=1)
        ok_button.grid(row=3, column=0)
        abort_button.grid(row=3, column=3)

    def set_price_markup1_ok(self, value, new_price_markup):
        """Применение введенного пользователем значения наценки 1"""
        try:
            self.price_markup1 = value.get()
            with open("markup1.dat", 'wb') as file:
                pickle.dump(self.price_markup1, file)
            new_price_markup.destroy()
            self.label_set_price_markup1.config(text="Наценка 1 = " + str(self.price_markup1) + " %")

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно быть целым числом, без разделительных знаков.")

    def cats_markup1(self):
        """Дочернее окно выбора категорий товаров, подлежащих применению наценки 1"""
        value = tkinter.StringVar()
        value.set(self.price_markup1)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите категории товаров для применения к ним выбранной наценки:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)

        if self.chosen_cats_markup1_group:
            for i in self.chosen_cats_markup1_group:
                self.chosen_cats_markup_general_group.remove(i)
                print(i)
            self.chosen_cats_markup1_group = []

        if self.cats_to_die:

            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        if not self.cats_to_die:
            self.invited_cats_on_markup = self.sorted_cats
            for i in self.invited_cats_on_markup:
                if i not in self.chosen_cats_markup_general_group:
                    choose_cats.insert("end", i)

        ok_button = tkinter.Button(available_cats, text="Ok", command=lambda: self.cats_markup1_ok(choose_cats, available_cats))
        abort_button = tkinter.Button(available_cats, text="Отмена", command=available_cats.destroy)
        choose_cats.grid(row=0, column=1, )
        ok_button.grid(row=2, column=0)
        abort_button.grid(row=2, column=4)
        xbar.grid(row=1, column=1, rowspan=1, sticky="we")
        ybar.grid(row=0, column=3, columnspan=1, sticky="ns")

    def cats_markup1_ok(self, choose_cats, available_cats):
        """Функция формирует список индексов категорий товаров, выбранных для применения наценки 1"""
        y = choose_cats.curselection()
        self.chosen_cats_markup1_group = []
        print(y)
        for i in y:
            self.chosen_cats_markup1_group.append(choose_cats.get(i))
        print(self.chosen_cats_markup1_group)

        for i in self.chosen_cats_markup1_group:
            print("добалено в общее множество " + i)
            self.chosen_cats_markup_general_group.add(i)

        available_cats.destroy()

    def cats_catcher(self):
        """Дочернее окно выбора категорий товаров, подлежащих удалению"""
        value = tkinter.DoubleVar()
        value.set(self.convertation_cource)
        available_cats = tkinter.Toplevel(self.root, bd=1)
        available_cats.title("Выберите лишние категории товаров:")
        available_cats.minsize(width=50, height=50)
        choose_cats = tkinter.Listbox(available_cats, selectmode="multiple", width=50, height=30)
        xbar = tkinter.Scrollbar(available_cats, orient='horizontal', command=choose_cats.xview)
        ybar = tkinter.Scrollbar(available_cats, orient='vertical', command=choose_cats.yview)
        choose_cats.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        for i in self.sorted_cats:
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

                if self.sorted_cats[i]:
                    self.cats_to_die.append(self.sorted_cats[i])
        print("кошки на смерть" + str(self.cats_to_die))
        for i in self.sorted_cats:
            if i not in self.cats_to_die:
                self.invited_cats_on_markup.append(i)

        print(self.invited_cats_on_markup)
        with open("deletecat.dat", "wb") as file:
            pickle.dump(self.cats_to_die, file)
        available_cats.destroy()

    def set_last_params(self):
       try:
           with open("hyperlink.dat", 'rb') as file15:
               self.hyperlink = pickle.load(file15)
             #  self.label_hyperlink.config(text=str(self.hyperlink))
       except FileNotFoundError:
           pass

       # try:
           # with open("deletecat.dat", 'rb') as file4:
               # self.cats_to_die = pickle.load(file4)

       # except FileNotFoundError:
           # pass

       try:
            with open("concourse.dat", 'rb') as file:
                self.convertation_cource = pickle.load(file)
                self.label_con_course.config(text="Курс RUB к EUR:" + str(self.convertation_cource))

       except FileNotFoundError:
            pass

       try:
           with open("markup1.dat", 'rb') as file5:
               self.price_markup1 = pickle.load(file5)
               self.label_set_price_markup1.config(text="Наценка 1 = " + str(self.price_markup1) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup2.dat", 'rb') as file6:
               self.price_markup2 = pickle.load(file6)
               self.label_set_price_markup2.config(text="Наценка 2 = " + str(self.price_markup2) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup3.dat", 'rb') as file7:
               self.price_markup3 = pickle.load(file7)
               self.label_set_price_markup3.config(text="Наценка 3 = " + str(self.price_markup3) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup4.dat", 'rb') as file8:
               self.price_markup4 = pickle.load(file8)
               self.label_set_price_markup4.config(text="Наценка 4 = " + str(self.price_markup4) + "%")

       except FileNotFoundError:
           pass

       try:
           with open("markup5.dat", 'rb') as file9:
               self.price_markup5 = pickle.load(file9)
               self.label_set_price_markup5.config(text="Наценка 5 = " + str(self.price_markup5) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup6.dat", 'rb') as file10:
               self.price_markup6 = pickle.load(file10)
               self.label_set_price_markup6.config(text="Наценка 6 = " + str(self.price_markup6) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup7.dat", 'rb') as file11:
               self.price_markup7 = pickle.load(file11)
               self.label_set_price_markup7.config(text="Наценка 7 = " + str(self.price_markup7) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup8.dat", 'rb') as file12:
               self.price_markup8 = pickle.load(file12)
               self.label_set_price_markup8.config(text="Наценка 8 = " + str(self.price_markup8) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup9.dat", 'rb') as file13:
               self.price_markup9 = pickle.load(file13)
               self.label_set_price_markup9.config(text="Наценка 9 = " + str(self.price_markup9) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("markup10.dat", 'rb') as file14:
               self.price_markup10 = pickle.load(file14)
               self.label_set_price_markup10.config(text="Наценка 10 = " + str(self.price_markup10) + "%")

       except FileNotFoundError:
            pass

       try:
           with open("input.dat", 'rb') as file2:
               self.csv_path = pickle.load(file2)

               self.input_entry.config(text=self.csv_path)

       except FileNotFoundError:
            pass

       try:
           with open("output.dat", 'rb') as file3:
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
            self.csvfile_name = str(self.csv_path.split('/').pop())

            self.input_entry.config(text=self.csv_path)
            with open("input.dat", 'wb') as file:
                pickle.dump(self.csv_path, file)

    def path_to_save(self):
        self.out = tf.askdirectory()
        if self.csv_path:
            try:
                self.output = self.out + '/' + self.csvfile_name
            except TypeError:
                self.output = self.out

        self.output_entry.config(text=self.output)
        with open("output.dat", 'wb') as file:
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
            with open("concourse.dat", 'wb') as file:
                pickle.dump(self.convertation_cource, file)
            new_course.destroy()
            self.label_con_course.config(text="Курс RUB к EUR:" + str(self.convertation_cource))

        except tkinter.TclError:
            angry = tm.showerror(title="Ошибка!", message="Вводимое значение должно содержать арабские цифры и точку в качестве разделителя.")

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
                anger = tm.showerror(title="Ошибка!", message="Прочитанный файл возможно ошибочен и не соответствует требуемой структуре данных."
                                                              "\n Удостоверьтесь, что входной файл корректен.")

        self.sorted_cats = list(set(self.cats))

        self.sorted_cats_id = list(set(self.cats_id))

    def buffer_handler(self):
        if self.cats_to_die:
            for line in self.buffer:
                if line['category_path'] not in self.cats_to_die:

                    self.buffer2.append(line)

        # elif not self.cats_to_die:
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

                # elif row['special_price_flag'] != 'X':
                else:
                    product_price = float(row['product_price'])
                    row['product_price'] = float(round(product_price / 2 * self.convertation_cource))
                self.buffer3.append(row)
            self.buffer2 = []
        self.serpentis_iter2()

    def serpentis_iter2(self):
            for row in self.buffer3:

                if self.price_markup1:
                    if row['category_path'] in self.chosen_cats_markup1_group:
                        markuped_price1 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price1 + self.price_markup1 * markuped_price1 / 100))
                if self.price_markup2:
                    if row['category_path'] in self.chosen_cats_markup2_group:
                        markuped_price2 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price2 + self.price_markup2 * markuped_price2 / 100))
                if self.price_markup3:
                    if row['category_path'] in self.chosen_cats_markup3_group:
                        markuped_price3 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price3 + self.price_markup3 * markuped_price3 / 100))
                if self.price_markup4:
                    if row['category_path'] in self.chosen_cats_markup4_group:
                        markuped_price4 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price4 + self.price_markup4 * markuped_price4 / 100))
                if self.price_markup5:
                    if row['category_path'] in self.chosen_cats_markup5_group:
                        markuped_price5 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price5 + self.price_markup5 * markuped_price5 / 100))
                if self.price_markup6:
                    if row['category_path'] in self.chosen_cats_markup6_group:
                        markuped_price6 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price6 + self.price_markup6 * markuped_price6 / 100))
                if self.price_markup7:
                    if row['category_path'] in self.chosen_cats_markup7_group:
                        markuped_price7 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price7 + self.price_markup7 * markuped_price7 / 100))
                if self.price_markup8:
                    if row['category_path'] in self.chosen_cats_markup8_group:
                        markuped_price8 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price8 + self.price_markup8 * markuped_price8 / 100))
                if self.price_markup9:
                    if row['category_path'] in self.chosen_cats_markup9_group:
                        markuped_price9 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price9 + self.price_markup9 * markuped_price9 / 100))
                if self.price_markup10:
                    if row['category_path'] in self.chosen_cats_markup10_group:
                        markuped_price10 = float(row['product_price'])
                        row['product_price'] = float(round(markuped_price10 + self.price_markup10 * markuped_price10 / 100))
                self.rows.append(row)
            self.buffer3 = []
            self.serpentis_iter3()

    def serpentis_iter3(self):
            for row in self.rows:
                for i in self.keys_to_kill:
                    if i in row.keys():
                        row.pop(i)

                row['name : Название'] = row.pop('product_name_EN')
                row['vendor : Производитель'] = row.pop('label_name')
                row['image : Иллюстрация'] = row.pop('product_image_1')
                row['supplier : Поставщик'] = 'orion.de'
                row['pre_order : Предзаказ'] = 0
                row['article : Артикул'] = row.pop('product_id')
                row['code_1c : 1C'] = ''
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
                row['price : Цена'] = row.pop('product_price')
                row['price_old : Старая цена'] = row['price : Цена']
                row['price2 : Цена 2'] = row['price : Цена']
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
                with open((self.output + "file.csv"), 'w', encoding='utf-8', newline='') as output_file:
                    writer = csv.DictWriter(output_file, dialect='orion', fieldnames=self.etha_keys, delimiter=';',
                                quotechar='"', doublequote=True, skipinitialspace=True)

                    writer.writeheader()

                    for line in self.outlines:
                        writer.writerow(line)

            self.final_msg = tm.showinfo(title="Готово.", message=("Результат обработки сохранен в  " + self.output))


snake = CsvParser()
snake.set_last_params()
snake.draw_me()

