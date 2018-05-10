# -- coding: utf-8 --
from __future__ import unicode_literals
import csv
import tkinter
import tkinter.filedialog as tf
import tkinter.messagebox as tm
import pickle



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
        self.outlines = []
        self.root = tkinter.Tk()
        self.root.title("Orion export v. 1.1")
        self.root.geometry = "300x250"
        self.button1 = tkinter.Button(self.root, text="Обработать", command=self.buffer_handler)
        self.button2 = tkinter.Button(self.root, text="Изменить курс рубля", command=self.change_con_course)
        self.button3 = tkinter.Button(self.root, text="Прочитать файл", command=self.reader)
        self.label_input = tkinter.Label(self.root, text="Путь к исходному файлу выгрузки:")
        self.input_entry = tkinter.Label(self.root, text=self.output)
        self.label_output = tkinter.Label(self.root, text="Указать путь к каталогу, куда будет помещен результат:")
        self.output_entry = tkinter.Label(self.root, text=self.output)
        self.label_cats_to_die = tkinter.Label(self.root, text="Категории товаров для удаления:")
        self.cats_to_die_dialog = tkinter.Button(self.root, text="Выбрать категории товаров для удаления.", command=self.cats_catcher)
        self.quit_button = tkinter.Button(self.root, text="Выход", command=self.root.destroy)
        self.set_input_path = tkinter.Button(self.root, text="...", command=self.open_file)
        self.set_output_path = tkinter.Button(self.root, text="...", command=self.path_to_save)
        self.label_con_course = tkinter.Label(self.root, text=("Курс RUB к EUR:" + str(self.convertation_cource)))
        self.final_msg = None

    def draw_me(self):
        """размещение и отображение элементов графического интерфейса главного окна"""
        self.label_input.grid(row=0, column=0)
        self.input_entry.grid(row=1, column=0)
        self.set_input_path.grid(row=1, column=1)
        self.label_output.grid(row=3, column=0)
        self.output_entry.grid(row=4, column=0)
        self.set_output_path.grid(row=4, column=1)
        self.button3.grid(row=5, column=0)
        self.label_cats_to_die.grid(row=6, column=0)
        self.cats_to_die_dialog.grid(row=7, column=0)
        self.label_con_course.grid(row=6, column=1)
        self.button2.grid(row=7, column=1)
        self.button1.grid(row=9, column=0)
        self.quit_button.grid(row=9, column=1)
        self.root.mainloop()

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
        """"""
        y = choose_cats.curselection()
        self.cats_to_die = []

        for i in y:

                if self.sorted_cats[i]:
                    self.cats_to_die.append(self.sorted_cats[i])
        with open("deletecat.dat", "wb") as file:
            pickle.dump(self.cats_to_die, file)
        available_cats.destroy()

    def set_last_params(self):
       try:
           with open("deletecat.dat", 'rb') as file4:
               self.cats_to_die = pickle.load(file4)

       except FileNotFoundError:
           pass

       try:
            with open("concourse.dat", 'rb') as file:
                self.convertation_cource = pickle.load(file)
                self.label_con_course.config(text="Курс RUB к EUR:" + str(self.convertation_cource))

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


    def reader(self):
        csv.register_dialect('orion', delimiter=';', quoting=csv.QUOTE_ALL)
        with open(self.csv_path, 'r', encoding='utf-8', newline='') as file:
            reader = csv.DictReader(file, dialect='orion', quotechar='"', doublequote=True, skipinitialspace=True)

            for line in reader:
                self.buffer.append(line)
                self.cats.append(line['category_path'])
                self.cats_id.append(line['category'])
        self.sorted_cats = list(set(self.cats))
        # в целях вероятной доработки, храним уникальные значения id категорий товаров
        self.sorted_cats_id = list(set(self.cats_id))

    def buffer_handler(self):
        if self.cats_to_die:
            for line in self.buffer:
                if line['category_path'] not in self.cats_to_die:
                    print(line)
                    self.rows.append(line)

        elif not self.cats_to_die:
            for line in self.buffer:
                self.rows.append(line)
        self.buffer = []
        self.serpentis()

    def serpentis(self):
        if self.rows:

            for row in self.rows:

                if row['special_price_flag'] == 'X':

                    product_price = float(row['product_price'])
                    row['product_price'] = round((product_price * self.convertation_cource), 2)

                elif row['special_price_flag'] != 'X':

                    product_price = float(row['product_price'])
                    row['product_price'] = round((product_price / 2 * self.convertation_cource), 2)

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


    def help(self):
        help_wind = tkinter.Toplevel(self.root, bd=1)
        help_wind.title("Справка.")
        help_wind.minsize(width=20, height=20)
        help_labler1 = tkinter.Label(help_wind, text="1. Сначала укажите путь к файлу выгрузки с сайта orion.de, кликнув по кнопке '...'.")
        help_labler2 = tkinter.Label(help_wind, text="2. Затем укажите каталог, в который будет сохранен результат обработки (вторая сверху кнопка '...').")
        help_labler3 = tkinter.Label(help_wind, text="3. Измените курс рубля, если это необходимо. Важно, при указании нового курса использовать точку, а не запятую.")
        help_labler4 = tkinter.Label(help_wind, text="4. Кликнике по кнопке 'Прочитать файл'.")
        help_labler5 = tkinter.Label(help_wind, text="5. Кликните по кнопке 'Выбрать категории товаров для удаления'.")
        help_labler6 = tkinter.Label(help_wind, text="6. В открывшемся окне, кликая левой клавишей мыши, выделите лишние категории.")
        help_labler7 = tkinter.Label(help_wind, text="7. Нажмите 'Обработать.'")

        ok_button = tkinter.Button(help_wind, text="Ok", command=help_wind.destroy)
        help_labler1.grid(row=0, column=0)
        help_labler2.grid(row=1, column=0)
        help_labler3.grid(row=2, column=0)
        help_labler4.grid(row=3, column=0)
        help_labler5.grid(row=4, column=0)
        help_labler6.grid(row=5, column=0)
        help_labler7.grid(row=6, column=0)
        ok_button.grid(row=7, column=0)


snake = CsvParser()
snake.help()
snake.set_last_params()
snake.draw_me()








