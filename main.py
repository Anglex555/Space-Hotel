import sys
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QTabWidget, QTableWidget, QTableWidgetItem, QLabel, QComboBox, QHBoxLayout, QPushButton, QDateEdit, QLineEdit, QMessageBox
from PyQt5.QtGui import QPixmap, QFontDatabase, QPalette, QColor, QFont
from PyQt5.QtCore import QFile, QTextStream, Qt, QPropertyAnimation, QEasingCurve


class HotelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Отельный менеджер")
        self.setGeometry(100, 100, 1280, 720)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)

        self.main_page = QWidget()
        self.hotels_tab = QWidget()
        self.inventory_tab = QWidget()
        self.regions_tab = QWidget()
        self.guests_tab = QWidget()
        self.pre_sales_tab = QWidget()
        self.sales_analysis_tab = QWidget()
        self.housekeeping_tab = QWidget()

        self.tab_widget.addTab(self.main_page, "Главная")
        self.tab_widget.addTab(self.hotels_tab, "Отели")
        self.tab_widget.addTab(self.inventory_tab, "Номерной фонд")
        self.tab_widget.addTab(self.regions_tab, "Регионы")
        self.tab_widget.addTab(self.guests_tab, "Гости")
        self.tab_widget.addTab(self.pre_sales_tab, "Предварительная продажа") 
        self.tab_widget.addTab(self.sales_analysis_tab, "Анализ эффективности продаж")
        self.tab_widget.addTab(self.housekeeping_tab, "Уборка номеров")

        self.init_main_page()
        self.init_hotels_tab()
        self.init_inventory_tab()
        self.init_regions_tab()
        self.init_guests_tab()
        self.init_pre_sales_tab()
        self.init_sales_analysis_tab()
        self.init_housekeeping_tab()

        self.apply_stylesheet()

    def apply_stylesheet(self):
        style_file = QFile("styles.css")
        style_file.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(style_file)
        self.setStyleSheet(stream.readAll())

    def init_main_page(self):
        layout = QVBoxLayout()
        self.main_page.setLayout(layout)

        title_label = QLabel("<h1><font color='#0072bc'>SPACE</font><font color='#ffffff'> Hotel</font></h1>")
        title_label.setStyleSheet("font-size: 54px;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        logo_label = QLabel()
        pixmap = QPixmap("image1.jpg")
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo_label)

        description_label = QLabel(
        "<p style='font-size: 20px; color: #0A1527;'>SPACE Hotel - это уникальная сеть отелей, которая предоставляет "
        "неповторимый опыт в путешествиях и отдыхе. Мы предлагаем комфортабельное размещение в "
        "космической тематике, обеспечивая вам роскошь и удовольствие в самых далеких уголках "
        "вселенной. Наше приложение для автоматизации позволяет легко и удобно бронировать номера и "
        "получать доступ к эксклюзивным предложениям и скидкам.</p>"
        )
        description_label.setWordWrap(True)
        layout.addWidget(description_label)

    def init_hotels_tab(self):
        layout = QVBoxLayout()
        self.hotels_tab.setLayout(layout)

        hotels_table = QTableWidget()
        hotels_table.setColumnCount(2)
        hotels_table.setHorizontalHeaderLabels(["Название отеля", "Регион"])

        hotels_data = pd.read_excel("hotels.xls")

        for index, row in hotels_data.iterrows():
            name = row["Отель"]
            region = row["Регион"]
            hotels_table.setRowCount(index + 1)
            hotels_table.setItem(index, 0, QTableWidgetItem(name))
            hotels_table.setItem(index, 1, QTableWidgetItem(region))

        layout.addWidget(hotels_table)

        font = QFont()
        font.setPointSize(12)
        hotels_table.horizontalHeader().setFont(font)
        hotels_table.verticalHeader().setFont(font)
        for i in range(hotels_table.columnCount()):
            hotels_table.horizontalHeaderItem(i).setFont(font)
            hotels_table.setColumnWidth(i, 300)
        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18) 
        hotels_table.setFont(font_content)
        hotels_table.setStyleSheet("color: #0A1527;")

        table_animation = QPropertyAnimation(hotels_table, b'size')
        table_animation.setDuration(1000)
        table_animation.setStartValue(hotels_table.size())
        table_animation.setEndValue(hotels_table.size() * 1.1)
        table_animation.setEasingCurve(QEasingCurve.OutBounce)
        table_animation.start()

    def init_inventory_tab(self):
        layout = QVBoxLayout()
        self.inventory_tab.setLayout(layout)

        inventory_table = QTableWidget()
        inventory_table.setColumnCount(4)
        inventory_table.setHorizontalHeaderLabels(["Номер", "Категория", "Количество мест", "Статус"])

        inventory_data = pd.read_excel("inventory.xls")

        for index, row in inventory_data.iterrows():
            number = row["Номер"]
            category = row["Категория"]
            seats = row["Количество мест"]
            status = row["Статус"]
            inventory_table.setRowCount(index + 1)
            inventory_table.setItem(index, 0, QTableWidgetItem(number))
            inventory_table.setItem(index, 1, QTableWidgetItem(category))
            inventory_table.setItem(index, 2, QTableWidgetItem(str(seats)))
            inventory_table.setItem(index, 3, QTableWidgetItem(status))

        layout.addWidget(inventory_table)

        font = QFont()
        font.setPointSize(12)
        inventory_table.horizontalHeader().setFont(font)
        inventory_table.verticalHeader().setFont(font)
        for i in range(inventory_table.columnCount()):
            inventory_table.horizontalHeaderItem(i).setFont(font)
            inventory_table.setColumnWidth(i, 200)

        table_animation = QPropertyAnimation(inventory_table, b'size')
        table_animation.setDuration(1000)
        table_animation.setStartValue(inventory_table.size())
        table_animation.setEndValue(inventory_table.size() * 1.1)
        table_animation.setEasingCurve(QEasingCurve.OutBounce)
        table_animation.start()

        def filter_inventory():
            selected_category = category_combobox.currentText()
            selected_seats = seats_combobox.currentText()

            for row in range(inventory_table.rowCount()):
                category_item = inventory_table.item(row, 1)
                seats_item = inventory_table.item(row, 2)

                if (selected_category == "Все" or category_item.text() == selected_category) and \
                        (selected_seats == "Все" or seats_item.text() == selected_seats):
                    inventory_table.setRowHidden(row, False)
                else:
                    inventory_table.setRowHidden(row, True)
        
        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18) 

        category_combobox = QComboBox()
        category_combobox.addItem("Все")
        for category in inventory_data["Категория"].unique():
            category_combobox.addItem(category)
        category_combobox.currentIndexChanged.connect(filter_inventory)
        category_combobox.setFont(font_content)
        category_combobox.setStyleSheet("color: #0A1527;")

        seats_combobox = QComboBox()
        seats_combobox.addItem("Все")
        for seats in inventory_data["Количество мест"].unique():
            seats_combobox.addItem(str(seats))
        seats_combobox.currentIndexChanged.connect(filter_inventory)
        seats_combobox.setFont(font_content) 
        seats_combobox.setStyleSheet("color: #0A1527;")

        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18) 
        inventory_table.setFont(font_content)
        inventory_table.setStyleSheet("color: #0A1527;")
        
        category_label = QLabel("Категория:")
        category_label.setFont(font_content)
        category_label.setStyleSheet("color: #0A1527;")
        seats_label = QLabel("Количество мест:")
        seats_label.setFont(font_content)
        seats_label.setStyleSheet("color: #0A1527;")

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(category_label)
        filter_layout.addWidget(category_combobox)
        filter_layout.addWidget(seats_label)
        filter_layout.addWidget(seats_combobox)
        layout.addLayout(filter_layout)

    def filter_inventory(self):
        text = self.search_line_edit.text()
        if text:
            filtered_data = self.inventory_data[self.inventory_data["Номер"].astype(str).str.contains(text, case=False)]
            self.populate_inventory_table(filtered_data)
        else:
            self.populate_inventory_table(self.inventory_data)

    def populate_inventory_table(self, data):
        self.inventory_table.clearContents()
        for index, row in data.iterrows():
            number = row["Номер"]
            category = row["Категория"]
            seats = row["Количество мест"]
            self.inventory_table.setRowCount(index + 1)
            self.inventory_table.setItem(index, 0, QTableWidgetItem(str(number)))
            self.inventory_table.setItem(index, 1, QTableWidgetItem(category))
            self.inventory_table.setItem(index, 2, QTableWidgetItem(str(seats)))

    def init_regions_tab(self):
        layout = QVBoxLayout()
        self.regions_tab.setLayout(layout)

        regions_table = QTableWidget()
        regions_table.setColumnCount(1)
        regions_table.setHorizontalHeaderLabels(["Регион"])

        regions_data = pd.read_excel("regions.xls")

        for index, row in regions_data.iterrows():
            region_name = row["Название"]
            regions_table.setRowCount(index + 1)
            regions_table.setItem(index, 0, QTableWidgetItem(region_name))

        layout.addWidget(regions_table)

        font = QFont()
        font.setPointSize(12)
        regions_table.horizontalHeader().setFont(font)
        regions_table.verticalHeader().setFont(font)
        regions_table.horizontalHeaderItem(0).setFont(font)
        regions_table.setColumnWidth(0, 400)
        
        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18)
        regions_table.setFont(font_content)
        regions_table.setStyleSheet("color: #0A1527;")

        table_animation = QPropertyAnimation(regions_table, b'size')
        table_animation.setDuration(1000)
        table_animation.setStartValue(regions_table.size())
        table_animation.setEndValue(regions_table.size() * 1.1)
        table_animation.setEasingCurve(QEasingCurve.OutBounce)
        table_animation.start()

    def init_guests_tab(self):
        layout = QVBoxLayout()
        self.guests_tab.setLayout(layout)

        guests_table = QTableWidget()
        guests_table.setColumnCount(2)
        guests_table.setHorizontalHeaderLabels(["ФИО", "Телефон"])

        data = [("Иванов Иван Иванович", "+7 (901) 123-45-67"),
                ("Петров Петр Петрович", "+7 (902) 234-56-78"),
                ("Сидоров Сидор Сидорович", "+7 (903) 345-67-89"),
                ("Козлова Ольга Викторовна", "+7 (904) 456-78-90"),
                ("Смирнова Елена Ивановна", "+7 (905) 567-89-01"),
                ("Васильева Анна Сергеевна", "+7 (906) 678-90-12"),
                ("Михайлов Михаил Михайлович", "+7 (907) 789-01-23"),
                ("Новиков Николай Николаевич", "+7 (908) 890-12-34")]
        guests_table.setRowCount(len(data))
        for i, (name, phone) in enumerate(data):
            guests_table.setItem(i, 0, QTableWidgetItem(name))
            guests_table.setItem(i, 1, QTableWidgetItem(phone))

        layout.addWidget(guests_table)

        font = QFont()
        font.setPointSize(12)
        guests_table.horizontalHeader().setFont(font)
        guests_table.verticalHeader().setFont(font)
        for i in range(guests_table.columnCount()):
            guests_table.horizontalHeaderItem(i).setFont(font)
            guests_table.setColumnWidth(i, 450)
        
        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18) 
        guests_table.setFont(font_content)
        guests_table.setStyleSheet("color: #0A1527;")

        table_animation = QPropertyAnimation(guests_table, b'size')
        table_animation.setDuration(1000)
        table_animation.setStartValue(guests_table.size())
        table_animation.setEndValue(guests_table.size() * 1.1)
        table_animation.setEasingCurve(QEasingCurve.OutBounce)
        table_animation.start()

    def init_pre_sales_tab(self):
        layout = QVBoxLayout()
        self.pre_sales_tab.setLayout(layout)

        self.pre_sales_table = QTableWidget()
        self.pre_sales_table.setColumnCount(7)
        self.pre_sales_table.setHorizontalHeaderLabels(["Дата", "Отель", "Номер", "Дата заезда", "Дата выезда", "Стоимость", "Количество гостей"])

        layout.addWidget(self.pre_sales_table)

        font = QFont()
        font.setPointSize(12)
        self.pre_sales_table.horizontalHeader().setFont(font)
        self.pre_sales_table.verticalHeader().setFont(font)
        for i in range(self.pre_sales_table.columnCount()):
            self.pre_sales_table.horizontalHeaderItem(i).setFont(font)
            self.pre_sales_table.setColumnWidth(i, 150)
        
        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18) 
        self.pre_sales_table.setFont(font_content)
        self.pre_sales_table.setStyleSheet("color: #0A1527;")

        table_animation = QPropertyAnimation(self.pre_sales_table, b'size')
        table_animation.setDuration(1000)
        table_animation.setStartValue(self.pre_sales_table.size())
        table_animation.setEndValue(self.pre_sales_table.size() * 1.1)
        table_animation.setEasingCurve(QEasingCurve.OutBounce)
        table_animation.start()

        date_layout = QHBoxLayout()
        layout.addLayout(date_layout)

        check_in_date_label = QLabel("Дата заезда:")
        check_in_date_label.setFont(font_content)
        check_in_date_label.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(check_in_date_label)

        self.check_in_date_edit = QDateEdit()
        self.check_in_date_edit.setCalendarPopup(True)
        self.check_in_date_edit.setFont(font_content)
        self.check_in_date_edit.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(self.check_in_date_edit)

        check_out_date_label = QLabel("Дата выезда:")
        check_out_date_label.setFont(font_content)
        check_out_date_label.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(check_out_date_label)

        self.check_out_date_edit = QDateEdit()
        self.check_out_date_edit.setCalendarPopup(True)
        self.check_out_date_edit.setFont(font_content)
        self.check_out_date_edit.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(self.check_out_date_edit)

        price_label = QLabel("Стоимость:")
        price_label.setFont(font_content)
        price_label.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(price_label)

        self.price_edit = QLineEdit()
        self.price_edit.setFont(font_content)
        self.price_edit.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(self.price_edit)

        booking_layout = QHBoxLayout()
        layout.addLayout(booking_layout)

        guests_label = QLabel("Количество гостей:")
        guests_label.setFont(font_content)
        guests_label.setStyleSheet("color: #0A1527;")
        booking_layout.addWidget(guests_label)

        self.guests_edit = QLineEdit()
        self.guests_edit.setFont(font_content)
        self.guests_edit.setStyleSheet("color: #0A1527;")
        booking_layout.addWidget(self.guests_edit)

        hotel_label = QLabel("Отель:")
        hotel_label.setFont(font_content)
        hotel_label.setStyleSheet("color: #0A1527;")
        booking_layout.addWidget(hotel_label)

        self.hotel_combobox = QComboBox()
        self.hotel_combobox.setFont(font_content)
        self.hotel_combobox.setStyleSheet("color: #0A1527;")
        booking_layout.addWidget(self.hotel_combobox)

        room_label = QLabel("Номер:")
        room_label.setFont(font_content)
        room_label.setStyleSheet("color: #0A1527;")
        booking_layout.addWidget(room_label)

        self.room_combobox = QComboBox()
        self.room_combobox.setFont(font_content)
        self.room_combobox.setStyleSheet("color: #0A1527;")
        booking_layout.addWidget(self.room_combobox)

        button_layout = QHBoxLayout()
        layout.addLayout(button_layout)

        book_button = QPushButton("Забронировать")
        book_button.setFont(font_content)
        book_button.setStyleSheet("color: #0A1527; background-color: #0072bc;")
        book_button.clicked.connect(self.book_room)
        button_layout.addWidget(book_button)

        hotels_data = pd.read_excel("hotels.xls")
        self.hotel_combobox.addItem("Выберите отель")
        for index, row in hotels_data.iterrows():
            hotel_name = row["Отель"]
            self.hotel_combobox.addItem(hotel_name)

        self.hotel_combobox.currentIndexChanged.connect(self.populate_room_combobox)
        self.load_data_from_excel()


    def populate_room_combobox(self):
        selected_hotel = self.hotel_combobox.currentText()

        inventory_data = pd.read_excel("inventory.xls")
        
        filtered_rooms = inventory_data[inventory_data["Отель"] == selected_hotel]["Номер"].tolist()

        self.room_combobox.clear()

        for room_number in filtered_rooms:
            self.room_combobox.addItem(str(room_number))


    def book_room(self):
        check_in_date = self.check_in_date_edit.date()
        check_out_date = self.check_out_date_edit.date()
        price = self.price_edit.text()
        guests = self.guests_edit.text()

        if not check_in_date or not check_out_date or not price or not guests:
            QMessageBox.warning(self, "Предупреждение", "Пожалуйста, заполните все поля для бронирования.")
            return

        if check_in_date >= check_out_date:
            QMessageBox.warning(self, "Предупреждение", "Дата заезда не может быть позже или равна дате выезда.")
            return

        selected_room = self.room_combobox.currentText()

        if selected_room in ["ZH2003", "ZH3001", "L10007", "L20018", "L30025", "CS210"]:
            room_guest_limit = 4
        else:
            room_guest_limit = 2

        if int(guests) > room_guest_limit:
            QMessageBox.warning(self, "Предупреждение", f"Для номера {selected_room} максимальное количество гостей - {room_guest_limit}.")
            return

        df_check_in = pd.DataFrame({
            "Дата заезда": [check_in_date.toString(Qt.ISODate)],
            "Дата выезда": [check_out_date.toString(Qt.ISODate)],
            "Отель": [self.hotel_combobox.currentText()],
            "Номер": [selected_room],
            "Список гостей": [guests]
        })
        self.append_to_excel(df_check_in, "check-in_log.xls")

        inventory_data = pd.read_excel("inventory.xls")
        inventory_data.loc[inventory_data["Номер"] == selected_room, "Статус"] = "занят"
        inventory_data.to_excel("inventory.xls", index=False)

        df_check_out = pd.DataFrame({
            "Дата выезда": [check_out_date.toString(Qt.ISODate)],
            "Отель": [self.hotel_combobox.currentText()],
            "Номер": [selected_room],
            "Список гостей": [guests]
        })
        self.append_to_excel(df_check_out, "departure_log.xls")

        row_position = self.pre_sales_table.rowCount()
        self.pre_sales_table.insertRow(row_position)
        self.pre_sales_table.setItem(row_position, 0, QTableWidgetItem(check_in_date.toString(Qt.ISODate)))
        self.pre_sales_table.setItem(row_position, 1, QTableWidgetItem(self.hotel_combobox.currentText()))
        self.pre_sales_table.setItem(row_position, 2, QTableWidgetItem(selected_room))
        self.pre_sales_table.setItem(row_position, 3, QTableWidgetItem(check_in_date.toString(Qt.ISODate)))
        self.pre_sales_table.setItem(row_position, 4, QTableWidgetItem(check_out_date.toString(Qt.ISODate)))
        self.pre_sales_table.setItem(row_position, 5, QTableWidgetItem(price))
        self.pre_sales_table.setItem(row_position, 6, QTableWidgetItem(str(guests)))
        self.save_data_to_excel()

        self.clear_booking_form()

    def append_to_excel(self, df, file_name):
        if os.path.exists(file_name):
            existing_data = pd.read_excel(file_name)
            df = pd.concat([existing_data, df], ignore_index=True)
        df.to_excel(file_name, index=False)

    def clear_booking_form(self):

        self.check_in_date_edit.clear()
        self.check_out_date_edit.clear()
        self.price_edit.clear()
        self.guests_edit.clear()

    def save_data_to_excel(self):
        data = []
        for row in range(self.pre_sales_table.rowCount()):
            row_data = []
            for column in range(self.pre_sales_table.columnCount()):
                item = self.pre_sales_table.item(row, column)
                if item is not None:
                    row_data.append(item.text())
                else:
                    row_data.append('')
            data.append(row_data)

        df = pd.DataFrame(data, columns=["Дата", "Отель", "Номер", "Дата заезда", "Дата выезда", "Стоимость", "Количество гостей"])
        df.to_excel("pre_sales.xls", index=False)
        
    def load_data_from_excel(self):
        try:
            df = pd.read_excel("pre_sales.xls")
            self.pre_sales_table.setRowCount(df.shape[0])
            for index, row in df.iterrows():
                for column in range(self.pre_sales_table.columnCount()):
                    item = QTableWidgetItem(str(row[column]))
                    self.pre_sales_table.setItem(index, column, item)
        except FileNotFoundError:
            QMessageBox.warning(self, "Предупреждение", "Файл pre_sales.xls не найден.")

    def init_sales_analysis_tab(self):
        layout = QVBoxLayout()
        self.sales_analysis_tab.setLayout(layout)

        date_layout = QHBoxLayout()
        layout.addLayout(date_layout)
        
        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18) 

        date_label = QLabel("Дата:")
        date_label.setFont(font_content)
        date_label.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(date_label)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setFont(font_content)
        self.date_edit.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(self.date_edit)

        hotel_label = QLabel("Отель:")
        hotel_label.setFont(font_content)
        hotel_label.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(hotel_label)

        self.hotel_combobox_analysis = QComboBox()
        self.hotel_combobox_analysis.setFont(font_content)
        self.hotel_combobox_analysis.setStyleSheet("color: #0A1527;")
        self.hotel_combobox_analysis.addItem("Все отели")
        hotels_data = pd.read_excel("hotels.xls")
        for index, row in hotels_data.iterrows():
            hotel_name = row["Отель"]
            self.hotel_combobox_analysis.addItem(hotel_name)
        date_layout.addWidget(self.hotel_combobox_analysis)

        calculate_button = QPushButton("Рассчитать")
        calculate_button.setFont(font_content)
        calculate_button.setStyleSheet("color: #0A1527; background-color: #0072bc;")
        calculate_button.clicked.connect(self.calculate_sales)
        layout.addWidget(calculate_button)

        self.sales_table = QTableWidget()
        self.sales_table.setColumnCount(6)
        self.sales_table.setHorizontalHeaderLabels(["Отель", "Кол-во ночей", "Сумма продаж", "% Загрузки", "Ср. цена ADR", "RevPAR"])
        layout.addWidget(self.sales_table)

    def calculate_sales(self):
        selected_date = self.date_edit.date().toPyDate()
        selected_hotel = self.hotel_combobox_analysis.currentText()

        pre_sales_data = pd.read_excel("pre_sales.xls")
        pre_sales_data["Дата заезда"] = pd.to_datetime(pre_sales_data["Дата заезда"]).dt.date
        pre_sales_data["Дата выезда"] = pd.to_datetime(pre_sales_data["Дата выезда"]).dt.date

        if selected_hotel != "Все отели":
            pre_sales_data = pre_sales_data[pre_sales_data["Отель"] == selected_hotel]

        total_rooms = len(pre_sales_data["Номер"].unique())
        total_nights_sold = 0
        total_sales_amount = 0

        for index, row in pre_sales_data.iterrows():
            if row["Дата заезда"] <= selected_date <= row["Дата выезда"]:
                nights = (row["Дата выезда"] - row["Дата заезда"]).days
                total_nights_sold += nights
                total_sales_amount += float(row["Стоимость"])

        occupancy_percentage = (total_nights_sold / (total_rooms * 365)) * 100 if total_rooms > 0 else 0
        adr = total_sales_amount / total_nights_sold if total_nights_sold > 0 else 0
        revpar = adr * occupancy_percentage / 100

        self.sales_table.setRowCount(1)
        self.sales_table.setItem(0, 0, QTableWidgetItem(selected_hotel))
        self.sales_table.setItem(0, 1, QTableWidgetItem(str(total_nights_sold)))
        self.sales_table.setItem(0, 2, QTableWidgetItem(str(total_sales_amount)))
        self.sales_table.setItem(0, 3, QTableWidgetItem("{:.2f}%".format(occupancy_percentage)))
        self.sales_table.setItem(0, 4, QTableWidgetItem("{:.2f}".format(adr)))
        self.sales_table.setItem(0, 5, QTableWidgetItem("{:.2f}".format(revpar)))

    def init_housekeeping_tab(self):
        layout = QVBoxLayout()
        self.housekeeping_tab.setLayout(layout)

        date_layout = QHBoxLayout()
        layout.addLayout(date_layout)

        font_content = QFont()
        font_content.setFamily("Arial")
        font_content.setPointSize(18)

        date_label = QLabel("Выберите дату для проверки:")
        date_label.setFont(font_content)
        date_label.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(date_label)

        self.date_for_checking = QDateEdit()
        self.date_for_checking.setCalendarPopup(True)
        self.date_for_checking.setFont(font_content)
        self.date_for_checking.setStyleSheet("color: #0A1527;")
        date_layout.addWidget(self.date_for_checking)

        check_button = QPushButton("Проверить")
        check_button.setFont(font_content)
        check_button.setStyleSheet("color: #0A1527; background-color: #0072bc;")
        check_button.clicked.connect(self.check_housekeeping_status)
        layout.addWidget(check_button)

    def check_housekeeping_status(self):
        selected_date = self.date_for_checking.date().toPyDate()

        try:
            pre_sales_data = pd.read_excel("pre_sales.xls")
        except FileNotFoundError:
            QMessageBox.warning(self, "Предупреждение", "Файл pre_sales.xls не найден.")
            return

        pre_sales_data["Дата заезда"] = pd.to_datetime(pre_sales_data["Дата заезда"]).dt.date
        pre_sales_data["Дата выезда"] = pd.to_datetime(pre_sales_data["Дата выезда"]).dt.date

        bookings_on_selected_date = pre_sales_data[(pre_sales_data["Дата заезда"] <= selected_date) & (pre_sales_data["Дата выезда"] >= selected_date)]

        if bookings_on_selected_date.empty:
            QMessageBox.information(self, "Информация", "На выбранную дату нет бронирований.")
            return

        for index, booking in bookings_on_selected_date.iterrows():
            check_out_date = booking["Дата выезда"]
            if check_out_date < selected_date:
                hotel = booking["Отель"]
                room = booking["Номер"]

                inventory_data = pd.read_excel("inventory.xls")
                inventory_data.loc[inventory_data["Номер"] == room, "Статус"] = "грязный"
                inventory_data.to_excel("inventory.xls", index=False)

                housekeeping_entry = pd.DataFrame({
                    "Дата уборки": [selected_date],
                    "Отель": [hotel],
                    "Номер": [room],
                    "ФИО горничной": ["Имя горничной"], 
                    "Уборка завершена": [False]
                })
                self.append_to_excel(housekeeping_entry, "housekeeping_log.xls")

        self.init_inventory_tab()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = HotelApp()
    window.show()
    sys.exit(app.exec_())
