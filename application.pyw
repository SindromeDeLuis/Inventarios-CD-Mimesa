from PySide6.QtCore import Qt, QCoreApplication, QDate, QDateTime, QLocale, QMetaObject, QObject, QPoint, QRect
from PySide6.QtGui import QColor, QPixmap, QGuiApplication
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QComboBox,
                               QTableWidget, QTableWidgetItem, QAbstractItemView, QFileDialog, QScrollArea,
                               QFrame, QHeaderView, QDialog, QTextEdit, QLineEdit, QMessageBox, QScrollBar)
import pandas as pd
import math
import os
import sys
import requests
from datetime import datetime


def get_image():
    global img_data
    url = "https://github.com/SindromeDeLuis/Inventarios-CD-Mimesa/blob/main/assets/GM%20Logo-2.png?raw=true"
    response = requests.get(url)
    if response.status_code == 200:
        img_data = response.content
    else:
        img_data = None
        print("Error al descargar la imagen")


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cargar Archivo")
        self.setFixedSize(600, 462)
        self.setStyleSheet("background-color: #003c72;")
        # Layout principal
        self.layout = QVBoxLayout()
        # Centrar el título
        self.label_titulo = QLabel(
            "Aplicacion Para Control De Inventario De Centros De Distribucion Grupo Mimesa")
        self.label_titulo.setStyleSheet("color: white; font: 12pt Arial;")
        self.label_titulo.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.label_titulo)
        # Centrar el status label
        self.status_label = QLabel("No se ha seleccionado ningún archivo.")
        self.status_label.setStyleSheet("color: white; font: 12pt Arial;")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.status_label)
        # Centrar el botón de seleccionar archivo
        self.button_select_file = QPushButton("Seleccionar archivo")
        self.button_select_file.setStyleSheet(
            "background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.button_select_file.clicked.connect(self.open_file_dialog)
        self.layout.addWidget(self.button_select_file,
                              alignment=Qt.AlignCenter)
        # Centrar el botón de listo
        self.button1 = QPushButton("Listo")
        self.button1.setStyleSheet(
            "background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.button1.clicked.connect(self.open_second_window)
        self.layout.addWidget(self.button1, alignment=Qt.AlignCenter)
        # Crear imagen logo
        self.contenedor_imagen = QWidget()
        self.contenedor_imagen.setGeometry(150, 200, 50, 50)
        get_image()
        self.pixmap = QPixmap()
        self.pixmap.loadFromData(img_data)
        self.imagen = QLabel()
        self.imagen.setPixmap(self.pixmap)
        resized_pixmap = self.pixmap.scaled(
            150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        imagen = QLabel(self.contenedor_imagen)
        imagen.setPixmap(resized_pixmap)
        imagen.move(500, 500)
        self.layout.addWidget(imagen)
        self.setLayout(self.layout)
        self.file_path = None

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo", "", "Archivos de Excel (*.xls *.xlsx)")
        if file_path:
            self.file_path = file_path
            file_name = os.path.basename(self.file_path)
            self.status_label.setText(f"Archivo seleccionado: {file_name}")

    def open_second_window(self):
        if not self.file_path:
            self.show_warning("¡Por favor, selecciona un archivo primero!")
            return

        self.second_window = SecondWindow(self.file_path)
        self.second_window.show()

    def show_warning(self, message):
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Advertencia")
        msg_box.setText(message)
        msg_box.exec_()


class SecondWindow(QWidget, QApplication):
    def __init__(self, file_path):
        super().__init__()
        self.setWindowTitle("Gestion de Inventario")
        # SI SE QUIERE CAMBIAR LA CANTIDAD DE PALETAS POR DEFECTO CAMBIAR AQUI, POR DEFECTO 30
        self.cantidad_de_paletas_a_enviar = 30
        self.iteraciones = 0
        width, height = self.screens()[0].size().toTuple()
        height = height-70
        self.setFixedSize(width, height)  # (1200, 600)
        self.move(0, 0)
        self.setStyleSheet("background-color: #003c72;")
        self.tm_estandar = 30
        self.tm_adicional = 0
        self.file_path = file_path
        # Crear el QFrame imagen logo
        self.tree_frame = QFrame(self)
        self.tree_frame.setStyleSheet("background-color: #003c72;")
        self.tree_frame.setGeometry(0, 0, width, height)
        self.pixmap = QPixmap()
        self.pixmap.loadFromData(img_data)
        self.imagen = QLabel(self.tree_frame)
        resized_pixmap = self.pixmap.scaled(
            180, 180, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.imagen.setPixmap(resized_pixmap)
        self.imagen.move(width-180-20, 0)  # (1100, 25)

        self.set_dataframe()

        if 'Localidad' not in self.df.columns or 'Categoria' not in self.df.columns:
            return
        self.localidades_unicas = self.df['Localidad'].unique().tolist()
        self.categorias_unicas = self.df['Categoria'].unique().tolist()

        # Localidad ComboBox
        self.label_localidad = QLabel("Localidad", self.tree_frame)
        self.label_localidad.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_localidad.move(25, 25)
        self.localidades_combobox = QComboBox(self.tree_frame)
        self.localidades_combobox.addItems(self.localidades_unicas)
        self.localidades_combobox.setStyleSheet("""
            QComboBox {
                background-color: #94cc1c;
                border: 1px solid #ccc;
                padding: 5px;
                margin: 10px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox QAbstractItemView {
                background-color: #1c3c6b;
                border: 1px solid #ccc;
                padding: 5px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox::drop-down {
                border: 1px solid #1c3c6b;
                background-color: #1c3c6b;
            }
        """)
        self.localidades_combobox.setFixedSize(
            150, 50)  # Fijar tamaño del ComboBox
        self.localidades_combobox.move(15, 45)
        self.localidades_combobox.currentIndexChanged.connect(
            self.update_table)
        # Localidad OrdenarPor
        self.label_OrdenarPor = QLabel("Ordenar Por", self.tree_frame)
        self.label_OrdenarPor.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_OrdenarPor.move(200, 27)
        self.OrdenarPor_combobox = QComboBox(self.tree_frame)
        self.OrdenarPor_combobox.addItems(["", "Código + Descripción del producto a despachar", "Inv en Origen TM", "Inv en Destino TM", "Planificado TM", "Tránsito TM",
                                           "% Target Original", "Target de Inventario", "Paletas Sugeridas", "Nuevo % Simulado", "Corr. Paletas", "Inv Final Simulado", "% Con Corrección"])
        self.OrdenarPor_combobox.setStyleSheet("""
            QComboBox {
                background-color: #94cc1c;
                border: 1px solid #ccc;
                padding: 5px;
                margin: 10px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox QAbstractItemView {
                background-color: #1c3c6b;
                border: 1px solid #ccc;
                padding: 5px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox::drop-down {
                border: 1px solid #1c3c6b;
                background-color: #1c3c6b;
            }
        """)
        self.OrdenarPor_combobox.setFixedSize(
            200, 50)  # Fijar tamaño del ComboBox
        self.OrdenarPor_combobox.move(190, 45)
        self.OrdenarPor_combobox.currentIndexChanged.connect(self.update_table)

        # label Nivel
        self.label_Nivel = QLabel("Nivel", self.tree_frame)
        self.label_Nivel.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_Nivel.move(200, 100)
        self.Nivel_combobox = QComboBox(self.tree_frame)
        self.Nivel_combobox.addItems(["", "Mayor", "Menor"])
        self.Nivel_combobox.setStyleSheet("""
            QComboBox {
                background-color: #94cc1c;
                border: 1px solid #ccc;
                padding: 5px;
                margin: 10px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox QAbstractItemView {
                background-color: #1c3c6b;
                border: 1px solid #ccc;
                padding: 5px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox::drop-down {
                border: 1px solid #1c3c6b;
                background-color: #1c3c6b;
            }
        """)
        self.Nivel_combobox.setFixedSize(150, 50)  # Fijar tamaño del ComboBox
        self.Nivel_combobox.move(190, 125)
        self.Nivel_combobox.currentIndexChanged.connect(self.update_table)

        # Generar SJ Botón dentro del QFrame
        self.button1 = QPushButton("Generar Propuesta de SJ", self.tree_frame)
        self.button1.setStyleSheet(
            "background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.button1.setFixedSize(250, 55)
        self.button1.move(width-250-25, 110)

        # Reiniciar proceso Botón dentro del QFrame
        self.reset_table_button = QPushButton("Reiniciar", self.tree_frame)
        self.reset_table_button.setStyleSheet(
            "background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.reset_table_button.setFixedSize(100, 32)
        self.reset_table_button.move(460, 132)

        # Iteraciones Label
        self.label_Nivel = QLabel(
            f"Iteraciones: {self.iteraciones}", self.tree_frame)
        self.label_Nivel.setStyleSheet("color: white; font: 12pt Arial;")
        self.label_Nivel.move(650, 160)

        # Categoria ComboBox
        self.label_categoria = QLabel("Categoria", self.tree_frame)
        self.label_categoria.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_categoria.move(25, 100)
        self.categorias_combobox = QComboBox(self.tree_frame)
        self.categorias_combobox.addItems(self.categorias_unicas)
        self.categorias_combobox.setStyleSheet("""
            QComboBox {
                background-color: #94cc1c;
                border: 1px solid #ccc;
                padding: 5px;
                margin: 10px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox QAbstractItemView {
                background-color: #1c3c6b;
                border: 1px solid #ccc;
                padding: 5px;
                color: white;
                font-family: Arial;
                font-size: 14px;
            }
            QComboBox::drop-down {
                border: 1px solid #1c3c6b;
                background-color: #1c3c6b;
            }
        """)
        self.categorias_combobox.setFixedSize(150, 50)
        self.categorias_combobox.move(15, 125)
        self.categorias_combobox.currentIndexChanged.connect(self.update_table)

        # Definir La tabla
        self.table = QTableWidget(self.tree_frame)
        self.table.setStyleSheet("background-color: white;")
        self.table.setGeometry(10, 185, width-30, height-220)

        # Definir label Resumen
        self.label_resumen = QLabel("Resumen", self.tree_frame)
        self.label_resumen.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_resumen.move(730, 25)

        # Definir TextField Resumen
        self.resumen = QTextEdit(self.tree_frame)
        self.resumen.setStyleSheet(
            "color: black; font: 14pt Arial; background-color: white;")
        self.resumen.setGeometry(650, 55, 250, 100)

        # Definir label paletas
        self.label_paletas = QLabel("Cantidad de Paletas", self.tree_frame)
        self.label_paletas.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_paletas.move(425, 75)

        # Definir TextFiel de paletas
        self.line_edit = QLineEdit(self.tree_frame)
        self.line_edit.setText(str(self.cantidad_de_paletas_a_enviar))
        self.line_edit.setStyleSheet(
            "color: black; font: 14pt Arial; background-color: white; text-align:center;")
        self.line_edit.setGeometry(460, 100, 100, 25)

        # Rellenar tabla con las columnas
        self.columns_to_display = [col for col in self.df.columns if col not in [
            '%Target Inv', '%Target Inv + Trans']]
        self.table.setColumnCount(len(self.columns_to_display))
        self.table.setHorizontalHeaderLabels(self.columns_to_display)
        self.table.resizeColumnsToContents()
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        # LLamar la funcion que calcula todo
        self.update_table()
        # Declara a que acciones estaran pendiende los botones y Textfield
        self.table.cellChanged.connect(self.calculo_Manual)
        self.button1.clicked.connect(self.guardarSj)
        self.line_edit.textChanged.connect(self.update_table)
        self.reset_table_button.clicked.connect(self.reset_table)

    def set_dataframe(self):
        # Cargar el archivo Excel
        self.df = pd.read_excel(self.file_path, skiprows=2)

        # Definir las nuevas columnas
        self.new_columns = [
            'Branchplant Origen', 'Branchplant Destino', 'Localidad', 'Categoria', 'FAMILIA 3',
            'Código + Descripción del producto a despachar', 'UOM Prim', 'Factor TM/PL', 'Factor Prim/PL', 'Inv en Origen TM',
            'Inv en Destino TM', 'Planificado TM', 'Tránsito TM', 'Target de Inventario',
            '% Target Original',
        ]

        # Renombrar las columnas existentes para que coincidan con las nuevas
        self.df.rename(columns={
            'Max of Factor PL - TM': 'Factor TM/PL',
            'Max of Factor PL - Prim': 'Factor Prim/PL',
            '%Target Inv + Trans + Plan (Python)': '% Target Original',
            'Inv Exist TM': 'Inv en Destino TM',
            'Cod + Descr': 'Código + Descripción del producto a despachar'
        }, inplace=True)

        # Reordenar las columnas según el nuevo formato
        self.df = self.df[self.new_columns]
        self.df = self.df[~self.df['% Target Original'].isna()]
        self.df = self.df[~self.df['Target de Inventario'].isna()]
        self.df = self.df.fillna(0)

        # Asegurarse de que las columnas existan en el DataFrame, si no, agregarlas
        for col in ['Código + Descripción del producto a despachar duplicado 1', 'Inv en Origen TM dupli', "Inv Final en Origen TM", 'Paletas Sugeridas', 'Nuevo % Simulado',
                    'Corr. Paletas', 'Inv Final Simulado', '% Con Corrección', 'Código + Descripción del producto a despachar duplicado 2',
                    '% Target Original dupli', '% Con Corrección dupli']:
            if col not in self.df.columns:
                # Si no existe, la columna se agrega con valor 0
                self.df[col] = 0

        columns_to_round = ['Inv en Origen TM', 'Target de Inventario',
                            'Inv en Origen TM dupli', 'Inv en Destino TM', 'Planificado TM', 'Tránsito TM']
        self.df[columns_to_round] = self.df[columns_to_round].round(5)

    def update_table(self):

        # Se Toman los datos de filtrado que estan en los comboxs y se hace una copia
        selected_localidad = self.localidades_combobox.currentText()
        selected_categoria = self.categorias_combobox.currentText()
        ordenar_por = self.OrdenarPor_combobox.currentText()
        nivel = self.Nivel_combobox.currentText()
        filtered_df = self.df.copy()

        # Filtrar el DataFrame
        if selected_localidad:
            filtered_df = filtered_df[filtered_df['Localidad']
                                      == selected_localidad]
        if selected_categoria:
            filtered_df = filtered_df[filtered_df['Categoria']
                                      == selected_categoria]

        # Se duplican los valores de algunas colunmas como los duplicados o % nuevo simulado con Targer original
        # Para poder hacer la primera iteracion
        filtered_df['Nuevo % Simulado'] = filtered_df['% Target Original']
        filtered_df['Código + Descripción del producto a despachar duplicado 1'] = filtered_df['Código + Descripción del producto a despachar']
        filtered_df['Código + Descripción del producto a despachar duplicado 2'] = filtered_df['Código + Descripción del producto a despachar']
        filtered_df['% Target Original dupli'] = filtered_df['% Target Original']
        filtered_df['Inv en Origen TM dupli'] = filtered_df['Inv en Origen TM']
        filtered_df['Inv Final en Origen TM'] = filtered_df['Inv en Origen TM']

        # palabra de control para el ciclo
        paletas_agregadas = 0
        # Lista de pedidos que no pueden despacharle nada por falta de inventario
        procesados = set()
        # Duplicado de la tabla menos las filas que no se pueden despachar por falta d inventario
        df_no_procesados = filtered_df.drop(procesados)
        # Ciclo infito hasta que paletas_agregadas sea igual a lo seleccionado en el TextField de cuantas palabras
        # Se quieren Agregar
        while paletas_agregadas < int(self.line_edit.text()):
            if filtered_df.empty:
                break

            # Filtrar las filas ya procesadas
            if df_no_procesados.empty:
                break

            # Se toma la fila con menor Nuevo % Simulado
            min_value_row_idx = df_no_procesados['Nuevo % Simulado'].idxmin()
            if min_value_row_idx is None:
                break

            # Se extren datos necesitarios para la validacion, factor de pl a tm, inv en origen,target
            factor_tm_pl = float(
                df_no_procesados.at[min_value_row_idx, 'Factor TM/PL'])
            inv_en_origen_tm = float(
                df_no_procesados.at[min_value_row_idx, 'Inv Final en Origen TM'])
            target = float(
                df_no_procesados.at[min_value_row_idx, 'Target de Inventario'])

            # 1r Condición de la validación Si factor de pl a tm es menor a origen tm
            # 2r Condición que inv en origen no se 0 o menor a 0
            # 3r que target sea diferente de 0
            if factor_tm_pl <= inv_en_origen_tm and inv_en_origen_tm > 0 and target != 0:
                filtered_df.at[min_value_row_idx, 'Paletas Sugeridas'] += 1
                # Si todas las condiciones se cumple se agrega 1 paleta
                paletas_agregadas += 1
            else:
                procesados.add(min_value_row_idx)

            # Realizar los cálculos para todas las filas
            for idx, row in filtered_df.iterrows():
                # Se sacan los datos de la tabla
                transit_tm = float(row['Tránsito TM'])
                planned_tm = float(row['Planificado TM'])
                inv_exist_tm = float(row['Inv en Destino TM'])
                target_inv = float(row['Target de Inventario'])
                paletas_inv = float(row['Paletas Sugeridas'])
                Factor_Conversion = float(row["Factor TM/PL"])
                origen = float(row["Inv en Origen TM"])

                try:
                    # Se calcula TargetOriginal
                    TargetOriginal = (
                        (transit_tm + planned_tm + inv_exist_tm) / target_inv) * 100

                    Paleta_a_Tm = Factor_Conversion * paletas_inv
                    # Nuevo Porcentaje %
                    nuevo_percent = (
                        (transit_tm + planned_tm + inv_exist_tm + Paleta_a_Tm) / target_inv) * 100
                    invfinal = Paleta_a_Tm + inv_exist_tm + transit_tm + planned_tm
                    # Se actualiza los nuevos valores
                    filtered_df.at[idx, 'Nuevo % Simulado'] = nuevo_percent
                    filtered_df.at[idx, 'Inv Final Simulado'] = invfinal
                    filtered_df.at[idx, '% Con Corrección'] = nuevo_percent
                    filtered_df.at[idx,
                                   '% Con Corrección dupli'] = nuevo_percent
                    filtered_df.at[idx, '% Target Original'] = TargetOriginal
                    filtered_df.at[idx, 'Inv Final en Origen TM'] = round(
                        origen-Paleta_a_Tm, 5)
                    filtered_df.at[idx,
                                   '% Target Original dupli'] = TargetOriginal
                except:
                    pass
                # Se vuelve a hacer la copia de la tabla con las filas que no cumplen eliminadas
                df_no_procesados = filtered_df.drop(procesados)
        # Se aplican los filtros
        if ordenar_por and nivel:
            ascending = True if nivel == "Menor" else False
            filtered_df = filtered_df.sort_values(
                by=ordenar_por, ascending=ascending)
        # Se parsea y se la agrega a los campos necesario el %
        self.table.setRowCount(filtered_df.shape[0])
        for row_idx, (_, row) in enumerate(filtered_df.iterrows()):
            for col_idx, col in enumerate(self.columns_to_display):
                value = row[col]
                if col in ['Inv Final Simulado']:
                    value = round(value, 5)
                if col in ['Nuevo % Simulado', '% Con Corrección', '% Con Corrección dupli', '% Target Original', '% Target Original dupli']:
                    value = round(float(value), 5)
                    value = f"{value:.2f}%"

                # Se carga en la tabla y se colorea
                item = QTableWidgetItem(str(value))
                self.table.setItem(row_idx, col_idx, item)
                if col_idx in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]:  # Gris
                    item.setBackground(QColor(169, 169, 169))  # Gris
                elif col_idx in range(15, 23):  # Azul
                    item.setBackground(QColor(173, 216, 230))  # Azul claro
                elif col_idx in range(23, 26):  # Verde
                    item.setBackground(QColor(144, 238, 144))  # Verde claro
                self.table.setItem(row_idx, col_idx, item)

        self.tm_adicional = 0
        self.resumen.setText(f"Total paletas a enviar: {paletas_agregadas}")

        # Bloquear edición en todas las columnas excepto la columna 20
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if col != 20:  # Bloquear edición en todas las columnas excepto la columna 20
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)

        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)

        self.label_Nivel.setText(f"Iteraciones: {self.iteraciones}")

    def calculo_Manual(self, row, column):
        if (column == 20):
            transit_tm = float(self.table.item(row, 12).text())
            planned_tm = float(self.table.item(row, 11).text())
            inv_exist_tm = float(self.table.item(row, 10).text())
            target_inv = float(self.table.item(row, 13).text())
            paletas_inv = float(self.table.item(row, 18).text())
            paletas_Agregada = float(self.table.item(row, 20).text())
            Factor_Conversion = float(self.table.item(row, 7).text())
            Paleta_A_TM = Factor_Conversion*(paletas_Agregada+paletas_inv)
            origen = float(self.table.item(row, 9).text())

            origenfinal = round(origen-Paleta_A_TM, 5)
            item3 = QTableWidgetItem(str(origenfinal))
            self.table.setItem(row, 17, item3)
            item3.setBackground(QColor(173, 216, 230))

            valor = ((transit_tm + planned_tm + inv_exist_tm +
                     Paleta_A_TM) / target_inv) * 100
            valor = round(valor, 2)

            inv_final2 = round(Paleta_A_TM+transit_tm +
                               planned_tm+inv_exist_tm, 2)
            item2 = QTableWidgetItem(str(inv_final2))
            self.table.setItem(row, 21, item2)
            item2.setBackground(QColor(173, 216, 230))  # Azul claro
            valor = f"{valor:.2f}%"
            item = QTableWidgetItem(valor)
            item3 = QTableWidgetItem(valor)
            self.table.setItem(row, 22, item)
            item.setBackground(QColor(173, 216, 230))  # Azul claro
            self.table.setItem(row, 25, item3)
            item3.setBackground(QColor(144, 238, 144))  # Verde claro

            self.suma_total = 0
            for fila in range(self.table.rowCount()):
                item = self.table.item(fila, 20)
                if item is not None:
                    self.suma_total += float(item.text())
            self.resumen.setText(f"Total paletas a enviar: {
                                 self.tm_estandar+self.suma_total}")
        else:
            pass

    def guardarSj(self):
        try:
            selected_localidad = self.localidades_combobox.currentText().lower()
            selected_categoria = self.categorias_combobox.currentText().lower()

            if not selected_localidad or not selected_categoria:
                self.mostrarMensaje(
                    "Error", "Debe seleccionar una localidad y una categoría.", tipo='error')
                return

            column_names = [self.table.horizontalHeaderItem(
                i).text().lower() for i in range(self.table.columnCount())]

            if 'localidad' not in column_names or 'categoria' not in column_names:
                self.mostrarMensaje(
                    "Error", "Las columnas 'localidad' y 'categoria' no están presentes en la tabla.", tipo='error')
                return

            localidad_index = column_names.index('localidad')
            categoria_index = column_names.index('categoria')
            tm_index = column_names.index('paletas sugeridas')

            filtered_data = []
            for row in range(self.table.rowCount()):
                item_localidad = self.table.item(row, localidad_index)
                item_categoria = self.table.item(row, categoria_index)
                item_tm = self.table.item(row, tm_index)

                if item_localidad and item_categoria and item_tm:
                    localidad_match = selected_localidad in item_localidad.text().lower()
                    categoria_match = selected_categoria in item_categoria.text().lower()
                    tm_value = float(item_tm.text())

                    if localidad_match and categoria_match and tm_value != 0:
                        self.table.setRowHidden(row, False)
                        row_data = []
                        for col in [5, 16, 6, 0, 1]:
                            item = self.table.item(row, col)
                            if item:
                                text = item.text()
                                if col == 3:
                                    text = text.split(' ')[0]
                                elif col == 16:
                                    item_col15 = self.table.item(row, 18)
                                    item_col14 = self.table.item(row, 8)
                                    item_col18 = self.table.item(row, 20)
                                    if item_col18 and item_col14:
                                        valor = (
                                            float(item_col15.text()) + float(item_col18.text()))
                                        valor2 = int(
                                            valor * float(item_col14.text()))
                                        valor3 = round(valor2, 2)
                                        text = str(valor3)
                                    else:
                                        text = '0'
                                row_data.append(text)
                            else:
                                row_data.append('')
                        row_data.insert(2, '1')
                        filtered_data.append(row_data)
                    else:
                        pass
                for idx, row2 in self.df.iterrows():
                    if row2['Código + Descripción del producto a despachar'] == self.table.item(row, 5).text():
                        self.df.at[idx, 'Inv en Origen TM'] = float(
                            self.table.item(row, 17).text())
                        if row2['Localidad'].lower() == selected_localidad:
                            paletas_inv = float(
                                self.table.item(row, 18).text())
                            paletas_Agregada = float(
                                self.table.item(row, 20).text())
                            Factor_Conversion = float(
                                self.table.item(row, 7).text())
                            Paleta_A_TM = Factor_Conversion * \
                                (paletas_Agregada+paletas_inv)
                            self.df.at[idx, 'Planificado TM'] += Paleta_A_TM

            if not filtered_data:
                self.mostrarMensaje(
                    "Información", "No se encontraron datos que coincidan con los criterios seleccionados.")
                return

            columns = ['Sku', 'Cantidad', 'Precio total', 'UOM Prim',
                       'Branchplant Origen', 'Branchplant Destino']
            df = pd.DataFrame(filtered_data, columns=columns)
            timestamp = datetime.now().strftime('%Y%m%d%H%M')
            file_name = f'Subida_Sj_{selected_localidad}_{
                selected_categoria}_{timestamp}.xlsx'
            script_dir = os.path.dirname(os.path.abspath(__file__))
            new_file_path = os.path.join(script_dir, file_name)
            df.to_excel(new_file_path, index=False)
            self.mostrarMensaje("Información", f"El archivo '{
                                file_name}' fue creado con éxito.")
            self.iteraciones += 1
            self.update_table()
        except Exception as e:
            self.mostrarMensaje(
                "Error", f"Se produjo un error: {e}", tipo='error')

    def reset_table(self):
        self.df = pd.read_excel(self.file_path, skiprows=2)
        self.line_edit.setText(str(self.cantidad_de_paletas_a_enviar))
        self.iteraciones = 0
        self.set_dataframe()
        self.update_table()

    def porcentaje_a_float(self, porcentaje_str):
        try:
            # Verifica si el valor es un porcentaje válido
            if isinstance(porcentaje_str, str) and porcentaje_str.endswith('%'):
                return float(porcentaje_str.strip('%')) / 100
            # Si el valor ya es un número, simplemente devuélvelo
            elif isinstance(porcentaje_str, (int, float)):
                return float(porcentaje_str)
            else:
                raise ValueError(f"Valor inesperado: {porcentaje_str}")
        except ValueError as e:
            print(f"Error al convertir porcentaje: {e}")
            return 0.0  # O cualquier valor por defecto que consideres apropiado

    def mostrarMensaje(self, titulo, mensaje, tipo='informacion'):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(titulo)
        msg_box.setText(mensaje)

        # Aplicar estilo CSS para cambiar el color de fondo
        msg_box.setStyleSheet("QMessageBox { background-color: white; }")

        if tipo == 'informacion':
            msg_box.setIcon(QMessageBox.Information)
        elif tipo == 'error':
            msg_box.setIcon(QMessageBox.Critical)

        msg_box.exec_()


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

"""
formado app
0 Branchplant Origen
1 Branchplant Destino
2 Localidad
3 Categoria
4 FAMILIA 3
5 Cod + Descr
6 UOM Prim
7 Factor TM/PL
8 Factor Prim/PL
9 Inv en Origen TM
10 Inv en Destino TM
11 Planificado TM
12 Tránsito TM
13 Target de Inventario
14 % Target Original 
15 Cod + Descr dupli
16 Inv en Origen TM ref
17 Inv en Origen Final TM
18 Paletas Sugeridas 
19 Nuevo % Simulado 
20 Corrección de Paletas
21 Inv Final Simulado
22 % Con Corrección
23 Cod + Descr dupli
24 % Target Original
25 % Con Corrección dupli
"""

"""
archivo de excel original
0 Localidad
1 Categoria
2 FAMILIA 3
3 Cod + Descr
4 Branchplant Origen
5 Branchplant Destino
6 UOM Prim
7 Inv en Origen TM
8 Inv Exist TM
9 Planificado TM
10 Tránsito TM
11 Target de Inventario
12 %Target Inv + Trans + Plan (Python)
13 Max of Factor PL - TM
14 Max of Factor PL - Prim
"""
