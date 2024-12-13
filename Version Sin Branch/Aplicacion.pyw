from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QComboBox, QTableWidget, QTableWidgetItem, QAbstractItemView, QFileDialog, QScrollArea, QFrame, QHeaderView, QTableWidgetItem, QDialog,QTextEdit
from PySide6.QtCore import Qt
import pandas as pd
import math
import os
import sys
import requests
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt
from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect)
from PySide6.QtWidgets import QWidget, QPushButton, QVBoxLayout, QFrame, QLabel, QComboBox, QTableWidget, QScrollArea, QAbstractItemView,QMessageBox
from PySide6.QtCore import Qt, QRect 
import pandas as pd
import requests
from PySide6.QtGui import QPixmap


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Cargar Archivo")
        self.setFixedSize(600, 462)  
        self.setStyleSheet("background-color: #003c72;")
        
        # Layout principal
        self.layout = QVBoxLayout()
        
        # Centrar el título
        self.label_titulo = QLabel("Aplicacion Para Control De Inventario De Centros De Distribucion Grupo Mimesa")
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
        self.button_select_file.setStyleSheet("background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.button_select_file.clicked.connect(self.open_file_dialog)
        self.layout.addWidget(self.button_select_file, alignment=Qt.AlignCenter)

        # Centrar el botón de listo
        self.button1 = QPushButton("Listo")
        self.button1.setStyleSheet("background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.button1.clicked.connect(self.open_second_window)
        self.layout.addWidget(self.button1, alignment=Qt.AlignCenter)
        
    
        self.contenedor_imagen = QWidget()  
        self.contenedor_imagen.setGeometry(150, 200, 50, 50) 
        
        
        self.url = "https://drive.google.com/uc?export=download&id=1se96KZPovcb8BrDnaDmRHlAwKvEMito-"
        self.response = requests.get(self.url)
        if self.response.status_code == 200:
            self.img_data = self.response.content
        else:
            print("Error al descargar la imagen")


        self.pixmap = QPixmap()
        self.pixmap.loadFromData(self.img_data)
        self.imagen = QLabel()
        self.imagen.setPixmap(self.pixmap)
        resized_pixmap = self.pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    

        imagen = QLabel(self.contenedor_imagen)  
        imagen.setPixmap(resized_pixmap)    
        imagen.move(500, 500) 
 
        self.layout.addWidget(imagen)

        self.setLayout(self.layout)
        
        self.file_path = None

    def open_file_dialog(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo", "", "Archivos de Excel (*.xls *.xlsx)")
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
        msg_box = QDialog(self)
        msg_box.setWindowTitle("Advertencia")
        msg_box.setText(message)
        msg_box.exec_()





class SecondWindow(QWidget):
    def __init__(self, file_path):
        super().__init__()
        self.setWindowTitle("Aplicacion Gestion De Inventario")
        self.setFixedSize(1000, 600) 
        self.setStyleSheet("background-color: #003c72;")
        self.tm_estandar=30
        self.tm_adicional=0
        
        self.file_path = file_path
        self.df = pd.read_excel(self.file_path, skiprows=2)
        
        # Crear el QFrame
        self.tree_frame = QFrame(self)
        self.tree_frame.setStyleSheet("background-color: #003c72;")
        self.tree_frame.setGeometry(0, 0, 1000, 600)
        
        self.url = "https://drive.google.com/uc?export=download&id=1se96KZPovcb8BrDnaDmRHlAwKvEMito-"
        self.response = requests.get(self.url)
        if self.response.status_code == 200:
            self.img_data = self.response.content
        else:
            print("Error al descargar la imagen")
        self.pixmap = QPixmap()
        self.pixmap.loadFromData(self.img_data)    
        self.imagen = QLabel(self.tree_frame)
        resized_pixmap = self.pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.imagen.setPixmap(resized_pixmap)
        self.imagen.move(825, 25)
        
        # Filtrar filas donde '%Target Inv + Trans + Plan' esté vacío o NaN
        self.df = self.df[~self.df['%Target Inv + Trans + Plan'].isna()]
        self.df = self.df[~self.df['Target de Inventario'].isna()]
        self.df = self.df.fillna(0).round(2) 
        # Asegurarse de que las columnas existan en el DataFrame, si no, agregarlas
        for col in ['TM',"Nuevo %", 'Inv Final','TM Adicional','% Con la TM Adicional']:
            if col not in self.df.columns:
                self.df[col] = 0  # Si no existe, la columna se agrega con valor 0
        
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
        self.localidades_combobox.setFixedSize(150, 50)  # Fijar tamaño del ComboBox
        self.localidades_combobox.move(15, 45)
        self.localidades_combobox.currentIndexChanged.connect(self.update_table)
        
        # Botón dentro del QFrame
        self.button1 = QPushButton("Generar SJ", self.tree_frame)
        self.button1.setStyleSheet("background-color: #94cc1c; color: white; font: 14pt Arial;")
        self.button1.setFixedSize(150, 50)
        self.button1.move(825, 125)

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

        self.table = QTableWidget(self.tree_frame)
        self.table.setStyleSheet("background-color: white;") 
        self.table.setGeometry(10,185,980,400)
        
        
        self.label_resumen = QLabel("Resumen", self.tree_frame)
        self.label_resumen.setStyleSheet("color: white; font: 14pt Arial;")
        self.label_resumen.move(455, 25)

        self.resumen= QTextEdit(self.tree_frame)
        self.resumen.setStyleSheet("color: black; font: 14pt Arial; background-color: white;")
        self.resumen.setGeometry(374, 55, 250, 100)
        
    
        
        # Inicializa la tabla
        self.columns_to_display = [col for col in self.df.columns if col not in ['%Target Inv', '%Target Inv + Trans']]
        self.table.setColumnCount(len(self.columns_to_display))
        self.table.setHorizontalHeaderLabels(self.columns_to_display)
        

        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.update_table()
        self.table.cellChanged.connect(self.calculo_Manual)
        self.button1.clicked.connect(self.guardarSj)

    def update_table(self):
        
        selected_localidad = self.localidades_combobox.currentText()
        selected_categoria = self.categorias_combobox.currentText()
        filtered_df = self.df.copy()
        if selected_localidad:
            filtered_df = filtered_df[filtered_df['Localidad'] == selected_localidad]
        if selected_categoria:
            filtered_df = filtered_df[filtered_df['Categoria'] == selected_categoria]
        filtered_df['Nuevo %'] = (filtered_df['%Target Inv + Trans + Plan'] * 100).round(2)
        for _ in range(30):  # 30 iteraciones para incrementar 'Paletas'
            if filtered_df.empty:
                break  # Si el DataFrame está vacío, salimos del bucle
            # Obtenemos el índice de la fila con el valor mínimo de 'Nuevo %'
            min_value_row_idx = filtered_df['Nuevo %'].idxmin()
            # Si no se encuentra un índice válido, salimos del bucle
            if min_value_row_idx is None:
                break  # Si no hay índice mínimo, terminamos el bucle      
            filtered_df.at[min_value_row_idx, 'TM'] += 1           
            # Obtenemos la fila correspondiente al índice mínimo
            min_value_row = filtered_df.loc[min_value_row_idx]
            transit_tm = min_value_row['Tránsito TM']
            planned_tm = min_value_row['Planificado TM']
            inv_exist_tm = min_value_row['Inv Exist TM']
            target_inv = min_value_row['Target de Inventario']
            paletas_inv= min_value_row['TM']

            nuevo_percent = ((transit_tm + planned_tm + inv_exist_tm+paletas_inv) / target_inv ) * 100
            
            invfinal=paletas_inv+inv_exist_tm
            
            filtered_df.at[min_value_row_idx, 'Nuevo %'] = nuevo_percent
            filtered_df['Inv Final'] = filtered_df['Inv Final'].astype(float)
            filtered_df.at[min_value_row_idx, 'Inv Final'] = invfinal
        self.table.setRowCount(filtered_df.shape[0])
        for row_idx, (_, row) in enumerate(filtered_df.iterrows()):
            for col_idx, col in enumerate(self.columns_to_display):
                value = row[col]
         
                if col in ['Inv Final']:
                    value= round(value,2)
                
                if col in ['%Target Inv', '%Target Inv + Trans + Plan']:
                    value = self.porcentaje_a_float(value)
                    value = round(value * 100, 2)
                    value = f"{value:.2f}%"
                    
                    
                if col in ['Nuevo %']:
                    value = self.porcentaje_a_float(value)
                    value = round(value, 2)
                    value = f"{value:.2f}%"

                item = QTableWidgetItem(str(value))
                self.table.setItem(row_idx, col_idx, item)
                
        self.tm_adicional=0
         
        self.resumen.setText(f"Total TM a enviar: {self.tm_estandar}")
        
        return self.df
 
    def calculo_Manual(self, row2, column):  
        if(column == 13):
            transit_tm = float(self.table.item(row2,7).text())
            planned_tm = float(self.table.item(row2,6).text())
            inv_exist_tm =float(self.table.item(row2,5).text())
            target_inv = float(self.table.item(row2,8).text())
            paletas_inv=float(self.table.item(row2,10).text()) 
            paletas_Agregada=float(self.table.item(row2,13).text())
            valor=((transit_tm + planned_tm + inv_exist_tm+paletas_inv+paletas_Agregada) / target_inv ) * 100
            valor = self.porcentaje_a_float(valor)
            valor = round(valor, 2)
            valor = f"{valor:.2f}%"
            item=QTableWidgetItem(valor)
            self.table.setItem(row2,14,item)         
            self.tm_adicional=self.tm_adicional+paletas_Agregada
            self.resumen.setText(f"Total TM a enviar: {self.tm_estandar+self.tm_adicional}")        
        else: 
            pass

    def guardarSj(self):
        try:
            selected_localidad = self.localidades_combobox.currentText().lower()
            selected_categoria = self.categorias_combobox.currentText().lower()
            
            column_names = [self.table.horizontalHeaderItem(i).text().lower() for i in range(self.table.columnCount())]
            
            if 'localidad' in column_names and 'categoria' in column_names:
                localidad_index = column_names.index('localidad')
                categoria_index = column_names.index('categoria')
                
                filtered_data = []
                for row in range(self.table.rowCount()):
                    item_localidad = self.table.item(row, localidad_index)
                    item_categoria = self.table.item(row, categoria_index)
                    
                    if item_localidad and item_categoria:
                        localidad_match = selected_localidad in item_localidad.text().lower()
                        categoria_match = selected_categoria in item_categoria.text().lower()
                        
                        if localidad_match and categoria_match:
                            self.table.setRowHidden(row, False)
                            row_data = []
                            suma_col10_col13 = 0
                            # Solo agregar las columnas 3 y 10
                            for col in [3, 10]:
                                item = self.table.item(row, col)
                                if item:
                                    text = item.text()
                                    if col == 3:
                                        text = text.split(' ')[0] 
                                    elif col == 10:
                                        item_col10 = self.table.item(row, 10)
                                        item_col13 = self.table.item(row, 13)
                                        if item_col10 and item_col13:
                                            suma_col10_col13 = float(item_col10.text()) + float(item_col13.text())
                                            text = str(suma_col10_col13)
                                        else:
                                            text = '0'
                                    row_data.append(text)
                                else:
                                    row_data.append('')
                            # Agregar la columna "Precio total" con valor 1
                            row_data.append('1')
                            if suma_col10_col13 != 0:
                                filtered_data.append(row_data)
                        else:
                            self.table.setRowHidden(row, True)
                
                # Guardar solo las columnas necesarias en un archivo Excel
                columns = ['Sku', 'TM', 'Precio total']  # Asignar los nombres de las columnas adecuadas
                df = pd.DataFrame(filtered_data, columns=columns)
                file_name = f'Subida_Sj_{selected_localidad}_{selected_categoria}.xlsx'
                df.to_excel(file_name, index=False)
                self.mostrarMensaje("Información", f"El archivo '{file_name}' fue creado con éxito.")
        except Exception as e:
            self.mostrarMensaje("Error", f"Se produjo un error: {e}", tipo='error')


    
    def porcentaje_a_float(self, porcentaje_str):
        """Convierte un string como '10.00%' o un número flotante a un valor flotante."""
        if isinstance(porcentaje_str, str): 
            return float(porcentaje_str.replace('%', '').strip()) / 100.0
        elif isinstance(porcentaje_str, float): 
            return porcentaje_str
        else:
            raise ValueError(f"Valor inesperado: {porcentaje_str}")
        
        
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
quiero modificar la crecion de este excel, solo voy a necesitar que se guarden la columna 3 y la 10 como se estan calculando 
en la forma actual, mas agregar una columna que el nombre sea "Precio total" y esta tendra en todas las filas un 1 

"""
