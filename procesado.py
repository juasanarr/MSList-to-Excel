import csv
import pandas as pd
import os

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit


eliminarPuntoComa = lambda str : str[:str.index(";")]

def encontrar_archivo(nombre, ruta_base):
    for dirpath, _, filenames in os.walk(ruta_base):
        if nombre in filenames:
            return os.path.join(dirpath, nombre)

def encuentraCaracter(s):
    while True:
        if s[0] == " ":
            s = s[1:]
        else:
            return s 

def correccion(s):
    res = ''
    mayus = False
    for c in s:
        if c == ",":
            mayus = True
        elif c == '"':
            mayus = False
        elif c != " " and mayus:
            c = str(c).lower()
            mayus = False
        res += c
    return res

def unirCorchetes(ls):
    guarda = False
    res = []
    for s in ls:
        if "[" in s:
            s = s.replace("[", "")
            guarda = True
        elif guarda:
            if "]" in s:
                s = s.replace("]", "")
                guarda = False
            res[-1] += ", " + s
            continue
        res.append(s)
    return res

def crear_excel(nombreCSV, nombreExcel):
    with open(nombreCSV, "r", newline="", encoding="latin-1") as brio:
        reader = csv.reader(brio)
        next(reader)
        for i, row in enumerate(reader):
            nrow = correccion(row[0]).split(",")[1:]
            if nrow: 
                if ";" in nrow[-1]:
                    nrow[-1] = eliminarPuntoComa(nrow[-1])
                j = 0
                res = []
                while j < len(nrow):
                    if nrow[j] == "":
                        res.append(" ")
                    elif "{" in nrow[j] or "}" in nrow[j] or nrow[j] == ' "':
                        pass 
                    elif res and (len(nrow[j]) > 1 and (not (encuentraCaracter(nrow[j])[0].isupper() or encuentraCaracter(nrow[j])[0] == '"') and "@" not in nrow[j])):
                        if res[-1] != (" "):
                            nrow[j] = nrow[j].replace('"', "")
                            res[-1] += nrow[j]
                        else:
                            nrow[j] = nrow[j].replace('"', "")
                            res.append(nrow[j])
                    elif "Label" in nrow[j]:
                        nrow[j] = nrow[j].replace('"', "")
                        ubi = (":").join(nrow[j].split(":")[1:])
                        res.append(ubi)
                    else:
                        nrow[j] = nrow[j].replace('"', "")
                        res.append(nrow[j])
                    j += 1
                res = unirCorchetes(res)
                if i == 0:
                    excel = pd.DataFrame(columns=res)
                else:
                    if not excel.empty and (0 in list(excel.iloc[-1])):
                        ultimo = list(excel.iloc[-1])
                        excel.iloc[-1] = ultimo[:ultimo.index(0)].append(res)
                    else:
                        nuevo_valor = {excel.columns[i] : res[i] if i < len(res) else 0 for i in range(len(excel.columns))}
                        excel = pd.concat([excel, pd.DataFrame([nuevo_valor])], ignore_index=True)

    excel.to_excel(nombreExcel + '.xlsx', index=False)

def main():
    entrada = str(input("Escribe el nombre o la ruta del csv importado"))
    entrada = str(encontrar_archivo(entrada + '.csv', "/."))
    salida = str(input("Escribe como se tiene que llamar el excel exportado"))
    crear_excel(entrada, salida)

class Ventana(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conversor de list a escel")
        self.setGeometry(200, 200, 500, 500)

        self.setStyleSheet("""
            QWidget {
                background-color: #f2f2f2;
            }
            QLineEdit {
                padding: 8px;
                border: 2px solid #ccc;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px;
                border: none;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QLabel {
                font-size: 16px;
            }
        """)

        # Widgets
        self.entrada = QLineEdit(self)
        self.entrada.setPlaceholderText("Escribe el nombre o la ruta del csv importado")

        self.salida = QLineEdit(self)
        self.salida.setPlaceholderText("Escribe como se tiene que llamar el excel exportado")

        self.boton = QPushButton("Convertir", self)
        self.etiqueta = QLabel("", self)

        self.exito = QLabel("Excel creado correctamente", self)
        self.exito.setStyleSheet("color : #28b463; background-color: #82e0aa;position: relative; text-align: center;top: 40px;left: 40px;")
        self.exito.setVisible(False)

        self.error = QLabel("No se ha encontrado el archivo nombrado, intentelo otra vez", self)
        self.error.setStyleSheet("color : #e74c3c; background-color: #f1948a ;position: relative; text-align: center;top: 40px;left: 40px;")
        self.error.setVisible(False)
        # Layout
        layout = QVBoxLayout()
        layout.addWidget(self.entrada)
        layout.addWidget(self.salida)
        layout.addWidget(self.boton)
        layout.addWidget(self.etiqueta)
        layout.addWidget(self.exito)
        layout.addWidget(self.error)
        self.setLayout(layout)

        # Conectar evento
        self.boton.clicked.connect(self.ventanillaExcel)

    def ventanillaExcel(self):
        entrada = str(encontrar_archivo(self.entrada.text() + '.csv', "/."))
        try:
            crear_excel(entrada, self.salida.text())
            self.error.setVisible(False)
            self.exito.setVisible(True)
        except FileNotFoundError:
            self.exito.setVisible(False)
            self.error.setVisible(True)

app = QApplication(sys.argv)
ventana = Ventana()
ventana.show()
sys.exit(app.exec_())