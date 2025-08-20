
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QLabel
from PyQt6.QtGui import QPalette, QColor, QFont
from PyQt6.QtCore import Qt
import backend

class DarkMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Informe de Producción - SEGUROS UNION")
        self.setGeometry(100, 100, 900, 600)
        self.setFont(QFont("Segoe UI", 11))
        self.initUI()
        self.applyDarkTheme()

    def initUI(self):
        central = QWidget()
        layout = QVBoxLayout()

        self.label = QLabel("Visualizador de Informe de Producción")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(self.label)

        self.btnCargar = QPushButton("Cargar Informe")
        self.btnCargar.clicked.connect(self.cargarDatos)
        layout.addWidget(self.btnCargar)

        self.tabla = QTableWidget()
        layout.addWidget(self.tabla)

        central.setLayout(layout)
        self.setCentralWidget(central)

    def applyDarkTheme(self):
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(30, 30, 30))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(220, 220, 220))
        palette.setColor(QPalette.ColorRole.Base, QColor(40, 40, 40))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(30, 30, 30))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(220, 220, 220))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(220, 220, 220))
        palette.setColor(QPalette.ColorRole.Text, QColor(220, 220, 220))
        palette.setColor(QPalette.ColorRole.Button, QColor(50, 50, 50))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(220, 220, 220))
        palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(45, 140, 240))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
        self.setPalette(palette)

    def cargarDatos(self):
        ruta, _ = QFileDialog.getOpenFileName(self, "Selecciona el archivo de informe", "", "Archivos XLSB (*.xlsb)")
        if ruta:
            df = backend.leer_excel(ruta)
            if df is not None:
                self.mostrarDatos(df)
            else:
                self.label.setText("Error al leer el archivo.")

    def mostrarDatos(self, df):
        self.tabla.clear()
        self.tabla.setRowCount(df.shape[0])
        self.tabla.setColumnCount(df.shape[1])
        self.tabla.setHorizontalHeaderLabels([str(col) for col in df.columns])
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
                self.tabla.setItem(i, j, item)
        self.label.setText(f"Mostrando {df.shape[0]} filas y {df.shape[1]} columnas")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DarkMainWindow()
    window.show()
    sys.exit(app.exec())
