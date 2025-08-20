
import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QLabel, QFrame)
from PyQt6.QtGui import QPalette, QColor, QFont
from PyQt6.QtCore import Qt
from PyQt6.QtCharts import QChart, QChartView, QBarSeries, QBarSet, QBarCategoryAxis, QValueAxis
import backend
import dashboard_utils

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
        main_layout = QVBoxLayout()

        self.label = QLabel("Dashboard de Producción - SEGUROS UNION")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label.setStyleSheet("font-size: 22px; font-weight: bold;")
        main_layout.addWidget(self.label)

        self.btnCargar = QPushButton("Cargar Informe")
        self.btnCargar.clicked.connect(self.cargarDatos)
        main_layout.addWidget(self.btnCargar)

        # Panel de métricas
        self.metric_frame = QFrame()
        self.metric_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.metric_layout = QHBoxLayout()
        self.metric_frame.setLayout(self.metric_layout)
        main_layout.addWidget(self.metric_frame)

        # Gráfica
        self.chart_view = QChartView()
        self.chart_view.setMinimumHeight(250)
        main_layout.addWidget(self.chart_view)

        # Tabla de datos
        self.tabla = QTableWidget()
        main_layout.addWidget(self.tabla)

        central.setLayout(main_layout)
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
                self.df = df
                self.metricas = dashboard_utils.obtener_metricas(df)
                self.mostrarDashboard()
            else:
                self.label.setText("Error al leer el archivo.")

    def mostrarDashboard(self):
        # Panel de métricas
        for i in reversed(range(self.metric_layout.count())):
            self.metric_layout.itemAt(i).widget().setParent(None)
        m = self.metricas
        def metric_label(text, value):
            lbl = QLabel(f"<b>{text}</b><br><span style='font-size:20px;'>{value}</span>")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet("background: #222; border-radius: 8px; padding: 12px; margin: 6px; color: #fff;")
            return lbl
        if 'total_primas' in m:
            self.metric_layout.addWidget(metric_label("Total Primas", f"${m['total_primas']:,}"))
        if 'total_polizas' in m:
            self.metric_layout.addWidget(metric_label("Total Pólizas", m['total_polizas']))
        if 'total_clientes' in m:
            self.metric_layout.addWidget(metric_label("Total Clientes", m['total_clientes']))
        if 'siniestros' in m:
            self.metric_layout.addWidget(metric_label("Siniestros", m['siniestros']))

        # Gráfica animada: primas por mes
        self.chart_view.setChart(QChart())
        chart = self.chart_view.chart()
        chart.setTitle("Primas por Mes")
        chart.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
        if 'primas_por_mes' in m:
            meses = [str(k) for k in m['primas_por_mes'].keys()]
            primas = [v for v in m['primas_por_mes'].values()]
            barset = QBarSet("Primas")
            barset.append(primas)
            series = QBarSeries()
            series.append(barset)
            chart.addSeries(series)
            axisX = QBarCategoryAxis()
            axisX.append(meses)
            chart.setAxisX(axisX, series)
            axisY = QValueAxis()
            axisY.setLabelFormat("$%.0f")
            chart.setAxisY(axisY, series)
        chart.setTheme(QChart.ChartTheme.ChartThemeDark)

        # Tabla de datos
        df = self.df
        self.tabla.clear()
        self.tabla.setRowCount(df.shape[0])
        self.tabla.setColumnCount(df.shape[1])
        self.tabla.setHorizontalHeaderLabels([str(col) for col in df.columns])
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                item = QTableWidgetItem(str(df.iat[i, j]))
                self.tabla.setItem(i, j, item)
        self.label.setText(f"Dashboard: {df.shape[0]} filas, {df.shape[1]} columnas")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DarkMainWindow()
    window.show()
    sys.exit(app.exec())
