
import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QLabel, QFrame, QComboBox, QDateEdit)
from PyQt6.QtCore import QDate
from PyQt6.QtGui import QPalette, QColor, QFont
from PyQt6.QtCore import Qt
import pandas as pd
from PyQt6.QtCharts import QChart, QChartView, QBarSeries, QBarSet, QBarCategoryAxis, QValueAxis, QPieSeries
import backend
import dashboard_utils

class DarkMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Informe de Producción - SEGUROS UNION")
        self.setGeometry(100, 100, 900, 600)
        self.setFont(QFont("Segoe UI", 11))
        # No crear widgets aquí, solo inicializar variables
        self.label = None
        self.btnCargar = None
        self.filtroRamo = None
        self.filtroAgente = None
        self.filtroCliente = None
        self.filtroFechaInicio = None
        self.filtroFechaFin = None
        self.btnExportar = None
        self.metric_frame = None
        self.metric_layout = None
        self.chart_view_bar = None
        self.chart_view_pie = None
        self.df = None
        self.df_filtrada = None
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

        # Filtros avanzados
        filtro_layout = QHBoxLayout()

        self.filtroRamo = QComboBox()
        self.filtroRamo.setPlaceholderText("Ramo (tipo de seguro)")
        self.filtroRamo.currentIndexChanged.connect(self.aplicarFiltros)
        filtro_layout.addWidget(self.filtroRamo)

        self.filtroAgente = QComboBox()
        self.filtroAgente.setPlaceholderText("Agente")
        self.filtroAgente.currentIndexChanged.connect(self.aplicarFiltros)
        filtro_layout.addWidget(self.filtroAgente)

        self.filtroCliente = QComboBox()
        self.filtroCliente.setPlaceholderText("Cliente")
        self.filtroCliente.currentIndexChanged.connect(self.aplicarFiltros)
        filtro_layout.addWidget(self.filtroCliente)

        self.filtroFechaInicio = QDateEdit()
        self.filtroFechaInicio.setCalendarPopup(True)
        self.filtroFechaInicio.setDisplayFormat("yyyy-MM-dd")
        self.filtroFechaInicio.setDate(QDate.currentDate().addMonths(-12))
        self.filtroFechaInicio.dateChanged.connect(self.aplicarFiltros)
        filtro_layout.addWidget(QLabel("Desde:"))
        filtro_layout.addWidget(self.filtroFechaInicio)

        self.filtroFechaFin = QDateEdit()
        self.filtroFechaFin.setCalendarPopup(True)
        self.filtroFechaFin.setDisplayFormat("yyyy-MM-dd")
        self.filtroFechaFin.setDate(QDate.currentDate())
        self.filtroFechaFin.dateChanged.connect(self.aplicarFiltros)
        filtro_layout.addWidget(QLabel("Hasta:"))
        filtro_layout.addWidget(self.filtroFechaFin)

        self.btnExportar = QPushButton("Exportar datos")
        self.btnExportar.clicked.connect(self.exportarDatos)
        filtro_layout.addWidget(self.btnExportar)

        main_layout.addLayout(filtro_layout)

        # Panel de métricas
        self.metric_frame = QFrame()
        self.metric_frame.setFrameShape(QFrame.Shape.StyledPanel)
        self.metric_layout = QHBoxLayout()
        self.metric_frame.setLayout(self.metric_layout)
        main_layout.addWidget(self.metric_frame)

        # Gráfica de barras
        self.chart_view_bar = QChartView()
        self.chart_view_bar.setMinimumHeight(250)
        main_layout.addWidget(self.chart_view_bar)

        # Gráfica de pastel
        self.chart_view_pie = QChartView()
        self.chart_view_pie.setMinimumHeight(250)
        main_layout.addWidget(self.chart_view_pie)

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
                self.df_filtrada = df
                # Actualizar filtros
                if 'TipoSeguro' in df.columns:
                    ramos = sorted(df['TipoSeguro'].dropna().unique())
                    self.filtroRamo.clear()
                    self.filtroRamo.addItem("Todos")
                    self.filtroRamo.addItems(ramos)
                if 'Agente' in df.columns:
                    agentes = sorted(df['Agente'].dropna().unique())
                    self.filtroAgente.clear()
                    self.filtroAgente.addItem("Todos")
                    self.filtroAgente.addItems(agentes)
                if 'Cliente' in df.columns:
                    clientes = sorted(df['Cliente'].dropna().unique())
                    self.filtroCliente.clear()
                    self.filtroCliente.addItem("Todos")
                    self.filtroCliente.addItems(clientes)
                self.aplicarFiltros()
            else:
                self.label.setText("Error al leer el archivo.")

    def mostrarDashboard(self):
        # Panel de métricas
        for i in reversed(range(self.metric_layout.count())):
            self.metric_layout.itemAt(i).widget().setParent(None)
        m = dashboard_utils.obtener_metricas(self.df_filtrada)
        def metric_label(text, value):
            lbl = QLabel(f"<b>{text}</b><br><span style='font-size:20px;'>{value}</span>")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet("background: #222; border-radius: 8px; padding: 12px; margin: 6px; color: #fff;")
            return lbl
        # Si no hay datos
        if self.df_filtrada is None or self.df_filtrada.empty:
            self.metric_layout.addWidget(metric_label("Sin datos", "-"))
            self.chart_view_bar.setChart(QChart())
            self.chart_view_pie.setChart(QChart())
            self.label.setText("No hay datos para mostrar.")
            return
        # Métricas principales
        if 'total_primas' in m:
            self.metric_layout.addWidget(metric_label("Total Primas", f"${m['total_primas']:,}"))
            # Promedio de primas
            promedio = self.df_filtrada['Prima'].mean() if 'Prima' in self.df_filtrada.columns else 0
            self.metric_layout.addWidget(metric_label("Promedio Prima", f"${promedio:,.2f}"))
        if 'total_polizas' in m:
            self.metric_layout.addWidget(metric_label("Total Pólizas", m['total_polizas']))
        if 'total_clientes' in m:
            self.metric_layout.addWidget(metric_label("Total Clientes", m['total_clientes']))
        if 'siniestros' in m:
            self.metric_layout.addWidget(metric_label("Siniestros", m['siniestros']))

        # Gráfica de barras: primas por mes
        self.chart_view_bar.setChart(QChart())
        chart_bar = self.chart_view_bar.chart()
        chart_bar.setTitle("Primas por Mes")
        chart_bar.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
        if 'primas_por_mes' in m and m['primas_por_mes']:
            meses = [str(k) for k in m['primas_por_mes'].keys()]
            primas = [v for v in m['primas_por_mes'].values()]
            barset = QBarSet("Primas")
            barset.append(primas)
            series = QBarSeries()
            series.append(barset)
            chart_bar.addSeries(series)
            axisX = QBarCategoryAxis()
            axisX.append(meses)
            chart_bar.setAxisX(axisX, series)
            axisY = QValueAxis()
            axisY.setLabelFormat("$%.0f")
            chart_bar.setAxisY(axisY, series)
        chart_bar.setTheme(QChart.ChartTheme.ChartThemeDark)

        # Gráfica de pastel: distribución por ramo
        self.chart_view_pie.setChart(QChart())
        chart_pie = self.chart_view_pie.chart()
        chart_pie.setTitle("Distribución por Ramo")
        chart_pie.setAnimationOptions(QChart.AnimationOption.SeriesAnimations)
        if 'tipos_seguros' in m and m['tipos_seguros']:
            pie_series = QPieSeries()
            for ramo, valor in m['tipos_seguros'].items():
                pie_series.append(str(ramo), valor)
            chart_pie.addSeries(pie_series)
        chart_pie.setTheme(QChart.ChartTheme.ChartThemeDark)

        self.label.setText("Dashboard visual: métricas y gráficos interactivos")

    def aplicarFiltros(self):
        if not hasattr(self, 'df') or self.df is None:
            self.df_filtrada = None
            self.mostrarDashboard()
            return
        df = self.df.copy()
        # Filtro por ramo
        if 'TipoSeguro' in df.columns:
            ramo = self.filtroRamo.currentText()
            if ramo and ramo != "Todos":
                df = df[df['TipoSeguro'].astype(str) == str(ramo)]
        # Filtro por agente
        if 'Agente' in df.columns:
            agente = self.filtroAgente.currentText()
            if agente and agente != "Todos":
                df = df[df['Agente'].astype(str) == str(agente)]
        # Filtro por cliente
        if 'Cliente' in df.columns:
            cliente = self.filtroCliente.currentText()
            if cliente and cliente != "Todos":
                df = df[df['Cliente'].astype(str) == str(cliente)]
        # Filtro por fechas
        if 'Fecha' in df.columns:
            fecha_inicio = self.filtroFechaInicio.date().toPyDate()
            fecha_fin = self.filtroFechaFin.date().toPyDate()
            fechas = pd.to_datetime(df['Fecha'], errors='coerce')
            df = df[(fechas >= pd.to_datetime(fecha_inicio)) & (fechas <= pd.to_datetime(fecha_fin))]
        self.df_filtrada = df
        self.mostrarDashboard()

    def exportarDatos(self):
        if not hasattr(self, 'df_filtrada') or self.df_filtrada.empty:
            return
        ruta, _ = QFileDialog.getSaveFileName(self, "Exportar datos", "datos_filtrados.csv", "Archivos CSV (*.csv);;Archivos Excel (*.xlsx)")
        if ruta:
            if ruta.endswith('.xlsx'):
                self.df_filtrada.to_excel(ruta, index=False)
            else:
                self.df_filtrada.to_csv(ruta, index=False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DarkMainWindow()
    window.show()
    sys.exit(app.exec())
