import sys
import os
import pandas as pd
import matplotlib.pyplot as plt
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QTableWidget, QTableWidgetItem, QTextEdit,
    QFileDialog, QMessageBox, QLineEdit
)
import qdarkstyle
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from fpdf import FPDF

# ----- FilterRow Widget for Multiple Filters -----
class FilterRow(QWidget):
    removed = QtCore.pyqtSignal(QWidget)  # Signal emitted when this row is to be removed

    def __init__(self, columns, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout()
        self.column_cb = QComboBox()
        self.column_cb.addItem("(None)")
        self.column_cb.addItems(columns)
        self.column_cb.setMinimumWidth(150)
        layout.addWidget(self.column_cb)

        self.operator_cb = QComboBox()
        self.operator_cb.addItems(["==", "!=", ">", "<", ">=", "<="])
        self.operator_cb.setMinimumWidth(80)
        layout.addWidget(self.operator_cb)

        self.value_cb = QComboBox()
        self.value_cb.setEditable(True)
        self.value_cb.setMinimumWidth(150)
        layout.addWidget(self.value_cb)

        self.remove_btn = QPushButton("Remove")
        self.remove_btn.setFixedWidth(80)
        self.remove_btn.clicked.connect(self.remove_self)
        layout.addWidget(self.remove_btn)

        self.setLayout(layout)

    def remove_self(self):
        self.removed.emit(self)

    def get_filter(self):
        """Return the filter condition (as a query string) from this row (or None if incomplete)."""
        col = self.column_cb.currentText()
        op = self.operator_cb.currentText()
        val = self.value_cb.currentText().strip()
        if col == "(None)" or val == "":
            return None
        try:
            # Attempt to interpret the value as numeric.
            float_val = float(val)
            return f"`{col}` {op} {float_val}"
        except ValueError:
            return f"`{col}` {op} '{val}'"

    def update_values(self, distinct_values):
        """Populate the value combo box with a sorted list of distinct values."""
        self.value_cb.clear()
        self.value_cb.addItems(sorted([str(x) for x in distinct_values]))

    def set_columns(self, columns):
        """Reset the column combo box with the given column list."""
        self.column_cb.clear()
        self.column_cb.addItem("(None)")
        self.column_cb.addItems(columns)

# ----- Data Tab -----
class DataTab(QWidget):
    data_changed = QtCore.pyqtSignal(pd.DataFrame)  # Signal emitted when data is loaded/filtered

    def __init__(self, parent=None):
        super().__init__(parent)
        self.original_df = None  # Holds the unfiltered data
        self.df = None         # Holds the current (possibly filtered) data
        self.filter_rows = []  # List of active FilterRow widgets
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Data source controls (Excel or CSV)
        source_layout = QHBoxLayout()
        self.source_cb = QComboBox()
        self.source_cb.addItems(["Excel", "CSV"])
        self.source_cb.currentIndexChanged.connect(self.update_source_ui)
        source_layout.addWidget(QLabel("Data Source:"))
        source_layout.addWidget(self.source_cb)

        self.load_btn = QPushButton("Load Data")
        self.load_btn.clicked.connect(self.load_data)
        source_layout.addWidget(self.load_btn)

        # Excel sheet selection (only visible when Excel is selected)
        self.sheet_label = QLabel("Sheet:")
        self.sheet_cb = QComboBox()
        self.sheet_cb.setMinimumWidth(150)
        self.sheet_cb.currentIndexChanged.connect(self.load_selected_sheet)
        source_layout.addWidget(self.sheet_label)
        source_layout.addWidget(self.sheet_cb)
        layout.addLayout(source_layout)

        # Data table to display loaded data
        self.table = QTableWidget()
        layout.addWidget(self.table)

        # Filtering controls
        filter_top_layout = QHBoxLayout()
        filter_top_layout.addWidget(QLabel("Filters:"))
        self.add_filter_btn = QPushButton("Add Filter")
        self.add_filter_btn.clicked.connect(self.add_filter_row)
        filter_top_layout.addWidget(self.add_filter_btn)
        self.apply_filter_btn = QPushButton("Apply Filters")
        self.apply_filter_btn.clicked.connect(self.apply_filters)
        filter_top_layout.addWidget(self.apply_filter_btn)
        self.reset_filter_btn = QPushButton("Reset Filters")
        self.reset_filter_btn.clicked.connect(self.reset_filters)
        filter_top_layout.addWidget(self.reset_filter_btn)
        layout.addLayout(filter_top_layout)

        # Container for dynamic filter rows
        self.filter_container = QVBoxLayout()
        layout.addLayout(self.filter_container)

        self.setLayout(layout)
        self.update_source_ui()

    def update_source_ui(self):
        """Show or hide Excel sheet controls based on the data source."""
        source = self.source_cb.currentText()
        if source == "Excel":
            self.sheet_label.show()
            self.sheet_cb.show()
        else:
            self.sheet_label.hide()
            self.sheet_cb.hide()

    def load_data(self):
        """Load data from an Excel or CSV file."""
        source = self.source_cb.currentText()
        if source == "Excel":
            file_path, _ = QFileDialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
            if file_path:
                try:
                    self.excel_file = pd.ExcelFile(file_path)
                    self.sheet_cb.clear()
                    self.sheet_cb.addItems(self.excel_file.sheet_names)
                    # Load the first sheet by default
                    self.df = self.excel_file.parse(self.excel_file.sheet_names[0])
                    self.original_df = self.df.copy()
                    self.update_table()
                    self.populate_filter_columns()
                    self.data_changed.emit(self.df)
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Failed to load Excel file:\n{e}")
        elif source == "CSV":
            file_path, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv)")
            if file_path:
                try:
                    self.df = pd.read_csv(file_path)
                    self.original_df = self.df.copy()
                    self.update_table()
                    self.populate_filter_columns()
                    self.data_changed.emit(self.df)
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Failed to load CSV file:\n{e}")

    def load_selected_sheet(self):
        """Reload data from the selected Excel sheet."""
        if self.source_cb.currentText() == "Excel" and hasattr(self, 'excel_file'):
            sheet_name = self.sheet_cb.currentText()
            try:
                self.df = self.excel_file.parse(sheet_name)
                self.original_df = self.df.copy()
                self.update_table()
                self.populate_filter_columns()
                self.data_changed.emit(self.df)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to load sheet:\n{e}")

    def update_table(self):
        """Display the current DataFrame in the table widget."""
        if self.df is not None:
            self.table.clear()
            self.table.setColumnCount(len(self.df.columns))
            self.table.setRowCount(len(self.df.index))
            self.table.setHorizontalHeaderLabels(self.df.columns)
            for i in range(len(self.df.index)):
                for j in range(len(self.df.columns)):
                    self.table.setItem(i, j, QTableWidgetItem(str(self.df.iat[i, j])))

    def populate_filter_columns(self):
        """Populate each filter row's column selector with the DataFrame columns."""
        if self.df is not None:
            columns = list(self.df.columns)
            for fr in self.filter_rows:
                fr.set_columns(columns)
                # If a filter row already has a selected column, update its values.
                if fr.column_cb.currentText() != "(None)":
                    distinct_vals = self.original_df[fr.column_cb.currentText()].dropna().unique()
                    fr.update_values(distinct_vals)

    def add_filter_row(self):
        """Add a new filter row to the UI."""
        if self.df is None:
            QMessageBox.information(self, "Info", "Load data first.")
            return
        columns = list(self.df.columns)
        fr = FilterRow(columns)
        fr.removed.connect(self.remove_filter_row)
        self.filter_rows.append(fr)
        self.filter_container.addWidget(fr)
        fr.column_cb.currentIndexChanged.connect(lambda: self.on_filter_column_changed(fr))

    def remove_filter_row(self, widget):
        """Remove a filter row from the UI."""
        self.filter_container.removeWidget(widget)
        widget.deleteLater()
        self.filter_rows.remove(widget)

    def on_filter_column_changed(self, fr):
        """Update the distinct values for a given filter row when its column changes."""
        col = fr.column_cb.currentText()
        if col != "(None)" and self.original_df is not None:
            distinct_vals = self.original_df[col].dropna().unique()
            fr.update_values(distinct_vals)

    def apply_filters(self):
        """Apply all active filters to the original data and update the table."""
        if self.original_df is None:
            return
        conditions = []
        for fr in self.filter_rows:
            filt = fr.get_filter()
            if filt:
                conditions.append(filt)
        if conditions:
            query_str = " and ".join(conditions)
            try:
                self.df = self.original_df.query(query_str)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to apply filters:\n{e}")
                return
        else:
            self.df = self.original_df.copy()
        self.update_table()
        self.data_changed.emit(self.df)

    def reset_filters(self):
        """Reset filters and show the original data."""
        self.df = self.original_df.copy()
        self.update_table()
        for fr in self.filter_rows:
            self.filter_container.removeWidget(fr)
            fr.deleteLater()
        self.filter_rows = []
        self.data_changed.emit(self.df)

# ----- Chart Tab (with improved chart mode selection) -----
class ChartTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.df = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        # Top controls: Chart Mode, Column(s), Chart Type, and axis labels
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("Chart Mode:"))
        self.chart_mode_cb = QComboBox()
        self.chart_mode_cb.addItems(["Single Numeric Column", "XY Chart"])
        self.chart_mode_cb.currentIndexChanged.connect(self.update_chart_mode_ui)
        mode_layout.addWidget(self.chart_mode_cb)

        # For Single Numeric Column mode
        self.single_column_cb = QComboBox()
        self.single_column_cb.setMinimumWidth(150)
        mode_layout.addWidget(QLabel("Column:"))
        mode_layout.addWidget(self.single_column_cb)

        # For XY Chart mode (initially hidden)
        self.x_column_cb = QComboBox()
        self.x_column_cb.setMinimumWidth(150)
        self.y_column_cb = QComboBox()
        self.y_column_cb.setMinimumWidth(150)
        self.x_column_cb.hide()
        self.y_column_cb.hide()
        self.xy_label = QLabel("X, Y:")
        self.xy_label.hide()
        mode_layout.addWidget(self.xy_label)
        mode_layout.addWidget(self.x_column_cb)
        mode_layout.addWidget(self.y_column_cb)

        # Chart type
        mode_layout.addWidget(QLabel("Chart Type:"))
        self.chart_type_cb = QComboBox()
        mode_layout.addWidget(self.chart_type_cb)

        # Axis label inputs
        self.xlabel_le = QLineEdit()
        self.xlabel_le.setPlaceholderText("X-axis Label")
        self.ylabel_le = QLineEdit()
        self.ylabel_le.setPlaceholderText("Y-axis Label")
        mode_layout.addWidget(QLabel("X Label:"))
        mode_layout.addWidget(self.xlabel_le)
        mode_layout.addWidget(QLabel("Y Label:"))
        mode_layout.addWidget(self.ylabel_le)

        # Generate and Export buttons
        self.generate_chart_btn = QPushButton("Generate Chart")
        self.generate_chart_btn.clicked.connect(self.generate_chart)
        mode_layout.addWidget(self.generate_chart_btn)
        self.export_chart_btn = QPushButton("Export Chart")
        self.export_chart_btn.clicked.connect(self.export_chart)
        mode_layout.addWidget(self.export_chart_btn)

        layout.addLayout(mode_layout)

        # Matplotlib canvas for interactive charts
        self.figure = Figure(figsize=(5, 4))
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)
        self.setLayout(layout)
        self.update_chart_mode_ui()

    def update_data(self, df):
        self.df = df
        self.update_numeric_columns()

    def update_numeric_columns(self):
        """Update the column selection controls with numeric columns only."""
        if self.df is not None:
            num_cols = [col for col in self.df.columns if pd.api.types.is_numeric_dtype(self.df[col])]
            # For single mode, update self.single_column_cb
            self.single_column_cb.clear()
            self.single_column_cb.addItem("(None)")
            self.single_column_cb.addItems(num_cols)
            # For XY mode, update both x and y selectors
            self.x_column_cb.clear()
            self.x_column_cb.addItem("(None)")
            self.x_column_cb.addItems(num_cols)
            self.y_column_cb.clear()
            self.y_column_cb.addItem("(None)")
            self.y_column_cb.addItems(num_cols)

    def update_chart_mode_ui(self):
        """Show/hide controls based on chart mode and update chart type options."""
        mode = self.chart_mode_cb.currentText()
        if mode == "Single Numeric Column":
            self.single_column_cb.show()
            self.xy_label.hide()
            self.x_column_cb.hide()
            self.y_column_cb.hide()
            # Available chart types for single column mode
            self.chart_type_cb.clear()
            self.chart_type_cb.addItems(["Histogram", "Bar Chart", "Line Chart", "Pie Chart"])
        else:  # XY Chart mode
            self.single_column_cb.hide()
            self.xy_label.show()
            self.x_column_cb.show()
            self.y_column_cb.show()
            # Available chart types for XY charting
            self.chart_type_cb.clear()
            self.chart_type_cb.addItems(["Scatter Plot", "Line Chart"])
        # Update numeric columns from the current data
        self.update_numeric_columns()

    def generate_chart(self):
        if self.df is None:
            QMessageBox.information(self, "Info", "Load data first.")
            return
        mode = self.chart_mode_cb.currentText()
        chart_type = self.chart_type_cb.currentText()
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        try:
            if mode == "Single Numeric Column":
                col = self.single_column_cb.currentText()
                if col == "(None)":
                    QMessageBox.information(self, "Info", "Select a numeric column for charting.")
                    return
                data = self.df[col].dropna()
                if data.empty:
                    QMessageBox.information(self, "Info", "No data in selected column.")
                    return
                if chart_type == "Histogram":
                    data_numeric = data.astype(float)
                    ax.hist(data_numeric, bins=20, color="skyblue", edgecolor="black")
                    ax.set_title(f"Histogram of {col}")
                    ax.set_xlabel(self.xlabel_le.text() if self.xlabel_le.text() else col)
                    ax.set_ylabel(self.ylabel_le.text() if self.ylabel_le.text() else "Frequency")
                elif chart_type == "Bar Chart":
                    counts = data.value_counts()
                    counts.plot(kind="bar", ax=ax, color="lightgreen", edgecolor="black")
                    ax.set_title(f"Bar Chart of {col}")
                    ax.set_xlabel(self.xlabel_le.text() if self.xlabel_le.text() else col)
                    ax.set_ylabel(self.ylabel_le.text() if self.ylabel_le.text() else "Count")
                elif chart_type == "Line Chart":
                    data_numeric = data.astype(float)
                    ax.plot(data_numeric.index, data_numeric, marker="o")
                    ax.set_title(f"Line Chart of {col}")
                    ax.set_xlabel(self.xlabel_le.text() if self.xlabel_le.text() else "Index")
                    ax.set_ylabel(self.ylabel_le.text() if self.ylabel_le.text() else col)
                elif chart_type == "Pie Chart":
                    counts = data.value_counts()
                    ax.pie(counts, labels=counts.index, autopct="%1.1f%%")
                    ax.set_title(f"Pie Chart of {col}")
            else:  # XY Chart mode
                x_col = self.x_column_cb.currentText()
                y_col = self.y_column_cb.currentText()
                if x_col == "(None)" or y_col == "(None)":
                    QMessageBox.information(self, "Info", "Select numeric columns for both X and Y.")
                    return
                x_data = self.df[x_col].dropna()
                y_data = self.df[y_col].dropna()
                if x_data.empty or y_data.empty:
                    QMessageBox.information(self, "Info", "No data in selected columns.")
                    return
                # To ensure equal length, take the intersection
                min_len = min(len(x_data), len(y_data))
                x_data = x_data.iloc[:min_len]
                y_data = y_data.iloc[:min_len]
                if chart_type == "Scatter Plot":
                    ax.scatter(x_data.astype(float), y_data.astype(float), color="purple")
                    ax.set_title(f"Scatter Plot of {x_col} vs {y_col}")
                elif chart_type == "Line Chart":
                    ax.plot(x_data.astype(float), y_data.astype(float), marker="o", linestyle="-")
                    ax.set_title(f"Line Chart of {x_col} vs {y_col}")
                ax.set_xlabel(self.xlabel_le.text() if self.xlabel_le.text() else x_col)
                ax.set_ylabel(self.ylabel_le.text() if self.ylabel_le.text() else y_col)
            self.canvas.draw()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to generate chart:\n{e}")

    def export_chart(self):
        if self.figure:
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Chart", "", "PNG Files (*.png)")
            if file_path:
                try:
                    self.figure.savefig(file_path)
                    QMessageBox.information(self, "Success", "Chart exported successfully.")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"Failed to export chart:\n{e}")

# ----- Report Tab (with improved export options) -----
class ReportTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.df = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.report_text = QTextEdit()
        self.report_text.setReadOnly(True)
        layout.addWidget(self.report_text)

        btn_layout = QHBoxLayout()
        self.generate_report_btn = QPushButton("Generate Report")
        self.generate_report_btn.clicked.connect(self.generate_report)
        btn_layout.addWidget(self.generate_report_btn)
        self.export_pdf_btn = QPushButton("Export Report as PDF")
        self.export_pdf_btn.clicked.connect(self.export_report_pdf)
        btn_layout.addWidget(self.export_pdf_btn)
        self.export_excel_btn = QPushButton("Export Report as Excel")
        self.export_excel_btn.clicked.connect(self.export_report_excel)
        btn_layout.addWidget(self.export_excel_btn)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def update_data(self, df):
        self.df = df

    def generate_report(self):
        if self.df is None:
            self.report_text.setText("No data loaded.")
            return
        report_lines = []
        total_records = len(self.df)
        columns = list(self.df.columns)
        report_lines.append("Report Summary:")
        report_lines.append(f"Total Records: {total_records}")
        report_lines.append(f"Columns: {', '.join(columns)}")
        report_lines.append("\nNumeric Column Statistics:")
        numeric_cols = self.df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            stats = self.df[numeric_cols].describe().to_string()
            report_lines.append(stats)
        else:
            report_lines.append("No numeric columns found.")
        report_text = "\n".join(report_lines)
        self.report_text.setText(report_text)

    def export_report_pdf(self):
        if self.df is None:
            QMessageBox.information(self, "Info", "No data to export.")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Report as PDF", "", "PDF Files (*.pdf)")
        if not file_path:
            return
        try:
            # Force landscape orientation ("L") on Letter-sized paper.
            pdf = FPDF(orientation="L", format="Letter")
            pdf.add_page()
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Report Summary", ln=True, align="C")
            pdf.ln(5)
            pdf.set_font("Arial", "", 12)
            # Create summary table
            summary_data = [
                ("Total Records", str(len(self.df))),
                ("Columns", ", ".join(list(self.df.columns)))
            ]
            for metric, value in summary_data:
                pdf.cell(50, 10, metric, border=1)
                pdf.cell(0, 10, value, border=1, ln=True)
            pdf.ln(5)
            # Numeric statistics table
            numeric_cols = self.df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                pdf.set_font("Arial", "B", 14)
                pdf.cell(0, 10, "Numeric Column Statistics", ln=True, align="C")
                pdf.ln(5)
                stats_df = self.df[numeric_cols].describe().reset_index()
                pdf.set_font("Arial", "B", 10)
                # Table header
                for col in stats_df.columns:
                    pdf.cell(30, 8, str(col), border=1)
                pdf.ln()
                pdf.set_font("Arial", "", 10)
                # Table rows
                for idx, row in stats_df.iterrows():
                    for item in row:
                        text = f"{item:.2f}" if isinstance(item, (float, int)) else str(item)
                        pdf.cell(30, 8, text, border=1)
                    pdf.ln()
            pdf.output(file_path)
            QMessageBox.information(self, "Success", "Report exported as PDF successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to export report as PDF:\n{e}")

    def export_report_excel(self):
        if self.df is None:
            QMessageBox.information(self, "Info", "No data to export.")
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Report as Excel", "", "Excel Files (*.xlsx)")
        if not file_path:
            return
        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Write the filtered data in a sheet called "Data"
                self.df.to_excel(writer, sheet_name="Data", index=False)
                # Write summary info in a sheet called "Summary"
                summary_df = pd.DataFrame({
                    "Metric": ["Total Records", "Columns"],
                    "Value": [len(self.df), ", ".join(list(self.df.columns))]
                })
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                # Write numeric statistics in a sheet called "Numeric Statistics" (if available)
                numeric_cols = self.df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    stats_df = self.df[numeric_cols].describe()
                    stats_df.to_excel(writer, sheet_name="Numeric Statistics")
            QMessageBox.information(self, "Success", "Report exported as Excel successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to export report as Excel:\n{e}")

# ----- Main Window -----
class DataInsightStudio(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Insight Studio")
        self.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
        self.resize(1000, 700)
        self.init_ui()

    def init_ui(self):
        self.tabs = QTabWidget()
        self.data_tab = DataTab(self)
        self.chart_tab = ChartTab(self)
        self.report_tab = ReportTab(self)

        self.tabs.addTab(self.data_tab, "Data")
        self.tabs.addTab(self.chart_tab, "Chart")
        self.tabs.addTab(self.report_tab, "Report")

        self.setCentralWidget(self.tabs)

        # When data is loaded or filtered, update the Chart and Report tabs.
        self.data_tab.data_changed.connect(self.on_data_changed)

    def on_data_changed(self, df):
        self.chart_tab.update_data(df)
        self.report_tab.update_data(df)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataInsightStudio()
    window.show()
    sys.exit(app.exec_())
