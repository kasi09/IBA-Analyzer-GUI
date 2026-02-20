"""
IBA-Analyzer-GUI
================
PySide6 GUI for reading iba PDA .dat files, replicating the ibaAnalyzer look and feel.
Uses iba_reader.py as backend for COM-based signal reading.
"""

import sys
import os

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QDockWidget, QTreeWidget, QTreeWidgetItem, QTableWidget,
    QTableWidgetItem, QTabWidget, QToolBar, QStatusBar, QMenuBar,
    QFileDialog, QMessageBox, QHeaderView, QLineEdit, QPushButton,
    QLabel, QSplitter, QProgressBar, QCheckBox, QColorDialog,
    QStyle, QAbstractItemView,
)
from PySide6.QtCore import (
    Qt, QThread, Signal, QSize, QSettings, QTimer,
)
from PySide6.QtGui import (
    QAction, QIcon, QColor, QPixmap, QPainter, QFont, QPalette,
    QKeySequence,
)

# Add parent directory so iba_reader can be imported
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from iba_reader import IbaReader, FILTER_ANALOG, FILTER_DIGITAL, FILTER_TEXT


# ---------------------------------------------------------------------------
# Color helpers
# ---------------------------------------------------------------------------
SIGNAL_COLORS = [
    QColor(0, 0, 255),       # Blue
    QColor(255, 0, 0),       # Red
    QColor(0, 128, 0),       # Green
    QColor(255, 165, 0),     # Orange
    QColor(128, 0, 128),     # Purple
    QColor(0, 128, 128),     # Teal
    QColor(128, 128, 0),     # Olive
    QColor(255, 0, 255),     # Magenta
    QColor(0, 0, 128),       # Navy
    QColor(128, 0, 0),       # Maroon
]


def color_icon(color, size=16):
    """Create a small square icon filled with the given color."""
    pixmap = QPixmap(size, size)
    pixmap.fill(color)
    return QIcon(pixmap)


# ---------------------------------------------------------------------------
# Worker thread for loading signals
# ---------------------------------------------------------------------------
class SignalLoaderThread(QThread):
    """Loads signals from a .dat file in a background thread."""
    finished = Signal(dict)
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, filepath):
        super().__init__()
        self.filepath = filepath

    def run(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                reader = IbaReader(self.filepath)
                reader.open()

                self.progress.emit("Lese analoge Signale...")
                analog = reader.get_signal_list(FILTER_ANALOG)

                self.progress.emit("Lese digitale Signale...")
                digital = reader.get_signal_list(FILTER_DIGITAL)

                self.progress.emit("Lese Text-Signale...")
                text = reader.get_signal_list(FILTER_TEXT)

                self.progress.emit("Lese Datei-Informationen...")
                version = reader.version

                reader.close()

                self.finished.emit({
                    'analog': analog,
                    'digital': digital,
                    'text': text,
                    'version': version,
                    'filepath': self.filepath,
                })
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            self.error.emit(str(e))


# ---------------------------------------------------------------------------
# Signal Tree Panel
# ---------------------------------------------------------------------------
class SignalTreePanel(QWidget):
    """Left panel: Signal tree with search tab."""
    signal_double_clicked = Signal(dict)  # emits signal info dict

    def __init__(self, parent=None):
        super().__init__(parent)
        self._signals_data = {'analog': [], 'digital': [], 'text': []}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.South)

        # --- Signale tab ---
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setAnimated(True)
        self.tree.setIndentation(16)
        self.tree.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.tabs.addTab(self.tree, "Signale")

        # --- Suchen tab ---
        search_widget = QWidget()
        search_layout = QVBoxLayout(search_widget)
        search_layout.setContentsMargins(4, 4, 4, 4)

        search_bar = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Suchmuster (z.B. *Speed*)")
        self.search_input.returnPressed.connect(self._do_search)
        search_bar.addWidget(self.search_input)

        self.search_btn = QPushButton("Suchen")
        self.search_btn.clicked.connect(self._do_search)
        search_bar.addWidget(self.search_btn)
        search_layout.addLayout(search_bar)

        self.search_results = QTreeWidget()
        self.search_results.setHeaderLabels(["ID", "Name", "Gruppe"])
        self.search_results.setRootIsDecorated(False)
        self.search_results.itemDoubleClicked.connect(self._on_search_result_double_clicked)
        header = self.search_results.header()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        search_layout.addWidget(self.search_results)

        self.tabs.addTab(search_widget, "Suchen")

        layout.addWidget(self.tabs)

    def populate(self, signals_data, filepath):
        """Populate the tree with signal data."""
        self._signals_data = signals_data
        self.tree.clear()

        filename = os.path.basename(filepath)
        root = QTreeWidgetItem(self.tree, [filename])
        root.setIcon(0, self.style().standardIcon(QStyle.SP_FileIcon))
        font = root.font(0)
        font.setBold(True)
        root.setFont(0, font)

        for label, key, icon_type in [
            ("Analog", 'analog', QStyle.SP_MediaVolume),
            ("Digital", 'digital', QStyle.SP_DialogApplyButton),
            ("Text", 'text', QStyle.SP_FileDialogDetailedView),
        ]:
            sigs = signals_data.get(key, [])
            if not sigs:
                continue

            type_node = QTreeWidgetItem(root, [f"{label} ({len(sigs)} Signale)"])
            type_node.setIcon(0, self.style().standardIcon(icon_type))
            font = type_node.font(0)
            font.setBold(True)
            type_node.setFont(0, font)

            # Group signals by group name
            groups = {}
            for s in sigs:
                g = s.get('group', '')
                if g not in groups:
                    groups[g] = []
                groups[g].append(s)

            for group_name in sorted(groups.keys()):
                group_signals = groups[group_name]
                if group_name:
                    group_node = QTreeWidgetItem(type_node, [group_name])
                    group_node.setIcon(0, self.style().standardIcon(QStyle.SP_DirIcon))
                    parent = group_node
                else:
                    parent = type_node

                for s in group_signals:
                    item = QTreeWidgetItem(parent, [f"{s['id']}  {s['name']}"])
                    item.setData(0, Qt.UserRole, s)
                    item.setToolTip(0, f"ID: {s['id']}\nName: {s['name']}\nGruppe: {s['group']}")

        root.setExpanded(True)

    def _on_item_double_clicked(self, item, column):
        sig_data = item.data(0, Qt.UserRole)
        if sig_data:
            self.signal_double_clicked.emit(sig_data)

    def _on_search_result_double_clicked(self, item, column):
        sig_data = item.data(0, Qt.UserRole)
        if sig_data:
            self.signal_double_clicked.emit(sig_data)

    def _do_search(self):
        pattern = self.search_input.text().strip()
        if not pattern:
            return

        self.search_results.clear()
        import fnmatch
        import re

        all_signals = (
            self._signals_data.get('analog', []) +
            self._signals_data.get('digital', []) +
            self._signals_data.get('text', [])
        )

        if any(c in pattern for c in ['*', '?']):
            matches = [s for s in all_signals if fnmatch.fnmatch(s['name'], pattern)]
        else:
            try:
                regex = re.compile(pattern, re.IGNORECASE)
                matches = [s for s in all_signals if regex.search(s['name'])]
            except re.error:
                matches = [s for s in all_signals if pattern.lower() in s['name'].lower()]

        for s in matches:
            item = QTreeWidgetItem(self.search_results, [s['id'], s['name'], s['group']])
            item.setData(0, Qt.UserRole, s)

        self.statusTip = f"{len(matches)} Ergebnis(se)"


# ---------------------------------------------------------------------------
# Signal Definitions Panel (Bottom Dock)
# ---------------------------------------------------------------------------
class SignalDefinitionsPanel(QWidget):
    """Bottom panel: Signal definitions table, statistics, overview."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._color_index = 0

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()
        self.tabs.setTabPosition(QTabWidget.South)

        # --- Signal Definitionen tab ---
        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels([
            "Anzeige", "Signalname", "Ausdruck", "Einheit", "Farbe"
        ])
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(22)
        self.tabs.addTab(self.table, "Signal Definitionen")

        # --- Statistik tab ---
        self.stats_table = QTableWidget(0, 6)
        self.stats_table.setHorizontalHeaderLabels([
            "Signal", "Min", "Max", "Mittelwert", "Datenpunkte", "Abtastrate"
        ])
        self.stats_table.horizontalHeader().setStretchLastSection(True)
        self.stats_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for i in range(1, 6):
            self.stats_table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
        self.stats_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.stats_table.setAlternatingRowColors(True)
        self.stats_table.verticalHeader().setDefaultSectionSize(22)
        self.tabs.addTab(self.stats_table, "Statistik")

        # --- Übersicht tab ---
        self.overview_widget = QWidget()
        overview_layout = QVBoxLayout(self.overview_widget)
        overview_layout.setContentsMargins(8, 8, 8, 8)
        self.overview_label = QLabel("Keine Datei geladen.")
        self.overview_label.setWordWrap(True)
        self.overview_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        overview_layout.addWidget(self.overview_label)
        self.tabs.addTab(self.overview_widget, "Übersicht")

        layout.addWidget(self.tabs)

    def add_signal(self, signal_info):
        """Add a signal to the definitions table."""
        # Check if already in table
        for row in range(self.table.rowCount()):
            expr_item = self.table.item(row, 2)
            if expr_item and expr_item.text() == signal_info['id']:
                return  # Already added

        row = self.table.rowCount()
        self.table.insertRow(row)

        # Anzeige (checkbox)
        checkbox = QCheckBox()
        checkbox.setChecked(True)
        checkbox_widget = QWidget()
        cb_layout = QHBoxLayout(checkbox_widget)
        cb_layout.addWidget(checkbox)
        cb_layout.setAlignment(Qt.AlignCenter)
        cb_layout.setContentsMargins(0, 0, 0, 0)
        self.table.setCellWidget(row, 0, checkbox_widget)

        # Signalname
        self.table.setItem(row, 1, QTableWidgetItem(signal_info['name']))

        # Ausdruck (expression / channel ID)
        self.table.setItem(row, 2, QTableWidgetItem(signal_info['id']))

        # Einheit
        self.table.setItem(row, 3, QTableWidgetItem(""))

        # Farbe
        color = SIGNAL_COLORS[self._color_index % len(SIGNAL_COLORS)]
        self._color_index += 1
        color_item = QTableWidgetItem()
        color_item.setBackground(color)
        color_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        self.table.setItem(row, 4, color_item)

    def clear_signals(self):
        """Remove all signals from the table."""
        self.table.setRowCount(0)
        self.stats_table.setRowCount(0)
        self._color_index = 0

    def get_selected_expressions(self):
        """Return list of expressions from selected/checked rows."""
        expressions = []
        for row in range(self.table.rowCount()):
            widget = self.table.cellWidget(row, 0)
            if widget:
                checkbox = widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    expr_item = self.table.item(row, 2)
                    if expr_item:
                        expressions.append(expr_item.text())
        return expressions

    def get_all_expressions(self):
        """Return all expressions in the table."""
        expressions = []
        for row in range(self.table.rowCount()):
            expr_item = self.table.item(row, 2)
            if expr_item:
                expressions.append(expr_item.text())
        return expressions

    def update_overview(self, signals_data, filepath, version):
        """Update the overview tab with file information."""
        analog_count = len(signals_data.get('analog', []))
        digital_count = len(signals_data.get('digital', []))
        text_count = len(signals_data.get('text', []))
        total = analog_count + digital_count + text_count

        self.overview_label.setText(
            f"<b>Datei:</b> {os.path.basename(filepath)}<br>"
            f"<b>Pfad:</b> {filepath}<br>"
            f"<b>ibaAnalyzer Version:</b> {version}<br>"
            f"<br>"
            f"<b>Signale gesamt:</b> {total}<br>"
            f"  Analog: {analog_count}<br>"
            f"  Digital: {digital_count}<br>"
            f"  Text: {text_count}"
        )


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    """Main application window replicating the ibaAnalyzer layout."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("IBA-Analyzer-GUI")
        self.resize(1200, 800)

        self._reader = None
        self._filepath = None
        self._signals_data = {}
        self._loader_thread = None

        self._settings = QSettings("kasi09", "IBA-Analyzer-GUI")

        self._setup_ui()
        self._create_menus()
        self._create_toolbar()
        self._restore_state()

    def _setup_ui(self):
        """Set up the main UI layout."""
        # Central widget (placeholder for future graph views)
        central = QWidget()
        central.setStyleSheet("background-color: #e0e0e0;")
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(0, 0, 0, 0)
        self.placeholder_label = QLabel("Signalansicht")
        self.placeholder_label.setAlignment(Qt.AlignCenter)
        self.placeholder_label.setStyleSheet(
            "color: #888; font-size: 18px; font-style: italic; background: transparent;"
        )
        central_layout.addWidget(self.placeholder_label)
        self.setCentralWidget(central)

        # Left dock: Signal tree
        self.signal_tree_dock = QDockWidget("Signale", self)
        self.signal_tree_dock.setObjectName("SignalTreeDock")
        self.signal_tree_dock.setMinimumWidth(250)
        self.signal_tree_panel = SignalTreePanel()
        self.signal_tree_panel.signal_double_clicked.connect(self._on_signal_selected)
        self.signal_tree_dock.setWidget(self.signal_tree_panel)
        self.addDockWidget(Qt.LeftDockWidgetArea, self.signal_tree_dock)

        # Bottom dock: Signal definitions
        self.signal_defs_dock = QDockWidget("Signal Definitionen", self)
        self.signal_defs_dock.setObjectName("SignalDefsDock")
        self.signal_defs_dock.setMinimumHeight(150)
        self.signal_defs_panel = SignalDefinitionsPanel()
        self.signal_defs_dock.setWidget(self.signal_defs_panel)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.signal_defs_dock)

        # Status bar
        self.statusBar().showMessage("Bereit")

        # Progress bar in status bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.setMaximumHeight(16)
        self.progress_bar.setVisible(False)
        self.statusBar().addPermanentWidget(self.progress_bar)

    def _create_menus(self):
        """Create the menu bar."""
        menubar = self.menuBar()

        # --- Datei ---
        file_menu = menubar.addMenu("&Datei")

        self.action_open = QAction("&Öffnen...", self)
        self.action_open.setShortcut(QKeySequence.Open)
        self.action_open.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        self.action_open.triggered.connect(self._open_file)
        file_menu.addAction(self.action_open)

        file_menu.addSeparator()

        self.action_export_csv = QAction("Export &CSV...", self)
        self.action_export_csv.setEnabled(False)
        self.action_export_csv.triggered.connect(self._export_csv)
        file_menu.addAction(self.action_export_csv)

        self.action_export_parquet = QAction("Export &Parquet...", self)
        self.action_export_parquet.setEnabled(False)
        self.action_export_parquet.triggered.connect(self._export_parquet)
        file_menu.addAction(self.action_export_parquet)

        self.action_export_video = QAction("Export &Video...", self)
        self.action_export_video.setEnabled(False)
        self.action_export_video.triggered.connect(self._export_video)
        file_menu.addAction(self.action_export_video)

        file_menu.addSeparator()

        self.action_close = QAction("Datei s&chließen", self)
        self.action_close.setEnabled(False)
        self.action_close.triggered.connect(self._close_file)
        file_menu.addAction(self.action_close)

        file_menu.addSeparator()

        self.action_exit = QAction("&Beenden", self)
        self.action_exit.setShortcut(QKeySequence.Quit)
        self.action_exit.triggered.connect(self.close)
        file_menu.addAction(self.action_exit)

        # --- Bearbeiten ---
        edit_menu = menubar.addMenu("&Bearbeiten")

        self.action_clear_signals = QAction("Signal-Liste &leeren", self)
        self.action_clear_signals.triggered.connect(self._clear_signal_table)
        edit_menu.addAction(self.action_clear_signals)

        self.action_remove_signal = QAction("Signal &entfernen", self)
        self.action_remove_signal.setShortcut(QKeySequence.Delete)
        self.action_remove_signal.triggered.connect(self._remove_selected_signal)
        edit_menu.addAction(self.action_remove_signal)

        # --- Ansicht ---
        view_menu = menubar.addMenu("&Ansicht")

        self.action_toggle_tree = self.signal_tree_dock.toggleViewAction()
        self.action_toggle_tree.setText("&Signal Panel")
        view_menu.addAction(self.action_toggle_tree)

        self.action_toggle_defs = self.signal_defs_dock.toggleViewAction()
        self.action_toggle_defs.setText("Signal &Definitionen")
        view_menu.addAction(self.action_toggle_defs)

        # --- Hilfe ---
        help_menu = menubar.addMenu("&Hilfe")

        self.action_about = QAction("Ü&ber...", self)
        self.action_about.triggered.connect(self._show_about)
        help_menu.addAction(self.action_about)

    def _create_toolbar(self):
        """Create the main toolbar."""
        toolbar = QToolBar("Hauptwerkzeugleiste")
        toolbar.setObjectName("MainToolbar")
        toolbar.setIconSize(QSize(20, 20))
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        toolbar.addAction(self.action_open)
        toolbar.addSeparator()
        toolbar.addAction(self.action_export_csv)
        toolbar.addAction(self.action_export_parquet)
        toolbar.addAction(self.action_export_video)

    # -----------------------------------------------------------------------
    # File operations
    # -----------------------------------------------------------------------
    def _open_file(self):
        """Open a .dat file."""
        last_dir = self._settings.value("last_directory", "")
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            "iba .dat Datei öffnen",
            last_dir,
            "iba Dateien (*.dat);;Alle Dateien (*.*)",
        )
        if not filepath:
            return

        self._settings.setValue("last_directory", os.path.dirname(filepath))
        self._load_file(filepath)

    def _load_file(self, filepath):
        """Load a .dat file using a background thread."""
        if self._loader_thread and self._loader_thread.isRunning():
            return

        self._filepath = filepath
        self.setWindowTitle(f"{os.path.basename(filepath)} - IBA-Analyzer-GUI")
        self.statusBar().showMessage(f"Lade {os.path.basename(filepath)}...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Indeterminate

        self.signal_defs_panel.clear_signals()

        self._loader_thread = SignalLoaderThread(filepath)
        self._loader_thread.finished.connect(self._on_signals_loaded)
        self._loader_thread.error.connect(self._on_load_error)
        self._loader_thread.progress.connect(lambda msg: self.statusBar().showMessage(msg))
        self._loader_thread.start()

    def _on_signals_loaded(self, result):
        """Handle successful signal loading."""
        self.progress_bar.setVisible(False)

        self._signals_data = {
            'analog': result['analog'],
            'digital': result['digital'],
            'text': result['text'],
        }

        self.signal_tree_panel.populate(self._signals_data, result['filepath'])
        self.signal_defs_panel.update_overview(
            self._signals_data, result['filepath'], result['version']
        )

        total = len(result['analog']) + len(result['digital']) + len(result['text'])
        self.statusBar().showMessage(
            f"Fertig — {total} Signale geladen "
            f"(Analog: {len(result['analog'])}, "
            f"Digital: {len(result['digital'])}, "
            f"Text: {len(result['text'])})"
        )

        self.action_export_csv.setEnabled(True)
        self.action_export_parquet.setEnabled(True)
        self.action_export_video.setEnabled(True)
        self.action_close.setEnabled(True)

    def _on_load_error(self, error_msg):
        """Handle signal loading error."""
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage("Fehler beim Laden")
        QMessageBox.critical(
            self, "Fehler",
            f"Fehler beim Laden der Datei:\n\n{error_msg}"
        )

    def _close_file(self):
        """Close the current file."""
        self._filepath = None
        self._signals_data = {}
        self.signal_tree_panel.tree.clear()
        self.signal_defs_panel.clear_signals()
        self.signal_defs_panel.overview_label.setText("Keine Datei geladen.")
        self.setWindowTitle("IBA-Analyzer-GUI")
        self.statusBar().showMessage("Bereit")
        self.action_export_csv.setEnabled(False)
        self.action_export_parquet.setEnabled(False)
        self.action_export_video.setEnabled(False)
        self.action_close.setEnabled(False)

    # -----------------------------------------------------------------------
    # Signal operations
    # -----------------------------------------------------------------------
    def _on_signal_selected(self, signal_info):
        """Handle double-click on a signal in the tree."""
        self.signal_defs_panel.add_signal(signal_info)
        self.signal_defs_panel.tabs.setCurrentIndex(0)

    def _clear_signal_table(self):
        """Clear all signals from the definitions table."""
        self.signal_defs_panel.clear_signals()

    def _remove_selected_signal(self):
        """Remove the selected signal from the definitions table."""
        table = self.signal_defs_panel.table
        rows = table.selectionModel().selectedRows()
        for index in sorted(rows, reverse=True):
            table.removeRow(index.row())

    # -----------------------------------------------------------------------
    # Export operations
    # -----------------------------------------------------------------------
    def _export_csv(self):
        """Export selected signals to CSV."""
        expressions = self.signal_defs_panel.get_selected_expressions()
        if not expressions:
            # If no signals in table, ask user
            QMessageBox.information(
                self, "Export CSV",
                "Keine Signale ausgewählt.\n"
                "Doppelklicken Sie auf Signale im Signalbaum, um sie zur Liste hinzuzufügen."
            )
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, "CSV exportieren", "", "CSV Dateien (*.csv);;Alle Dateien (*.*)"
        )
        if not filepath:
            return

        self.statusBar().showMessage("Exportiere CSV...")
        try:
            with IbaReader(self._filepath) as reader:
                reader.export_csv(expressions, filepath)
            self.statusBar().showMessage(f"CSV exportiert: {filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"CSV Export fehlgeschlagen:\n\n{e}")
            self.statusBar().showMessage("CSV Export fehlgeschlagen")

    def _export_parquet(self):
        """Export selected signals to Parquet."""
        expressions = self.signal_defs_panel.get_selected_expressions()
        if not expressions:
            QMessageBox.information(
                self, "Export Parquet",
                "Keine Signale ausgewählt.\n"
                "Doppelklicken Sie auf Signale im Signalbaum, um sie zur Liste hinzuzufügen."
            )
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, "Parquet exportieren", "", "Parquet Dateien (*.parquet);;Alle Dateien (*.*)"
        )
        if not filepath:
            return

        self.statusBar().showMessage("Exportiere Parquet...")
        try:
            with IbaReader(self._filepath) as reader:
                reader.export_parquet(expressions, filepath)
            self.statusBar().showMessage(f"Parquet exportiert: {filepath}")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Parquet Export fehlgeschlagen:\n\n{e}")
            self.statusBar().showMessage("Parquet Export fehlgeschlagen")

    def _export_video(self):
        """Export embedded video from .dat file."""
        if not self._filepath:
            return

        filepath, _ = QFileDialog.getSaveFileName(
            self, "Video exportieren", "", "MP4 Dateien (*.mp4);;Alle Dateien (*.*)"
        )
        if not filepath:
            return

        self.statusBar().showMessage("Exportiere Video...")
        try:
            with IbaReader(self._filepath) as reader:
                result = reader.export_video(filepath)
            size_mb = result['size'] / (1024 * 1024)
            self.statusBar().showMessage(
                f"Video exportiert: {result['name']} ({size_mb:.1f} MB)"
            )
        except RuntimeError as e:
            QMessageBox.warning(self, "Video Export", str(e))
            self.statusBar().showMessage("Kein Video gefunden")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", f"Video Export fehlgeschlagen:\n\n{e}")
            self.statusBar().showMessage("Video Export fehlgeschlagen")

    # -----------------------------------------------------------------------
    # Help
    # -----------------------------------------------------------------------
    def _show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self, "Über IBA-Analyzer-GUI",
            "<b>IBA-Analyzer-GUI</b><br><br>"
            "GUI zum Lesen von iba PDA .dat Dateien.<br><br>"
            "Verwendet ibaAnalyzer COM Interface als Backend.<br><br>"
            "© 2026 kasi09 — MIT Lizenz"
        )

    # -----------------------------------------------------------------------
    # State persistence
    # -----------------------------------------------------------------------
    def _restore_state(self):
        """Restore window geometry and state from settings."""
        geometry = self._settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
        state = self._settings.value("windowState")
        if state:
            self.restoreState(state)

    def closeEvent(self, event):
        """Save window state on close."""
        self._settings.setValue("geometry", self.saveGeometry())
        self._settings.setValue("windowState", self.saveState())
        super().closeEvent(event)

    # -----------------------------------------------------------------------
    # Drag & drop support
    # -----------------------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith('.dat'):
                    event.acceptProposedAction()
                    return

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            filepath = url.toLocalFile()
            if filepath.lower().endswith('.dat'):
                self._load_file(filepath)
                return


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    app = QApplication(sys.argv)
    app.setApplicationName("IBA-Analyzer-GUI")
    app.setOrganizationName("kasi09")

    # Accept drag & drop
    window = MainWindow()
    window.setAcceptDrops(True)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
