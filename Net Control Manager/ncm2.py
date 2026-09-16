import sys
import csv
import json
import os
from datetime import datetime

from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QTextDocument, QColor, QKeySequence, QShortcut
from PyQt6.QtPrintSupport import QPrinter
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QTableWidget,
    QTableWidgetItem, QTabWidget, QTextEdit, QGroupBox, QSplitter,
    QFileDialog, QMessageBox, QHeaderView, QFrame, QCheckBox, QDialog,
    QListWidget, QListWidgetItem, QInputDialog, QLayout, QSizePolicy
)

PROFILES_FILE = "profiles.json"

DEFAULT_PROFILES = {
    "Standard ARES Net": {
        "mode": "Regular",
        "station_types": ["Base", "Mobile", "NCS", "Alternate NCS", "NTS", "OES"],
        "traffic_types": ["Question", "Announcement", "Priority", "Emergency", "General Note"],
        "net_statuses": {
            "ACTIVE": {"text": "#000000", "bg": "#00FF00"},
            "CLOSED": {"text": "#00FF00", "bg": "#000000"},
            "STANDBY": {"text": "#FFFF00", "bg": "#000000"}
        },
        "scripts": [
            {
                "title": "Preamble",
                "text": "Welcome to the Amateur Radio Emergency / Ragchew Net. This is [CALLSIGN], acting as Net Control Station for tonight.\\n\\nIs there any emergency or priority traffic on frequency? Please call now."
            }
        ]
    }
}

def generate_stylesheet(font_size=10):
    return f"""
QMainWindow, QWidget, QDialog {{
    background-color: #C0C0C0;
    color: #000000;
    font-family: 'Tahoma', 'MS Sans Serif', sans-serif;
    font-size: {font_size}pt;
}}
QGroupBox {{
    border: 2px solid;
    border-color: #FFFFFF #808080 #808080 #FFFFFF;
    margin-top: 22px;
    font-weight: bold;
    color: #000080;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 8px;
    padding: 0 4px;
    background-color: #C0C0C0;
}}
QLineEdit, QComboBox, QTextEdit, QListWidget {{
    background-color: #FFFFFF;
    color: #000000;
    border: 2px solid;
    border-color: #808080 #FFFFFF #FFFFFF #808080;
    padding: 1px 2px;
    selection-background-color: #000080;
    selection-color: #FFFFFF;
}}
QCheckBox {{
    font-weight: bold;
    color: #000000;
}}
QPushButton {{
    background-color: #C0C0C0;
    color: #000000;
    border: 2px solid;
    border-color: #FFFFFF #808080 #808080 #FFFFFF;
    padding: 2px 4px;
    font-weight: bold;
}}
QPushButton:focus {{
    border: 2px dotted #000000;
}}
QPushButton:pressed {{
    border-color: #808080 #FFFFFF #FFFFFF #808080;
    padding: 3px 3px 1px 5px;
}}
QTableWidget {{
    background-color: #FFFFFF;
    color: #000000;
    gridline-color: #C0C0C0;
    border: 2px solid;
    border-color: #808080 #FFFFFF #FFFFFF #808080;
}}
QHeaderView::section {{
    background-color: #C0C0C0;
    color: #000000;
    padding: 1px 2px;
    border: 2px solid;
    border-color: #FFFFFF #808080 #808080 #FFFFFF;
    font-weight: bold;
}}
QTabWidget::pane {{
    border: 2px solid;
    border-color: #FFFFFF #808080 #808080 #FFFFFF;
    background-color: #C0C0C0;
}}
QTabBar::tab {{
    background-color: #C0C0C0;
    color: #000000;
    padding: 2px 6px;
    border: 2px solid;
    border-color: #FFFFFF #808080 #C0C0C0 #FFFFFF;
    margin-right: 2px;
}}
QTabBar::tab:selected {{
    border-color: #FFFFFF #808080 #C0C0C0 #FFFFFF;
    font-weight: bold;
    background-color: #C0C0C0;
}}
"""

BTN_START_STYLE = """
QPushButton {
    background-color: #008000;
    color: #FFFFFF;
    border: 2px solid;
    border-color: #FFFFFF #004000 #004000 #FFFFFF;
    padding: 2px 4px;
    font-weight: bold;
}
QPushButton:pressed {
    border-color: #004000 #FFFFFF #FFFFFF #004000;
    padding: 3px 3px 1px 5px;
}
"""

BTN_GREEN_STYLE = BTN_START_STYLE

BTN_CLOSE_STYLE = """
QPushButton {
    background-color: #CC0000;
    color: #FFFFFF;
    border: 2px solid;
    border-color: #FFFFFF #660000 #660000 #FFFFFF;
    padding: 2px 4px;
    font-weight: bold;
}
QPushButton:pressed {
    border-color: #660000 #FFFFFF #FFFFFF #660000;
    padding: 3px 3px 1px 5px;
}
"""

BTN_ORANGE_STYLE = """
QPushButton {
    background-color: #FF8C00;
    color: #FFFFFF;
    border: 2px solid;
    border-color: #FFFFFF #8B4500 #8B4500 #FFFFFF;
    padding: 2px 4px;
    font-weight: bold;
}
QPushButton:pressed {
    border-color: #8B4500 #FFFFFF #FFFFFF #8B4500;
    padding: 3px 3px 1px 5px;
}
"""

BTN_CYAN_STYLE = """
QPushButton {
    background-color: #00FFFF;
    color: #000000;
    border: 2px solid;
    border-color: #FFFFFF #008080 #008080 #FFFFFF;
    padding: 2px 4px;
    font-weight: bold;
}
QPushButton:pressed {
    border-color: #008080 #FFFFFF #FFFFFF #008080;
    padding: 3px 3px 1px 5px;
}
"""

BTN_SAVE_STYLE = """
QPushButton {
    background-color: #CCCC00;
    color: #000000;
    border: 2px solid;
    border-color: #FFFFFF #666600 #666600 #FFFFFF;
    padding: 2px 4px;
    font-weight: bold;
}
QPushButton:pressed {
    border-color: #666600 #FFFFFF #FFFFFF #666600;
    padding: 3px 3px 1px 5px;
}
"""

BTN_RESTORE_STYLE = BTN_ORANGE_STYLE
BTN_DANGER_STYLE = BTN_CLOSE_STYLE

def load_profiles_from_disk():
    if not os.path.exists(PROFILES_FILE):
        save_profiles_to_disk(DEFAULT_PROFILES)
        return DEFAULT_PROFILES.copy()
    try:
        with open(PROFILES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading profiles: {e}")
        return DEFAULT_PROFILES.copy()

def save_profiles_to_disk(profiles_data):
    try:
        with open(PROFILES_FILE, "w", encoding="utf-8") as f:
            json.dump(profiles_data, f, indent=2)
    except Exception as e:
        print(f"Error saving profiles: {e}")

def parse_color_map(val):
    if isinstance(val, dict):
        return val.get("text", "#FFFF00"), val.get("bg", "#000000")
    elif isinstance(val, str) and val.startswith("#"):
        return "#000000", val
    return "#FFFF00", "#000000"


class UppercaseLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.textChanged.connect(self.handle_text_changed)

    def handle_text_changed(self, text):
        upper_text = text.upper()
        if text != upper_text:
            pos = self.cursorPosition()
            self.setText(upper_text)
            self.setCursorPosition(pos)


class EnterToggleCheckBox(QCheckBox):
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.animateClick()
        else:
            super().keyPressEvent(event)


class EnterClickButton(QPushButton):
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.animateClick()
        else:
            super().keyPressEvent(event)


class ProfileBadgeLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.Box)
        self.setFrameShadow(QFrame.Shadow.Sunken)
        self.setLineWidth(2)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(220, 38)
        self.setWordWrap(True)
        self.setToolTip("Loaded Profile")

        mono_font = QFont("Consolas", 8, QFont.Weight.Bold)
        mono_font.setStyleHint(QFont.StyleHint.Monospace)
        self.setFont(mono_font)

        self.setStyleSheet("""
            QLabel {
                background-color: #000000;
                color: #00FFFF;
                border: 2px solid #808080;
                padding: 1px 2px;
                font-family: 'Consolas', 'Courier New', monospace;
            }
        """)
        self.set_profile("NONE")

    def set_profile(self, profile_name):
        disp = profile_name.upper()
        if len(disp) > 40:
            disp = disp[:37] + "..."
        self.setText(disp)


class StatusBadgeLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.Box)
        self.setFrameShadow(QFrame.Shadow.Sunken)
        self.setLineWidth(2)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(110, 38)

        mono_font = QFont("Consolas", 8, QFont.Weight.Bold)
        mono_font.setStyleHint(QFont.StyleHint.Monospace)
        self.setFont(mono_font)
        self.set_state("INACTIVE", "#FFFF00", "#000000")

    def set_state(self, state_name, text_hex="#FFFF00", bg_hex="#000000"):
        state_name = state_name.upper()
        self.setText(f"STATUS:\n{state_name}")

        fg = text_hex if text_hex else "#FFFF00"
        bg = bg_hex if bg_hex else "#000000"

        self.setStyleSheet(f"""
            QLabel {{
                background-color: {bg};
                color: {fg};
                border: 2px solid #808080;
                padding: 1px 2px;
                font-family: 'Consolas', 'Courier New', monospace;
            }}
        """)


class DynamicTimerLabel(QLabel):
    INITIAL_TIME_SECONDS = 600

    def __init__(self, parent=None):
        super().__init__(parent)
        self.time_remaining = self.INITIAL_TIME_SECONDS
        self.flash_state = False

        self.setFrameShape(QFrame.Shape.Box)
        self.setFrameShadow(QFrame.Shadow.Sunken)
        self.setLineWidth(2)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("Click box to reset ID timer to 10:00")
        self.setFixedSize(85, 38)

        mono_font = QFont("Consolas", 8, QFont.Weight.Bold)
        mono_font.setStyleHint(QFont.StyleHint.Monospace)
        self.setFont(mono_font)

        self.countdown_timer = QTimer(self)
        self.countdown_timer.setInterval(1000)
        self.countdown_timer.timeout.connect(self.update_countdown)

        self.flash_timer = QTimer(self)
        self.flash_timer.setInterval(1000)
        self.flash_timer.timeout.connect(self.toggle_flash)

        self.apply_normal_style()
        self.update_display_text()

    def update_countdown(self):
        if self.time_remaining > 0:
            self.time_remaining -= 1
            self.update_display_text()

            if self.time_remaining == 0:
                self.start_flashing()

    def update_display_text(self):
        mins = self.time_remaining // 60
        secs = self.time_remaining % 60
        self.setText(f"ID TIMER\n{mins:02d}:{secs:02d}")

    def start_flashing(self):
        self.setText("!!! ID !!!\nNOW")
        if not self.flash_timer.isActive():
            self.flash_timer.start()

    def toggle_flash(self):
        self.flash_state = not self.flash_state
        if self.flash_state:
            self.setStyleSheet("""
                QLabel {
                    background-color: #FFFFFF;
                    color: #FF0000;
                    border: 2px solid #800000;
                    padding: 1px 2px;
                    font-family: 'Consolas', 'Courier New', monospace;
                }
            """)
        else:
            self.setStyleSheet("""
                QLabel {
                    background-color: #FF0000;
                    color: #FFFFFF;
                    border: 2px solid #800000;
                    padding: 1px 2px;
                    font-family: 'Consolas', 'Courier New', monospace;
                }
            """)

    def apply_normal_style(self):
        self.setStyleSheet("""
            QLabel {
                background-color: #000000;
                color: #00FF00;
                border: 2px solid #808080;
                padding: 1px 2px;
                font-family: 'Consolas', 'Courier New', monospace;
            }
        """)

    def reset_timer(self, is_net_active=False):
        self.flash_timer.stop()
        self.time_remaining = self.INITIAL_TIME_SECONDS
        self.update_display_text()
        self.apply_normal_style()

        if is_net_active and not self.countdown_timer.isActive():
            self.countdown_timer.start()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            parent_net_active = False
            main_window = self.window()
            if isinstance(main_window, NetControlApp):
                parent_net_active = main_window.is_net_active

            self.reset_timer(is_net_active=parent_net_active)
        super().mousePressEvent(event)


class NetDurationLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.elapsed_seconds = 0
        self.start_time_str = "--:--"
        self.end_time_str = "--:--"

        self.setFrameShape(QFrame.Shape.Box)
        self.setFrameShadow(QFrame.Shadow.Sunken)
        self.setLineWidth(2)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setFixedSize(235, 38)
        self.setToolTip("Net Duration & Start/End Times")

        mono_font = QFont("Consolas", 8, QFont.Weight.Bold)
        mono_font.setStyleHint(QFont.StyleHint.Monospace)
        self.setFont(mono_font)

        self.duration_timer = QTimer(self)
        self.duration_timer.setInterval(1000)
        self.duration_timer.timeout.connect(self.tick)

        self.apply_style()
        self.update_display()

    def apply_style(self):
        self.setStyleSheet("""
            QLabel {
                background-color: #000000;
                color: #00FFFF;
                border: 2px solid #808080;
                padding: 1px 2px;
                font-family: 'Consolas', 'Courier New', monospace;
            }
        """)

    def tick(self):
        self.elapsed_seconds += 1
        self.update_display()

    def update_display(self):
        hrs = self.elapsed_seconds // 3600
        mins = (self.elapsed_seconds % 3600) // 60
        self.setText(f"DURATION | START {self.start_time_str}\n  {hrs:02d}:{mins:02d}   | END   {self.end_time_str}")

    def start_duration(self):
        self.start_time_str = datetime.now().strftime("%H:%M")
        self.end_time_str = "--:--"
        if not self.duration_timer.isActive():
            self.duration_timer.start()
        self.update_display()

    def stop_duration(self):
        if self.duration_timer.isActive():
            self.end_time_str = datetime.now().strftime("%H:%M")
        self.duration_timer.stop()
        self.update_display()

    def reset_duration(self):
        self.stop_duration()
        self.elapsed_seconds = 0
        self.start_time_str = "--:--"
        self.end_time_str = "--:--"
        self.update_display()


class StationEditDialog(QDialog):
    def __init__(self, station_data, station_types, is_skywarn=False, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Station Details")
        self.setFixedSize(380, 220)

        self.is_skywarn = is_skywarn
        layout = QVBoxLayout(self)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Callsign:"))
        self.input_call = UppercaseLineEdit()
        self.input_call.setText(station_data.get("callsign", ""))
        row1.addWidget(self.input_call)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Station Type:"))
        self.combo_type = QComboBox()
        self.combo_type.addItems(station_types)
        self.combo_type.setCurrentText(station_data.get("type", ""))
        row2.addWidget(self.combo_type)
        layout.addLayout(row2)

        if not is_skywarn:
            row_chk = QHBoxLayout()
            self.chk_traffic = EnterToggleCheckBox("Traffic")
            self.chk_announcement = EnterToggleCheckBox("Announcement")
            
            detail_str = station_data.get("detail", "")
            if "Traffic" in detail_str:
                self.chk_traffic.setChecked(True)
            if "Announcement" in detail_str:
                self.chk_announcement.setChecked(True)

            row_chk.addWidget(self.chk_traffic)
            row_chk.addWidget(self.chk_announcement)
            layout.addLayout(row_chk)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Location / Info:" if is_skywarn else "Info / Grid:"))
        self.input_info = QLineEdit()
        self.input_info.setText(station_data.get("detail" if is_skywarn else "info", ""))
        row3.addWidget(self.input_info)
        layout.addLayout(row3)

        btn_layout = QHBoxLayout()
        btn_save = EnterClickButton("Save")
        btn_save.clicked.connect(self.accept)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def get_updated_data(self):
        call = self.input_call.text().strip().upper()
        st_type = self.combo_type.currentText()
        info_val = self.input_info.text().strip() or "-"

        if self.is_skywarn:
            col_detail = info_val
            col_info = "-"
        else:
            items = []
            if self.chk_traffic.isChecked():
                items.append("Traffic")
            if self.chk_announcement.isChecked():
                items.append("Announcement")
            col_detail = ", ".join(items) if items else "-"
            col_info = info_val

        return call, st_type, col_detail, col_info


class ProfileSelectDialog(QDialog):
    def __init__(self, profiles, current_profile="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Net Profile")
        self.setFixedSize(480, 260)
        self.selected_profile = None

        layout = QVBoxLayout(self)

        lbl = QLabel("Choose Net Control Profile:")
        lbl.setFont(QFont("Tahoma", 9, QFont.Weight.Bold))
        layout.addWidget(lbl)

        lists_layout = QHBoxLayout()

        gen_box = QGroupBox("General Nets")
        gen_layout = QVBoxLayout(gen_box)
        self.list_general = QListWidget()
        gen_layout.addWidget(self.list_general)

        sky_box = QGroupBox("SKYWARN Nets")
        sky_layout = QVBoxLayout(sky_box)
        self.list_skywarn = QListWidget()
        sky_layout.addWidget(self.list_skywarn)

        lists_layout.addWidget(gen_box)
        lists_layout.addWidget(sky_box)
        layout.addLayout(lists_layout)

        for name, data in profiles.items():
            item = QListWidgetItem(name)
            if data.get("mode") == "SKYWARN":
                self.list_skywarn.addItem(item)
                if name == current_profile:
                    self.list_skywarn.setCurrentItem(item)
            else:
                self.list_general.addItem(item)
                if name == current_profile:
                    self.list_general.setCurrentItem(item)

        self.list_general.itemClicked.connect(lambda: self.list_skywarn.clearSelection())
        self.list_skywarn.itemClicked.connect(lambda: self.list_general.clearSelection())

        self.list_general.itemDoubleClicked.connect(lambda item: self.finish_selection(item.text()))
        self.list_skywarn.itemDoubleClicked.connect(lambda item: self.finish_selection(item.text()))

        btn_layout = QHBoxLayout()
        btn_select = EnterClickButton("Select Profile")
        btn_select.clicked.connect(self.accept_selection)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_select)
        btn_layout.addWidget(btn_cancel)

        layout.addLayout(btn_layout)

    def finish_selection(self, profile_name):
        self.selected_profile = profile_name
        self.accept()

    def accept_selection(self):
        selected_sky = self.list_skywarn.selectedItems()
        selected_gen = self.list_general.selectedItems()

        if selected_sky:
            self.selected_profile = selected_sky[0].text()
            self.accept()
        elif selected_gen:
            self.selected_profile = selected_gen[0].text()
            self.accept()


class ProfileEditorDialog(QDialog):
    def __init__(self, profiles, current_profile="", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Net Profile Creator & Editor")
        self.resize(760, 740)
        self.profiles = json.loads(json.dumps(profiles))
        self.active_profile_name = None

        layout = QVBoxLayout(self)

        top_row = QHBoxLayout()
        top_row.addWidget(QLabel("Select Profile to Edit:"))
        self.combo_profiles = QComboBox()
        self.combo_profiles.setFixedWidth(380)
        self.combo_profiles.currentIndexChanged.connect(self.load_selected_profile)
        top_row.addWidget(self.combo_profiles)

        btn_new_prof = QPushButton("New Profile")
        btn_new_prof.clicked.connect(self.create_new_profile)
        btn_del_prof = QPushButton("Delete")
        btn_del_prof.clicked.connect(self.delete_current_profile)

        top_row.addWidget(btn_new_prof)
        top_row.addWidget(btn_del_prof)
        top_row.addStretch()
        layout.addLayout(top_row)

        mode_row = QHBoxLayout()
        mode_row.addWidget(QLabel("Profile Name:"))
        self.input_prof_name = QLineEdit()
        self.input_prof_name.setFixedWidth(380)
        mode_row.addWidget(self.input_prof_name)

        mode_row.addWidget(QLabel("Net Mode:"))
        self.combo_mode = QComboBox()
        self.combo_mode.addItems(["Regular", "SKYWARN"])
        self.combo_mode.currentTextChanged.connect(self.on_mode_changed)
        mode_row.addWidget(self.combo_mode)
        mode_row.addStretch()
        layout.addLayout(mode_row)

        types_row = QHBoxLayout()

        station_group = QGroupBox("Selectable Station Types (One per line)")
        st_layout = QVBoxLayout(station_group)
        self.text_station_types = QTextEdit()
        self.text_station_types.setFixedHeight(90)
        st_layout.addWidget(self.text_station_types)
        types_row.addWidget(station_group)

        traffic_group = QGroupBox("Traffic / Condition Types (One per line)")
        tr_layout = QVBoxLayout(traffic_group)
        self.text_traffic_types = QTextEdit()
        self.text_traffic_types.setFixedHeight(90)
        tr_layout.addWidget(self.text_traffic_types)
        types_row.addWidget(traffic_group)

        layout.addLayout(types_row)

        self.statuses_group = QGroupBox("Net Statuses Color Map (Hex Code e.g. #FF0000 or #FFFF00)")
        statuses_layout = QVBoxLayout(self.statuses_group)

        self.table_statuses = QTableWidget(0, 3)
        self.table_statuses.setHorizontalHeaderLabels(["Status Name", "Text Hex", "Background Hex"])
        self.table_statuses.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table_statuses.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_statuses.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table_statuses.setColumnWidth(0, 180)
        self.table_statuses.setFixedHeight(140)

        st_btns = QHBoxLayout()
        btn_add_st = QPushButton("Add Status")
        btn_add_st.clicked.connect(self.add_status_row)
        btn_del_st = QPushButton("Remove Status")
        btn_del_st.clicked.connect(self.delete_status_row)
        st_btns.addWidget(btn_add_st)
        st_btns.addWidget(btn_del_st)
        st_btns.addStretch()

        statuses_layout.addWidget(self.table_statuses)
        statuses_layout.addLayout(st_btns)
        layout.addWidget(self.statuses_group)

        scripts_group = QGroupBox("Script Tabs Management (Type \\n for line breaks)")
        scripts_layout = QVBoxLayout(scripts_group)

        self.table_scripts = QTableWidget(0, 2)
        self.table_scripts.setHorizontalHeaderLabels(["Tab Title", "Script Preamble Text"])
        self.table_scripts.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table_scripts.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_scripts.setColumnWidth(0, 150)
        self.table_scripts.setMinimumHeight(160)

        sc_btns = QHBoxLayout()
        btn_add_script = QPushButton("Add Tab")
        btn_add_script.clicked.connect(self.add_script_row)
        btn_del_script = QPushButton("Remove Selected Tab")
        btn_del_script.clicked.connect(self.delete_script_row)

        sc_btns.addWidget(btn_add_script)
        sc_btns.addWidget(btn_del_script)
        sc_btns.addStretch()

        scripts_layout.addWidget(self.table_scripts)
        scripts_layout.addLayout(sc_btns)
        layout.addWidget(scripts_group)

        bottom_btns = QHBoxLayout()
        btn_save = EnterClickButton("Save All Profiles")
        btn_save.clicked.connect(self.save_and_close)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        bottom_btns.addStretch()
        bottom_btns.addWidget(btn_save)
        bottom_btns.addWidget(btn_cancel)
        layout.addLayout(bottom_btns)

        self.populate_combo(current_profile)

    def on_mode_changed(self, mode_text):
        self.statuses_group.setEnabled(True)

    def populate_combo(self, select_name=""):
        self.combo_profiles.blockSignals(True)
        self.combo_profiles.clear()
        for name in self.profiles.keys():
            self.combo_profiles.addItem(name)
        self.combo_profiles.blockSignals(False)

        if select_name in self.profiles:
            idx = self.combo_profiles.findText(select_name)
            if idx >= 0:
                self.combo_profiles.setCurrentIndex(idx)

        self.load_selected_profile()

    def load_selected_profile(self):
        if self.active_profile_name and self.active_profile_name in self.profiles:
            self.sync_form_to_data(self.active_profile_name)

        name = self.combo_profiles.currentText()
        if not name or name not in self.profiles:
            self.input_prof_name.clear()
            self.text_station_types.clear()
            self.text_traffic_types.clear()
            self.table_statuses.setRowCount(0)
            self.table_scripts.setRowCount(0)
            self.active_profile_name = None
            return

        self.active_profile_name = name
        prof = self.profiles[name]

        self.input_prof_name.setText(name)
        mode = prof.get("mode", "Regular")
        self.combo_mode.setCurrentText(mode)
        self.on_mode_changed(mode)

        self.text_station_types.setText("\n".join(prof.get("station_types", [])))

        traffic_types = prof.get("traffic_types", ["Question", "Announcement", "Priority", "Emergency", "General Note"])
        self.text_traffic_types.setText("\n".join(traffic_types))

        net_statuses = prof.get("net_statuses", {})
        self.table_statuses.setRowCount(0)
        for st_name, val in net_statuses.items():
            row = self.table_statuses.rowCount()
            self.table_statuses.insertRow(row)
            self.table_statuses.setItem(row, 0, QTableWidgetItem(st_name))

            t_hex, b_hex = parse_color_map(val)
            self.table_statuses.setItem(row, 1, QTableWidgetItem(t_hex))
            self.table_statuses.setItem(row, 2, QTableWidgetItem(b_hex))

        scripts = prof.get("scripts", [])
        self.table_scripts.setRowCount(0)
        for row, sc in enumerate(scripts):
            self.table_scripts.insertRow(row)
            self.table_scripts.setItem(row, 0, QTableWidgetItem(sc.get("title", "")))
            self.table_scripts.setItem(row, 1, QTableWidgetItem(sc.get("text", "")))

    def sync_form_to_data(self, target_name):
        if target_name not in self.profiles:
            return

        st_types = [t.strip() for t in self.text_station_types.toPlainText().split("\n") if t.strip()]
        tr_types = [t.strip() for t in self.text_traffic_types.toPlainText().split("\n") if t.strip()]

        statuses_map = {}
        for r in range(self.table_statuses.rowCount()):
            s_item = self.table_statuses.item(r, 0)
            t_item = self.table_statuses.item(r, 1)
            b_item = self.table_statuses.item(r, 2)

            s_name = s_item.text().strip().upper() if s_item else ""
            t_hex = t_item.text().strip() if t_item else "#FFFF00"
            b_hex = b_item.text().strip() if b_item else "#000000"

            if s_name:
                statuses_map[s_name] = {"text": t_hex, "bg": b_hex}

        if not statuses_map:
            statuses_map = {
                "ACTIVE": {"text": "#000000", "bg": "#00FF00"},
                "CLOSED": {"text": "#00FF00", "bg": "#000000"},
                "STANDBY": {"text": "#FFFF00", "bg": "#000000"}
            }

        scripts = []
        for r in range(self.table_scripts.rowCount()):
            t_item = self.table_scripts.item(r, 0)
            c_item = self.table_scripts.item(r, 1)
            title = t_item.text().strip() if t_item else "Script"
            text = c_item.text() if c_item else ""
            if title:
                scripts.append({"title": title, "text": text})

        self.profiles[target_name] = {
            "mode": self.combo_mode.currentText(),
            "station_types": st_types,
            "traffic_types": tr_types,
            "net_statuses": statuses_map,
            "scripts": scripts
        }

    def create_new_profile(self):
        name, ok = QInputDialog.getText(self, "New Profile", "Enter Profile Name:")
        if ok and name.strip():
            name = name.strip()
            if name in self.profiles:
                QMessageBox.warning(self, "Duplicate Profile", f"Profile '{name}' already exists.")
                return
            self.profiles[name] = {
                "mode": "Regular",
                "station_types": ["Base", "Mobile", "NCS"],
                "traffic_types": ["Question", "Announcement", "Priority", "Emergency", "General Note"],
                "net_statuses": {
                    "ACTIVE": {"text": "#000000", "bg": "#00FF00"},
                    "CLOSED": {"text": "#00FF00", "bg": "#000000"},
                    "STANDBY": {"text": "#FFFF00", "bg": "#000000"}
                },
                "scripts": [{"title": "Preamble", "text": "Welcome to the net..."}]
            }
            self.populate_combo(name)

    def delete_current_profile(self):
        curr = self.combo_profiles.currentText()
        if not curr or curr not in self.profiles:
            return

        if len(self.profiles) <= 1:
            QMessageBox.warning(self, "Cannot Delete", "You must keep at least one profile.")
            return

        confirm = QMessageBox.question(
            self, "Confirm Delete", f"Delete profile '{curr}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if confirm == QMessageBox.StandardButton.Yes:
            del self.profiles[curr]
            self.active_profile_name = None
            self.populate_combo()

    def add_status_row(self):
        row = self.table_statuses.rowCount()
        self.table_statuses.insertRow(row)
        self.table_statuses.setItem(row, 0, QTableWidgetItem("NEW_STATUS"))
        self.table_statuses.setItem(row, 1, QTableWidgetItem("#FFFF00"))
        self.table_statuses.setItem(row, 2, QTableWidgetItem("#000000"))

    def delete_status_row(self):
        row = self.table_statuses.currentRow()
        if row >= 0:
            self.table_statuses.removeRow(row)

    def add_script_row(self):
        row = self.table_scripts.rowCount()
        self.table_scripts.insertRow(row)
        self.table_scripts.setItem(row, 0, QTableWidgetItem(f"Script {row + 1}"))
        self.table_scripts.setItem(row, 1, QTableWidgetItem(""))

    def delete_script_row(self):
        row = self.table_scripts.currentRow()
        if row >= 0:
            self.table_scripts.removeRow(row)

    def save_and_close(self):
        if self.active_profile_name:
            new_name = self.input_prof_name.text().strip()
            if not new_name:
                QMessageBox.warning(self, "Validation Error", "Profile name cannot be empty.")
                return

            if new_name != self.active_profile_name:
                self.profiles[new_name] = self.profiles.pop(self.active_profile_name)
                self.active_profile_name = new_name

            self.sync_form_to_data(new_name)

        save_profiles_to_disk(self.profiles)
        self.accept()


class StationSearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Search Station Roster")
        self.setFixedSize(320, 110)

        layout = QVBoxLayout(self)
        lbl = QLabel("Enter Callsign to Search:")
        lbl.setFont(QFont("Tahoma", 9, QFont.Weight.Bold))
        layout.addWidget(lbl)

        self.input_search = UppercaseLineEdit()
        self.input_search.setPlaceholderText("e.g., W9A")
        self.input_search.returnPressed.connect(self.accept)
        layout.addWidget(self.input_search)

        btn_layout = QHBoxLayout()
        btn_find = EnterClickButton("Find")
        btn_find.clicked.connect(self.accept)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_find)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def get_search_term(self):
        return self.input_search.text().strip().upper()


class TrafficSearchDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Search Traffic Log")
        self.setFixedSize(320, 110)

        layout = QVBoxLayout(self)
        lbl = QLabel("Enter Callsign to Search Traffic:")
        lbl.setFont(QFont("Tahoma", 9, QFont.Weight.Bold))
        layout.addWidget(lbl)

        self.input_search = UppercaseLineEdit()
        self.input_search.setPlaceholderText("e.g., W9A")
        self.input_search.returnPressed.connect(self.accept)
        layout.addWidget(self.input_search)

        btn_layout = QHBoxLayout()
        btn_find = EnterClickButton("Find")
        btn_find.clicked.connect(self.accept)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_find)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def get_search_term(self):
        return self.input_search.text().strip().upper()


class TrafficEditDialog(QDialog):
    def __init__(self, current_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Traffic Details")
        self.setFixedSize(380, 180)

        layout = QVBoxLayout(self)
        lbl = QLabel("Edit Traffic Note / Details:")
        lbl.setFont(QFont("Tahoma", 9, QFont.Weight.Bold))
        layout.addWidget(lbl)

        self.text_edit = QTextEdit()
        self.text_edit.setText(current_text)
        layout.addWidget(self.text_edit)

        btn_layout = QHBoxLayout()
        btn_save = EnterClickButton("Save")
        btn_save.clicked.connect(self.accept)

        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def get_updated_text(self):
        return self.text_edit.toPlainText().strip()


class NetControlApp(QMainWindow):
    AUTOSAVE_FILE = "autosave_session.json"

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Amateur Radio Net Control Manager")
        self.setMinimumSize(1240, 920)
        self.resize(1240, 940)

        self.base_font_size = 10
        self.legends_visible = True

        self.is_net_active = False
        self.net_state_name = "INACTIVE"
        self.current_mode = "Regular"
        self.net_statuses = {}
        self.stations = []
        self.activities = []

        self.profiles = load_profiles_from_disk()
        self.current_profile_name = ""

        self.init_ui()
        self.apply_font_size()
        self.setup_shortcuts()
        self.setup_autosave()

    def apply_font_size(self):
        app = QApplication.instance()
        if app:
            app.setStyleSheet(generate_stylesheet(self.base_font_size))

    def increase_font_size(self):
        if self.base_font_size < 20:
            self.base_font_size += 2
            self.apply_font_size()

    def decrease_font_size(self):
        if self.base_font_size > 8:
            self.base_font_size -= 2
            self.apply_font_size()

    def setup_shortcuts(self):
        self.shortcut_focus_callsign = QShortcut(QKeySequence("F2"), self)
        self.shortcut_focus_callsign.activated.connect(self.focus_callsign_input)

        self.shortcut_search = QShortcut(QKeySequence("F3"), self)
        self.shortcut_search.activated.connect(self.open_search_dialog)

        self.shortcut_edit_station = QShortcut(QKeySequence("Ctrl+E"), self)
        self.shortcut_edit_station.activated.connect(self.edit_station_details)

        self.shortcut_focus_traffic_call = QShortcut(QKeySequence("F4"), self)
        self.shortcut_focus_traffic_call.activated.connect(self.focus_traffic_station_input)

        self.shortcut_traffic_search = QShortcut(QKeySequence("F5"), self)
        self.shortcut_traffic_search.activated.connect(self.open_traffic_search_dialog)

        self.shortcut_select_profile = QShortcut(QKeySequence("F6"), self)
        self.shortcut_select_profile.activated.connect(self.open_profile_select_dialog)

        self.shortcut_edit_profile = QShortcut(QKeySequence("F7"), self)
        self.shortcut_edit_profile.activated.connect(self.open_profile_editor_dialog)

        self.shortcut_toggle_legends = QShortcut(QKeySequence("F12"), self)
        self.shortcut_toggle_legends.activated.connect(self.toggle_legends_visibility)

        self.shortcut_save_session = QShortcut(QKeySequence("Ctrl+S"), self)
        self.shortcut_save_session.activated.connect(self.manual_save_session)

        self.shortcut_reset_net = QShortcut(QKeySequence("Ctrl+R"), self)
        self.shortcut_reset_net.activated.connect(self.reset_entire_net)

        self.shortcut_traffic_now = QShortcut(QKeySequence("Ctrl+N"), self)
        self.shortcut_traffic_now.activated.connect(self.set_traffic_time_now)

        self.shortcut_set_status = QShortcut(QKeySequence("Ctrl+Alt+S"), self)
        self.shortcut_set_status.activated.connect(self.apply_selected_net_status)

        self.shortcut_zoom_in_1 = QShortcut(QKeySequence("Ctrl++"), self)
        self.shortcut_zoom_in_1.activated.connect(self.increase_font_size)
        self.shortcut_zoom_in_2 = QShortcut(QKeySequence("Ctrl+="), self)
        self.shortcut_zoom_in_2.activated.connect(self.increase_font_size)

        self.shortcut_zoom_out_1 = QShortcut(QKeySequence("Ctrl+-"), self)
        self.shortcut_zoom_out_1.activated.connect(self.decrease_font_size)
        self.shortcut_zoom_out_2 = QShortcut(QKeySequence("Ctrl+_"), self)
        self.shortcut_zoom_out_2.activated.connect(self.decrease_font_size)

        self.shortcut_standby = QShortcut(QKeySequence("Ctrl+T"), self)
        self.shortcut_standby.activated.connect(self.toggle_station_standby)

        self.shortcut_checkout = QShortcut(QKeySequence("Ctrl+U"), self)
        self.shortcut_checkout.activated.connect(self.checkout_station)

        self.shortcut_delete_station = QShortcut(QKeySequence("Delete"), self)
        self.shortcut_delete_station.activated.connect(self.delete_station)

        self.shortcut_next_script = QShortcut(QKeySequence("Ctrl+."), self)
        self.shortcut_next_script.activated.connect(self.next_script_tab)

        self.shortcut_prev_script = QShortcut(QKeySequence("Ctrl+,"), self)
        self.shortcut_prev_script.activated.connect(self.prev_script_tab)

        self.shortcut_log_traffic_1 = QShortcut(QKeySequence("Ctrl+Return"), self)
        self.shortcut_log_traffic_1.activated.connect(self.log_traffic_item)
        self.shortcut_log_traffic_2 = QShortcut(QKeySequence("Ctrl+Enter"), self)
        self.shortcut_log_traffic_2.activated.connect(self.log_traffic_item)

        self.shortcut_traffic_edit = QShortcut(QKeySequence("Ctrl+Shift+E"), self)
        self.shortcut_traffic_edit.activated.connect(self.edit_traffic_entry)

        self.shortcut_traffic_delete = QShortcut(QKeySequence("Ctrl+Delete"), self)
        self.shortcut_traffic_delete.activated.connect(self.delete_traffic_entry)

    def setup_autosave(self):
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setInterval(15000)
        self.autosave_timer.timeout.connect(self.autosave_session)
        self.autosave_timer.start()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        main_layout = QVBoxLayout(main_widget)
        main_layout.setSizeConstraint(QLayout.SizeConstraint.SetNoConstraint)
        main_layout.setContentsMargins(4, 4, 4, 4)
        main_layout.setSpacing(3)

        # Top Header
        header_layout = QHBoxLayout()
        header_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        header_layout.setSpacing(6)

        self.profile_badge = ProfileBadgeLabel()

        center_group = QWidget()
        center_layout = QHBoxLayout(center_group)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(6)

        self.status_badge = StatusBadgeLabel()

        self.btn_toggle_net = QPushButton("Start Net\n(Activate)")
        self.btn_toggle_net.setFixedSize(90, 38)
        self.btn_toggle_net.setStyleSheet(BTN_START_STYLE)
        self.btn_toggle_net.clicked.connect(self.toggle_net_mode)

        self.id_timer_label = DynamicTimerLabel()

        center_layout.addWidget(self.status_badge)
        center_layout.addWidget(self.btn_toggle_net)
        center_layout.addWidget(self.id_timer_label)

        self.net_duration_label = NetDurationLabel()

        header_layout.addWidget(self.profile_badge)
        header_layout.addStretch()
        header_layout.addWidget(center_group)
        header_layout.addStretch()
        header_layout.addWidget(self.net_duration_label)

        main_layout.addLayout(header_layout)

        # Main Splitter
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.left_splitter = QSplitter(Qt.Orientation.Vertical)

        # Pane 1: Scripts Viewer
        script_group = QGroupBox("Net Control Scripts")
        script_layout = QVBoxLayout(script_group)
        script_layout.setContentsMargins(3, 3, 3, 3)

        self.script_tabs = QTabWidget()

        self.script_legend_frame = QFrame()
        self.script_legend_frame.setFrameShape(QFrame.Shape.Panel)
        self.script_legend_frame.setFrameShadow(QFrame.Shadow.Sunken)
        script_legend_layout = QHBoxLayout(self.script_legend_frame)
        script_legend_layout.setContentsMargins(4, 1, 4, 1)

        script_legend_label = QLabel("[Ctrl+.] Next Tab  |  [Ctrl+,] Prev Tab")
        script_legend_label.setStyleSheet("color: #000080;")
        script_legend_layout.addWidget(script_legend_label)

        script_layout.addWidget(self.script_tabs)
        script_layout.addWidget(self.script_legend_frame)
        self.left_splitter.addWidget(script_group)

        # Pane 2: Mark Net Traffic Input
        self.traffic_group = QGroupBox("Mark Net Traffic")
        traffic_layout = QVBoxLayout(self.traffic_group)
        traffic_layout.setContentsMargins(3, 3, 3, 3)

        self.t_row1 = QHBoxLayout()
        self.t_row1.setSpacing(3)

        self.lbl_traffic_station = QLabel("Station:")

        self.combo_traffic_call = QComboBox()
        self.combo_traffic_call.setEditable(True)
        self.combo_traffic_call.setLineEdit(UppercaseLineEdit(self.combo_traffic_call))
        self.combo_traffic_call.setPlaceholderText("Callsign")
        self.combo_traffic_call.setFixedWidth(100)

        self.lbl_traffic_type = QLabel("Type:")
        self.combo_traffic_type = QComboBox()
        self.combo_traffic_type.setFixedWidth(150)

        self.t_row1.addWidget(self.lbl_traffic_station)
        self.t_row1.addWidget(self.combo_traffic_call)
        self.t_row1.addSpacing(4)
        self.t_row1.addWidget(self.lbl_traffic_type)
        self.t_row1.addWidget(self.combo_traffic_type)
        self.t_row1.addStretch()

        self.t_row2 = QHBoxLayout()
        self.t_row2.setSpacing(3)

        self.lbl_traffic_time = QLabel("Time:")
        self.input_traffic_time = QLineEdit()
        self.input_traffic_time.setPlaceholderText("HH:MM")
        self.input_traffic_time.setFixedWidth(55)

        self.btn_traffic_now = QPushButton("Now")
        self.btn_traffic_now.clicked.connect(self.set_traffic_time_now)

        self.lbl_traffic_loc = QLabel("Location:")
        self.input_traffic_loc = QLineEdit()
        self.input_traffic_loc.setPlaceholderText("Location / Grid")

        self.t_row2.addWidget(self.lbl_traffic_time)
        self.t_row2.addWidget(self.input_traffic_time)
        self.t_row2.addWidget(self.btn_traffic_now)
        self.t_row2.addSpacing(4)
        self.t_row2.addWidget(self.lbl_traffic_loc)
        self.t_row2.addWidget(self.input_traffic_loc)

        self.input_traffic_note = QTextEdit()
        self.input_traffic_note.setPlaceholderText("Enter description, question details, or weather report notes...")

        t_actions_row = QHBoxLayout()
        self.btn_add_traffic = QPushButton("Log\nTraffic")
        self.btn_add_traffic.setStyleSheet(BTN_GREEN_STYLE)
        self.btn_add_traffic.clicked.connect(self.log_traffic_item)
        t_actions_row.addWidget(self.btn_add_traffic)

        self.status_select_widget = QWidget()
        status_select_layout = QHBoxLayout(self.status_select_widget)
        status_select_layout.setContentsMargins(0, 0, 0, 0)
        status_select_layout.setSpacing(3)

        status_vbox = QVBoxLayout()
        status_vbox.setContentsMargins(0, 0, 0, 0)
        status_vbox.setSpacing(1)

        self.lbl_set_status = QLabel("Net Status:")
        self.combo_net_statuses = QComboBox()
        self.combo_net_statuses.setMinimumWidth(130)

        status_vbox.addWidget(self.lbl_set_status)
        status_vbox.addWidget(self.combo_net_statuses)

        self.btn_set_status = QPushButton("Set\nStatus")
        self.btn_set_status.setStyleSheet(BTN_CYAN_STYLE)
        self.btn_set_status.clicked.connect(self.apply_selected_net_status)

        status_select_layout.addLayout(status_vbox)
        status_select_layout.addWidget(self.btn_set_status)

        t_actions_row.addWidget(self.status_select_widget)

        self.traffic_legend_frame = QFrame()
        self.traffic_legend_frame.setFrameShape(QFrame.Shape.Panel)
        self.traffic_legend_frame.setFrameShadow(QFrame.Shadow.Sunken)
        traffic_legend_layout = QVBoxLayout(self.traffic_legend_frame)
        traffic_legend_layout.setContentsMargins(4, 2, 4, 2)
        traffic_legend_layout.setSpacing(1)

        self.traffic_legend_lbl1 = QLabel("[F4] Focus Station Field  |  [Ctrl+Ent] Log Activity")
        self.traffic_legend_lbl1.setStyleSheet("color: #000080;")
        self.traffic_legend_lbl1.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.traffic_legend_lbl2 = QLabel("[Ctrl+N] Now  |  [Ctrl+Alt+S] Set Status")
        self.traffic_legend_lbl2.setStyleSheet("color: #000080;")
        self.traffic_legend_lbl2.setAlignment(Qt.AlignmentFlag.AlignCenter)

        traffic_legend_layout.addWidget(self.traffic_legend_lbl1)
        traffic_legend_layout.addWidget(self.traffic_legend_lbl2)

        traffic_layout.addLayout(self.t_row1)
        traffic_layout.addLayout(self.t_row2)
        traffic_layout.addWidget(self.input_traffic_note)
        traffic_layout.addLayout(t_actions_row)
        traffic_layout.addWidget(self.traffic_legend_frame)
        self.left_splitter.addWidget(self.traffic_group)

        # Right Column
        self.right_splitter = QSplitter(Qt.Orientation.Vertical)

        # Pane 3: Station Roster
        station_group = QGroupBox("Station Management Roster")
        station_layout = QVBoxLayout(station_group)
        station_layout.setContentsMargins(3, 3, 3, 3)

        self.checkin_layout = QHBoxLayout()
        self.checkin_layout.setSpacing(3)

        self.input_callsign = UppercaseLineEdit()
        self.input_callsign.setPlaceholderText("Callsign (F2)")
        self.input_callsign.returnPressed.connect(self.checkin_station)

        self.lbl_roster_type = QLabel("Type:")
        self.combo_station_type = QComboBox()

        self.chk_traffic = EnterToggleCheckBox("Traffic")
        self.chk_announcement = EnterToggleCheckBox("Announcement")

        self.lbl_roster_loc = QLabel("Info / Grid:")
        self.input_roster_loc = QLineEdit()
        self.input_roster_loc.setPlaceholderText("City / Grid / Info")
        self.input_roster_loc.returnPressed.connect(self.checkin_station)

        btn_checkin = EnterClickButton("Check In")
        btn_checkin.setStyleSheet(BTN_GREEN_STYLE)
        btn_checkin.clicked.connect(self.checkin_station)

        self.checkin_layout.addWidget(self.input_callsign)
        self.checkin_layout.addWidget(self.lbl_roster_type)
        self.checkin_layout.addWidget(self.combo_station_type)
        self.checkin_layout.addWidget(self.chk_traffic)
        self.checkin_layout.addWidget(self.chk_announcement)
        self.checkin_layout.addWidget(self.lbl_roster_loc)
        self.checkin_layout.addWidget(self.input_roster_loc)
        self.checkin_layout.addWidget(btn_checkin)

        self.station_table = QTableWidget(0, 7)
        self.station_table.horizontalHeader().setFixedHeight(22)
        self.station_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        st_actions = QHBoxLayout()
        st_actions.setSpacing(3)

        btn_search_roster = QPushButton("Search\nRoster")
        btn_search_roster.clicked.connect(self.open_search_dialog)

        btn_edit_station = QPushButton("Edit\nStation")
        btn_edit_station.clicked.connect(self.edit_station_details)

        btn_toggle_standby = QPushButton("Toggle\nStatus")
        btn_toggle_standby.clicked.connect(self.toggle_station_standby)

        btn_checkout = QPushButton("Check\nOut")
        btn_checkout.setStyleSheet(BTN_ORANGE_STYLE)
        btn_checkout.clicked.connect(self.checkout_station)

        btn_delete_station = QPushButton("Delete\nStation")
        btn_delete_station.setStyleSheet(BTN_DANGER_STYLE)
        btn_delete_station.clicked.connect(self.delete_station)

        btn_clear_roster = QPushButton("Clear\nRoster")
        btn_clear_roster.clicked.connect(self.clear_station_roster)

        # Integrated double-row Hotkeys Frame into the actions bar
        self.legend_frame = QFrame()
        self.legend_frame.setFrameShape(QFrame.Shape.Panel)
        self.legend_frame.setFrameShadow(QFrame.Shadow.Sunken)
        legend_layout = QVBoxLayout(self.legend_frame)
        legend_layout.setContentsMargins(4, 1, 4, 1)
        legend_layout.setSpacing(1)

        legend_label1 = QLabel("[F2] Add  |  [F3] Search  |  [Ctrl+E] Edit")
        legend_label1.setStyleSheet("color: #000080;")
        legend_label1.setAlignment(Qt.AlignmentFlag.AlignCenter)

        legend_label2 = QLabel("[Ctrl+T] Standby  |  [Ctrl+U] Out  |  [Del] Delete")
        legend_label2.setStyleSheet("color: #000080;")
        legend_label2.setAlignment(Qt.AlignmentFlag.AlignCenter)

        legend_layout.addWidget(legend_label1)
        legend_layout.addWidget(legend_label2)

        counter_frame = QFrame()
        counter_frame.setFrameShape(QFrame.Shape.Panel)
        counter_frame.setFrameShadow(QFrame.Shadow.Sunken)
        counter_frame.setFixedWidth(115)

        counter_layout = QHBoxLayout(counter_frame)
        counter_layout.setContentsMargins(4, 1, 4, 1)

        self.lbl_station_counts = QLabel("Total: 0\nMobile: 0")
        self.lbl_station_counts.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        mono_font = QFont("Consolas", 8, QFont.Weight.Bold)
        mono_font.setStyleHint(QFont.StyleHint.Monospace)
        self.lbl_station_counts.setFont(mono_font)
        self.lbl_station_counts.setStyleSheet("color: #000080;")
        counter_layout.addWidget(self.lbl_station_counts)

        st_actions.addWidget(btn_search_roster)
        st_actions.addWidget(btn_edit_station)
        st_actions.addWidget(btn_toggle_standby)
        st_actions.addWidget(btn_checkout)
        st_actions.addWidget(btn_delete_station)
        st_actions.addWidget(btn_clear_roster)
        st_actions.addWidget(self.legend_frame)
        st_actions.addStretch()
        st_actions.addWidget(counter_frame)

        station_layout.addLayout(self.checkin_layout)
        station_layout.addWidget(self.station_table)
        station_layout.addLayout(st_actions)

        self.right_splitter.addWidget(station_group)

        # Pane 4: Activity Log Box
        self.activity_group = QGroupBox("Live Net Activity & Traffic Log")
        activity_layout = QVBoxLayout(self.activity_group)
        activity_layout.setContentsMargins(3, 3, 3, 3)

        self.activity_table = QTableWidget(0, 5)
        self.activity_table.horizontalHeader().setFixedHeight(22)
        self.activity_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        act_actions = QHBoxLayout()
        act_actions.setSpacing(3)
        btn_traffic_search = QPushButton("Search Traffic")
        btn_traffic_search.clicked.connect(self.open_traffic_search_dialog)

        btn_traffic_edit = QPushButton("Edit Details")
        btn_traffic_edit.clicked.connect(self.edit_traffic_entry)

        btn_traffic_delete = QPushButton("Delete Entry")
        btn_traffic_delete.clicked.connect(self.delete_traffic_entry)

        btn_clear_log = QPushButton("Clear Log")
        btn_clear_log.clicked.connect(self.clear_activity_log)

        act_actions.addWidget(btn_traffic_search)
        act_actions.addWidget(btn_traffic_edit)
        act_actions.addWidget(btn_traffic_delete)
        act_actions.addWidget(btn_clear_log)
        act_actions.addStretch()

        self.activity_legend_frame = QFrame()
        self.activity_legend_frame.setFrameShape(QFrame.Shape.Panel)
        self.activity_legend_frame.setFrameShadow(QFrame.Shadow.Sunken)
        activity_legend_layout = QHBoxLayout(self.activity_legend_frame)
        activity_legend_layout.setContentsMargins(4, 1, 4, 1)

        activity_legend_label = QLabel("[F5] Search Traffic  |  [Ctrl+Shift+E] Edit Details  |  [Ctrl+Del] Delete Entry")
        activity_legend_label.setStyleSheet("color: #000080;")
        activity_legend_layout.addWidget(activity_legend_label)

        activity_layout.addWidget(self.activity_table)
        activity_layout.addLayout(act_actions)
        activity_layout.addWidget(self.activity_legend_frame)
        self.right_splitter.addWidget(self.activity_group)

        self.left_splitter.setSizes([460, 460])
        self.left_splitter.setStretchFactor(0, 1)
        self.left_splitter.setStretchFactor(1, 1)

        self.right_splitter.setSizes([460, 460])
        self.right_splitter.setStretchFactor(0, 1)
        self.right_splitter.setStretchFactor(1, 1)

        self.main_splitter.addWidget(self.left_splitter)
        self.main_splitter.addWidget(self.right_splitter)
        self.main_splitter.setSizes([380, 860])

        main_layout.addWidget(self.main_splitter, 1)

        # Footer (Converted to roster button dynamic sizing using \n titles)
        footer_layout = QHBoxLayout()
        footer_layout.setSpacing(4)

        self.btn_save_session = QPushButton("Save\nSession")
        self.btn_save_session.setStyleSheet(BTN_SAVE_STYLE)
        self.btn_save_session.clicked.connect(self.manual_save_session)

        self.btn_restore = QPushButton("Restore\nSession")
        self.btn_restore.setStyleSheet(BTN_RESTORE_STYLE)
        self.btn_restore.clicked.connect(self.restore_session)

        self.footer_legend_frame = QFrame()
        self.footer_legend_frame.setFrameShape(QFrame.Shape.Panel)
        self.footer_legend_frame.setFrameShadow(QFrame.Shadow.Sunken)

        self.footer_legend_layout = QVBoxLayout(self.footer_legend_frame)
        self.footer_legend_layout.setContentsMargins(6, 2, 6, 2)
        self.footer_legend_layout.setSpacing(2)

        self.footer_legend_lbl1 = QLabel("[F12] Hide Legends  |  [F6] Select Profile  |  [F7] Create Profile  |  [Ctrl+S] Save Session")
        self.footer_legend_lbl1.setStyleSheet("color: #000080;")
        self.footer_legend_lbl1.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.footer_legend_lbl2 = QLabel("[Ctrl+] Zoom In  |  [Ctrl-] Zoom Out  |  [Ctrl+R] Reset Net  |  [Alt+F4] Close")
        self.footer_legend_lbl2.setStyleSheet("color: #000080;")
        self.footer_legend_lbl2.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.footer_legend_layout.addWidget(self.footer_legend_lbl1)
        self.footer_legend_layout.addWidget(self.footer_legend_lbl2)

        self.btn_export_csv = QPushButton("Export\nCSV")
        self.btn_export_csv.clicked.connect(self.export_csv)

        self.btn_export_pdf = QPushButton("Export\nPDF")
        self.btn_export_pdf.clicked.connect(self.export_pdf)

        self.btn_reset = QPushButton("Reset\nNet")
        self.btn_reset.setStyleSheet(BTN_DANGER_STYLE)
        self.btn_reset.clicked.connect(self.reset_entire_net)

        self.btn_exit = QPushButton("Exit\nApp")
        self.btn_exit.setStyleSheet(BTN_DANGER_STYLE)
        self.btn_exit.clicked.connect(self.close)

        footer_layout.addWidget(self.btn_save_session)
        footer_layout.addWidget(self.btn_restore)
        footer_layout.addStretch()
        footer_layout.addWidget(self.footer_legend_frame)
        footer_layout.addStretch()
        footer_layout.addWidget(self.btn_export_csv)
        footer_layout.addWidget(self.btn_export_pdf)
        footer_layout.addWidget(self.btn_reset)
        footer_layout.addWidget(self.btn_exit)

        main_layout.addLayout(footer_layout)

    def toggle_legends_visibility(self):
        self.legends_visible = not self.legends_visible

        self.script_legend_frame.setVisible(self.legends_visible)
        self.traffic_legend_frame.setVisible(self.legends_visible)
        self.legend_frame.setVisible(self.legends_visible)
        self.activity_legend_frame.setVisible(self.legends_visible)

        if self.legends_visible:
            self.footer_legend_lbl1.setText("[F12] Hide Legends  |  [F6] Select Profile  |  [F7] Create Profile  |  [Ctrl+S] Save Session")
            self.footer_legend_lbl2.setText("[Ctrl+] Zoom In  |  [Ctrl-] Zoom Out  |  [Ctrl+R] Reset Net  |  [Alt+F4] Close")
            self.footer_legend_lbl2.setVisible(True)
        else:
            self.footer_legend_lbl1.setText("[F12] Display Legends")
            self.footer_legend_lbl2.setVisible(False)

    def rebuild_ui_for_mode(self, mode_name, traffic_types, net_statuses):
        self.current_mode = mode_name
        self.net_statuses = net_statuses

        self.combo_traffic_type.clear()
        self.combo_traffic_type.addItems(traffic_types)

        self.station_table.clear()
        self.activity_table.clear()

        if mode_name == "SKYWARN":
            self.traffic_group.setTitle("Spotter Reports and Net Traffic")
            self.activity_group.setTitle("SKYWARN Activity Log")

            self.lbl_traffic_time.setVisible(True)
            self.input_traffic_time.setVisible(True)
            self.btn_traffic_now.setVisible(True)
            self.lbl_traffic_loc.setVisible(True)
            self.input_traffic_loc.setVisible(True)

            self.lbl_traffic_type.setText("Condition:")

            self.status_select_widget.setVisible(True)
            self.combo_net_statuses.clear()
            self.combo_net_statuses.addItems(list(net_statuses.keys()))

            self.traffic_legend_lbl2.setVisible(True)

            self.chk_traffic.setVisible(False)
            self.chk_announcement.setVisible(False)
            self.lbl_roster_loc.setText("Location:")

            self.station_table.setColumnCount(6)
            self.station_table.setHorizontalHeaderLabels(["Call", "Station Type", "Location", "Status", "In", "Out"])

            self.station_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(0, 60)

            self.station_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
            self.station_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)

            self.station_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(3, 90)

            self.station_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(4, 50)

            self.station_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(5, 50)

            self.activity_table.setColumnCount(5)
            self.activity_table.setHorizontalHeaderLabels(["Time", "Condition", "Location", "Source", "Narrative"])
            self.activity_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        else:
            self.traffic_group.setTitle("Mark Net Traffic")
            self.activity_group.setTitle("Live Net Activity & Traffic Log")

            self.lbl_traffic_time.setVisible(False)
            self.input_traffic_time.setVisible(False)
            self.btn_traffic_now.setVisible(False)
            self.lbl_traffic_loc.setVisible(False)
            self.input_traffic_loc.setVisible(False)

            self.lbl_traffic_type.setText("Type:")

            self.status_select_widget.setVisible(False)
            self.traffic_legend_lbl2.setVisible(False)

            self.chk_traffic.setVisible(True)
            self.chk_announcement.setVisible(True)
            self.lbl_roster_loc.setText("Info / Grid:")

            self.station_table.setColumnCount(7)
            self.station_table.setHorizontalHeaderLabels(["Call", "Station Type", "Additional", "Info / Grid", "Status", "In", "Out"])

            self.station_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(0, 60)

            self.station_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)

            self.station_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(2, 110)

            self.station_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

            self.station_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(4, 90)

            self.station_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(5, 50)

            self.station_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)
            self.station_table.setColumnWidth(6, 50)

            self.activity_table.setColumnCount(4)
            self.activity_table.setHorizontalHeaderLabels(["Time", "Callsign", "Traffic Tag", "Details"])
            self.activity_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            self.activity_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        self.refresh_station_table()

    def set_traffic_time_now(self):
        if self.current_mode == "SKYWARN":
            self.input_traffic_time.setText(datetime.now().strftime("%H:%M"))

    def apply_selected_net_status(self):
        if not self.status_select_widget.isVisible():
            return

        st_name = self.combo_net_statuses.currentText()
        if not st_name:
            return

        val = self.net_statuses.get(st_name, {})
        t_hex, b_hex = parse_color_map(val)

        self.net_state_name = st_name
        self.status_badge.set_state(st_name, t_hex, b_hex)
        self.add_activity_log("NET CONTROL", "STATUS CHANGE", "-", f"=== NET STATUS SET TO {st_name} ===")

        if self.current_mode == "SKYWARN":
            st_clean = st_name.strip().lower()
            matched = False

            for i in range(self.script_tabs.count()):
                tab_title = self.script_tabs.tabText(i).strip().lower()
                if tab_title == st_clean:
                    self.script_tabs.setCurrentIndex(i)
                    matched = True
                    break

            if not matched:
                for i in range(self.script_tabs.count()):
                    tab_title = self.script_tabs.tabText(i).strip().lower()
                    if st_clean in tab_title or tab_title in st_clean:
                        self.script_tabs.setCurrentIndex(i)
                        break

    def prompt_profile_on_startup(self):
        self.profiles = load_profiles_from_disk()

        if not self.profiles:
            QMessageBox.information(self, "No Profiles", "No profiles found. Please create a profile to continue.")
            self.open_profile_editor_dialog()
            self.profiles = load_profiles_from_disk()

            if not self.profiles:
                sys.exit(0)

        dialog = ProfileSelectDialog(self.profiles, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected_profile:
            self.apply_profile(dialog.selected_profile)
        else:
            first_name = list(self.profiles.keys())[0]
            self.apply_profile(first_name)

    def apply_profile(self, profile_name):
        if profile_name not in self.profiles:
            return

        self.stations.clear()
        self.activities.clear()
        self.activity_table.setRowCount(0)

        self.id_timer_label.reset_timer(is_net_active=False)
        self.net_duration_label.reset_duration()

        self.is_net_active = False
        self.btn_toggle_net.setText("Start Net\n(Activate)")
        self.btn_toggle_net.setStyleSheet(BTN_START_STYLE)

        self.current_profile_name = profile_name
        self.profile_badge.set_profile(profile_name)

        prof = self.profiles[profile_name]
        mode = prof.get("mode", "Regular")
        tr_types = prof.get("traffic_types", ["Question", "Announcement", "Priority", "Emergency", "General Note"])
        st_map = prof.get("net_statuses", {})

        self.rebuild_ui_for_mode(mode, tr_types, st_map)

        self.combo_station_type.clear()
        self.combo_station_type.addItems(prof.get("station_types", []))

        self.script_tabs.clear()
        for sc in prof.get("scripts", []):
            self.script_tabs.addTab(
                self.create_script_tab(sc.get("text", "")),
                sc.get("title", "Script")
            )

        if mode == "SKYWARN":
            if self.combo_net_statuses.count() > 0:
                self.apply_selected_net_status()
        else:
            self.net_state_name = "INACTIVE"
            self.status_badge.set_state("INACTIVE", "#FFFF00", "#000000")

        self.refresh_callsign_combo()
        self.main_splitter.setSizes([380, 860])

    def open_profile_select_dialog(self):
        self.profiles = load_profiles_from_disk()
        dialog = ProfileSelectDialog(self.profiles, self.current_profile_name, self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected_profile:
            self.apply_profile(dialog.selected_profile)

    def open_profile_editor_dialog(self):
        self.profiles = load_profiles_from_disk()
        dialog = ProfileEditorDialog(self.profiles, self.current_profile_name, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.profiles = load_profiles_from_disk()
            if self.current_profile_name in self.profiles:
                self.apply_profile(self.current_profile_name)
            elif self.profiles:
                self.apply_profile(list(self.profiles.keys())[0])

    def focus_callsign_input(self):
        self.input_callsign.setFocus()
        self.input_callsign.selectAll()

    def focus_traffic_station_input(self):
        self.combo_traffic_call.setFocus()

    def next_script_tab(self):
        count = self.script_tabs.count()
        if count > 0:
            curr = self.script_tabs.currentIndex()
            self.script_tabs.setCurrentIndex((curr + 1) % count)

    def prev_script_tab(self):
        count = self.script_tabs.count()
        if count > 0:
            curr = self.script_tabs.currentIndex()
            self.script_tabs.setCurrentIndex((curr - 1 + count) % count)

    def edit_station_details(self):
        selected = self.station_table.currentRow()
        if selected < 0 or selected >= len(self.stations):
            QMessageBox.information(self, "Edit Station", "Please select a station from the roster to edit.")
            return

        st = self.stations[selected]
        prof = self.profiles.get(self.current_profile_name, {})
        st_types = prof.get("station_types", ["Base", "Mobile", "NCS"])

        dialog = StationEditDialog(st, st_types, is_skywarn=(self.current_mode == "SKYWARN"), parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            call, st_type, col_detail, col_info = dialog.get_updated_data()
            if not call:
                QMessageBox.warning(self, "Input Error", "Callsign cannot be empty.")
                return

            st["callsign"] = call
            st["type"] = st_type
            st["detail"] = col_detail
            st["info"] = col_info

            self.refresh_station_table()
            self.refresh_callsign_combo()
            self.station_table.selectRow(selected)

    def delete_station(self):
        focused = QApplication.focusWidget()
        if isinstance(focused, (QLineEdit, QTextEdit)):
            return

        selected = self.station_table.currentRow()
        if selected < 0 or selected >= len(self.stations):
            QMessageBox.information(self, "Delete Station", "Please select a station from the roster to delete.")
            return

        st = self.stations[selected]
        confirm = QMessageBox.question(
            self, "Confirm Delete", f"Are you sure you want to delete station '{st['callsign']}' from the roster?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.stations.pop(selected)
            self.refresh_station_table()
            self.refresh_callsign_combo()

    def clear_station_roster(self):
        if not self.stations:
            return

        confirm = QMessageBox.question(
            self, "Confirm Clear Roster", "Are you sure you want to clear all checked-in stations from the roster?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.stations.clear()
            self.refresh_station_table()
            self.refresh_callsign_combo()

    def open_traffic_search_dialog(self):
        dialog = TrafficSearchDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            term = dialog.get_search_term()
            if not term:
                return

            found = False
            for row, act in enumerate(self.activities):
                call = act.get("callsign", act.get("source", ""))
                if term in call:
                    self.activity_table.selectRow(row)
                    item = self.activity_table.item(row, 0)
                    if item:
                        self.activity_table.scrollToItem(item)
                    found = True
                    break

            if not found:
                QMessageBox.information(self, "Search Result", f"No traffic entry found for callsign '{term}'.")

    def edit_traffic_entry(self):
        selected = self.activity_table.currentRow()
        if selected < 0 or selected >= len(self.activities):
            QMessageBox.information(self, "Edit Traffic", "Please select a traffic entry to edit.")
            return

        act = self.activities[selected]
        dialog = TrafficEditDialog(act.get("note", act.get("narrative", "")), self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_note = dialog.get_updated_text()
            if new_note:
                if self.current_mode == "SKYWARN":
                    act["narrative"] = new_note
                    self.activity_table.setItem(selected, 4, QTableWidgetItem(new_note))
                else:
                    act["note"] = new_note
                    self.activity_table.setItem(selected, 3, QTableWidgetItem(new_note))

    def delete_traffic_entry(self):
        selected = self.activity_table.currentRow()
        if selected < 0 or selected >= len(self.activities):
            QMessageBox.information(self, "Delete Traffic", "Please select a traffic entry to delete.")
            return

        call = self.activities[selected].get("callsign", self.activities[selected].get("source", ""))
        confirm = QMessageBox.question(
            self, "Confirm Delete", f"Are you sure you want to delete the selected entry for {call}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.activities.pop(selected)
            self.activity_table.removeRow(selected)

    def clear_activity_log(self):
        if not self.activities:
            return

        confirm = QMessageBox.question(
            self, "Confirm Clear Log", "Are you sure you want to clear the entire activity log?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.activities.clear()
            self.activity_table.setRowCount(0)

    def reset_entire_net(self):
        confirm = QMessageBox.question(
            self, "Confirm Net Reset", "Are you sure you want to reset the entire net?\n\nThis will clear all stations, activity logs, and reset session timers.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if confirm == QMessageBox.StandardButton.Yes:
            self.stations.clear()
            self.activities.clear()
            self.refresh_station_table()
            self.refresh_callsign_combo()
            self.activity_table.setRowCount(0)

            self.id_timer_label.reset_timer(is_net_active=False)
            self.net_duration_label.reset_duration()

            if self.is_net_active:
                self.toggle_net_mode()

            val = self.net_statuses.get("INACTIVE", {})
            t_hex, b_hex = parse_color_map(val)
            self.status_badge.set_state("INACTIVE", t_hex, b_hex)

    def autosave_session(self):
        data = {
            "is_net_active": self.is_net_active,
            "net_state_name": self.net_state_name,
            "stations": self.stations,
            "activities": self.activities,
            "time_remaining": self.id_timer_label.time_remaining,
            "elapsed_seconds": self.net_duration_label.elapsed_seconds,
            "start_time_str": self.net_duration_label.start_time_str,
            "end_time_str": self.net_duration_label.end_time_str,
            "current_profile": self.current_profile_name,
            "base_font_size": self.base_font_size
        }
        try:
            with open(self.AUTOSAVE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Autosave Error: {e}")

    def manual_save_session(self):
        self.autosave_session()
        QMessageBox.information(self, "Save Session", "Current net session state successfully saved to disk.")

    def restore_session(self):
        if not os.path.exists(self.AUTOSAVE_FILE):
            QMessageBox.warning(self, "Restore Session", "No autosave session file found.")
            return

        try:
            with open(self.AUTOSAVE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.is_net_active = data.get("is_net_active", False)
            self.net_state_name = data.get("net_state_name", "INACTIVE")
            self.stations = data.get("stations", [])
            self.activities = data.get("activities", [])

            saved_font_size = data.get("base_font_size", 10)
            self.base_font_size = saved_font_size
            self.apply_font_size()

            saved_prof = data.get("current_profile", "")
            if saved_prof and saved_prof in self.profiles:
                self.current_profile_name = saved_prof
                self.profile_badge.set_profile(saved_prof)

                prof = self.profiles[saved_prof]
                mode = prof.get("mode", "Regular")
                tr_types = prof.get("traffic_types", ["Question", "Announcement", "Priority", "Emergency", "General Note"])
                st_map = prof.get("net_statuses", {})

                self.rebuild_ui_for_mode(mode, tr_types, st_map)

                self.combo_station_type.clear()
                self.combo_station_type.addItems(prof.get("station_types", []))

                self.script_tabs.clear()
                for sc in prof.get("scripts", []):
                    self.script_tabs.addTab(
                        self.create_script_tab(sc.get("text", "")),
                        sc.get("title", "Script")
                    )

            time_rem = data.get("time_remaining", 600)
            self.id_timer_label.time_remaining = time_rem
            self.id_timer_label.update_display_text()

            self.net_duration_label.elapsed_seconds = data.get("elapsed_seconds", 0)
            self.net_duration_label.start_time_str = data.get("start_time_str", "--:--")
            self.net_duration_label.end_time_str = data.get("end_time_str", "--:--")
            self.net_duration_label.update_display()

            val = self.net_statuses.get(self.net_state_name, {})
            t_hex, b_hex = parse_color_map(val)
            self.status_badge.set_state(self.net_state_name, t_hex, b_hex)

            if self.is_net_active:
                self.btn_toggle_net.setText("Close Net")
                self.btn_toggle_net.setStyleSheet(BTN_CLOSE_STYLE)
                if not self.net_duration_label.duration_timer.isActive():
                    self.net_duration_label.duration_timer.start()
                self.id_timer_label.countdown_timer.start()
            else:
                self.btn_toggle_net.setText("Start Net\n(Activate)")
                self.btn_toggle_net.setStyleSheet(BTN_START_STYLE)
                self.net_duration_label.stop_duration()
                self.id_timer_label.countdown_timer.stop()

            self.refresh_station_table()
            self.refresh_callsign_combo()

            self.activity_table.setRowCount(0)
            for act in self.activities:
                row = self.activity_table.rowCount()
                self.activity_table.insertRow(row)

                if self.current_mode == "SKYWARN":
                    self.activity_table.setItem(row, 0, QTableWidgetItem(act.get("timestamp", "")))
                    self.activity_table.setItem(row, 1, QTableWidgetItem(act.get("condition", act.get("type", ""))))
                    self.activity_table.setItem(row, 2, QTableWidgetItem(act.get("location", "-")))
                    self.activity_table.setItem(row, 3, QTableWidgetItem(act.get("source", act.get("callsign", ""))))
                    self.activity_table.setItem(row, 4, QTableWidgetItem(act.get("narrative", act.get("note", ""))))
                else:
                    self.activity_table.setItem(row, 0, QTableWidgetItem(act.get("timestamp", "")))
                    self.activity_table.setItem(row, 1, QTableWidgetItem(act.get("callsign", "")))
                    self.activity_table.setItem(row, 2, QTableWidgetItem(act.get("type", "")))
                    self.activity_table.setItem(row, 3, QTableWidgetItem(act.get("note", "")))

            QMessageBox.information(self, "Restore Session", "Session successfully restored from autosave!")
        except Exception as e:
            QMessageBox.critical(self, "Restore Error", f"Failed to restore session:\n{str(e)}")

    def open_search_dialog(self):
        dialog = StationSearchDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            term = dialog.get_search_term()
            if not term:
                return

            found = False
            for row, st in enumerate(self.stations):
                if term in st["callsign"]:
                    self.station_table.selectRow(row)
                    item = self.station_table.item(row, 0)
                    if item:
                        self.station_table.scrollToItem(item)
                    found = True
                    break

            if not found:
                QMessageBox.information(self, "Search Result", f"Callsign '{term}' not found in roster.")

    def create_script_tab(self, default_text):
        editor = QTextEdit()
        formatted_text = str(default_text).replace("\\n", "\n")
        editor.setPlainText(formatted_text)
        editor.setReadOnly(True)
        return editor

    def toggle_net_mode(self):
        self.is_net_active = not self.is_net_active
        if self.is_net_active:
            self.net_state_name = "ACTIVE"
            val = self.net_statuses.get("ACTIVE", {})
            t_hex, b_hex = parse_color_map(val)

            self.status_badge.set_state("ACTIVE", t_hex, b_hex)
            self.btn_toggle_net.setText("Close Net")
            self.btn_toggle_net.setStyleSheet(BTN_CLOSE_STYLE)
            self.add_activity_log("NET CONTROL", "General Note", "-", "=== NET SESSION STARTED ===")

            self.id_timer_label.reset_timer(is_net_active=True)
            self.net_duration_label.reset_duration()
            self.net_duration_label.start_duration()
        else:
            self.net_state_name = "CLOSED"
            val = self.net_statuses.get("CLOSED", {})
            t_hex, b_hex = parse_color_map(val)

            self.status_badge.set_state("CLOSED", t_hex, b_hex)
            self.btn_toggle_net.setText("Start Net\n(Activate)")
            self.btn_toggle_net.setStyleSheet(BTN_START_STYLE)
            self.add_activity_log("NET CONTROL", "General Note", "-", "=== NET SESSION PLACED ON STANDBY / CLOSED ===")

            self.id_timer_label.countdown_timer.stop()
            self.net_duration_label.stop_duration()

    def format_additional_attributes(self, has_traffic, has_announcement):
        items = []
        if has_traffic:
            items.append("Traffic")
        if has_announcement:
            items.append("Announcement")
        return ", ".join(items) if items else "-"

    def checkin_station(self):
        call = self.input_callsign.text().strip().upper()
        st_type = self.combo_station_type.currentText()
        now_str = datetime.now().strftime("%H:%M")
        info_val = self.input_roster_loc.text().strip() or "-"

        if not call:
            QMessageBox.warning(self, "Input Error", "Callsign is required for check-in.")
            return

        if self.current_mode == "SKYWARN":
            col_detail = info_val
            col_info = "-"
        else:
            has_traffic = self.chk_traffic.isChecked()
            has_announcement = self.chk_announcement.isChecked()
            col_detail = self.format_additional_attributes(has_traffic, has_announcement)
            col_info = info_val

        for st in self.stations:
            if st["callsign"] == call:
                st["type"] = st_type
                st["detail"] = col_detail
                st["info"] = col_info
                st["status"] = "Active"
                st["time_in"] = now_str
                st["time_out"] = ""
                self.refresh_station_table()
                self.refresh_callsign_combo()
                self.clear_checkin_inputs()
                return

        station_data = {
            "callsign": call,
            "type": st_type,
            "detail": col_detail,
            "info": col_info,
            "status": "Active",
            "time_in": now_str,
            "time_out": ""
        }
        self.stations.append(station_data)
        self.refresh_station_table()
        self.refresh_callsign_combo()
        self.clear_checkin_inputs()

    def clear_checkin_inputs(self):
        self.input_callsign.clear()
        self.input_roster_loc.clear()
        if self.combo_station_type.count() > 0:
            self.combo_station_type.setCurrentIndex(0)
        self.chk_traffic.setChecked(False)
        self.chk_announcement.setChecked(False)
        self.input_callsign.setFocus()

    def refresh_station_table(self):
        self.station_table.setRowCount(0)
        total_count = len(self.stations)
        mobile_count = 0

        self.station_table.setUpdatesEnabled(False)
        try:
            for row, st in enumerate(self.stations):
                self.station_table.insertRow(row)

                st_type = st["type"]
                if "Mobile" in st_type:
                    mobile_count += 1

                if self.current_mode == "SKYWARN":
                    items = [
                        QTableWidgetItem(st["callsign"]),
                        QTableWidgetItem(st["type"]),
                        QTableWidgetItem(st.get("detail", "-")),
                        QTableWidgetItem(st["status"]),
                        QTableWidgetItem(st["time_in"]),
                        QTableWidgetItem(st["time_out"])
                    ]
                    status_col_idx = 3
                else:
                    items = [
                        QTableWidgetItem(st["callsign"]),
                        QTableWidgetItem(st["type"]),
                        QTableWidgetItem(st.get("detail", "-")),
                        QTableWidgetItem(st.get("info", "-")),
                        QTableWidgetItem(st["status"]),
                        QTableWidgetItem(st["time_in"]),
                        QTableWidgetItem(st["time_out"])
                    ]
                    status_col_idx = 4

                status = st["status"]
                for col_idx, item in enumerate(items):
                    if status == "Standby":
                        item.setBackground(QColor("#FFFF00"))
                        item.setForeground(QColor("#000000"))
                    elif status == "Checked Out":
                        item.setBackground(QColor("#CC0000"))
                        item.setForeground(QColor("#FFFFFF"))
                    else:
                        item.setBackground(QColor("#FFFFFF"))
                        if col_idx == status_col_idx:
                            item.setForeground(QColor("#008000"))
                        else:
                            item.setForeground(QColor("#000000"))

                    self.station_table.setItem(row, col_idx, item)
        finally:
            self.station_table.setUpdatesEnabled(True)

        self.lbl_station_counts.setText(f"Total: {total_count}\nMobile: {mobile_count}")

    def refresh_callsign_combo(self):
        current = self.combo_traffic_call.currentText()
        self.combo_traffic_call.clear()
        self.combo_traffic_call.addItem("")
        for st in self.stations:
            if st["status"] != "Checked Out":
                self.combo_traffic_call.addItem(st["callsign"])
        self.combo_traffic_call.setCurrentText(current)

    def toggle_station_standby(self):
        selected = self.station_table.currentRow()
        if selected < 0:
            return
        st = self.stations[selected]
        if st["status"] == "Active":
            st["status"] = "Standby"
        elif st["status"] == "Standby":
            st["status"] = "Active"
        self.refresh_station_table()
        self.station_table.selectRow(selected)

    def checkout_station(self):
        selected = self.station_table.currentRow()
        if selected < 0:
            return
        st = self.stations[selected]
        now_str = datetime.now().strftime("%H:%M")
        st["status"] = "Checked Out"
        st["time_out"] = now_str
        self.refresh_station_table()
        self.refresh_callsign_combo()
        self.station_table.selectRow(selected)

    def log_traffic_item(self):
        call = self.combo_traffic_call.currentText().strip().upper() or "NET"
        traffic_type = self.combo_traffic_type.currentText()
        note = self.input_traffic_note.toPlainText().strip()
        time_val = self.input_traffic_time.text().strip() or datetime.now().strftime("%H:%M")
        loc_val = self.input_traffic_loc.text().strip() or "-"

        if not note:
            QMessageBox.warning(self, "Input Error", "Please provide details for the entry.")
            return

        self.add_activity_log(call, traffic_type, loc_val, note, time_val)
        self.input_traffic_note.clear()
        self.input_traffic_loc.clear()

    def add_activity_log(self, call, traffic_type, location, note, custom_time=None):
        time_str = custom_time if custom_time else datetime.now().strftime("%H:%M")

        row = self.activity_table.rowCount()
        self.activity_table.insertRow(row)

        if self.current_mode == "SKYWARN":
            activity_item = {
                "timestamp": time_str,
                "condition": traffic_type,
                "location": location,
                "source": call,
                "narrative": note
            }
            self.activities.append(activity_item)

            self.activity_table.setItem(row, 0, QTableWidgetItem(time_str))
            self.activity_table.setItem(row, 1, QTableWidgetItem(traffic_type))
            self.activity_table.setItem(row, 2, QTableWidgetItem(location))
            self.activity_table.setItem(row, 3, QTableWidgetItem(call))
            self.activity_table.setItem(row, 4, QTableWidgetItem(note))
        else:
            activity_item = {
                "timestamp": time_str,
                "callsign": call,
                "type": traffic_type,
                "note": note
            }
            self.activities.append(activity_item)

            self.activity_table.setItem(row, 0, QTableWidgetItem(time_str))
            self.activity_table.setItem(row, 1, QTableWidgetItem(call))

            type_item = QTableWidgetItem(traffic_type)
            if traffic_type == "Emergency":
                type_item.setForeground(QColor("#FF0000"))
            elif traffic_type == "Priority":
                type_item.setForeground(QColor("#804000"))

            self.activity_table.setItem(row, 2, type_item)
            self.activity_table.setItem(row, 3, QTableWidgetItem(note))

        self.activity_table.scrollToBottom()

    def export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export CSV Log", "net_log.csv", "CSV Files (*.csv)")
        if not path:
            return

        try:
            with open(path, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["=== STATION ROSTER ==="])
                if self.current_mode == "SKYWARN":
                    writer.writerow(["Call", "Station Type", "Location", "Status", "In", "Out"])
                    for st in self.stations:
                        writer.writerow([st.get("callsign", ""), st.get("type", ""), st.get("detail", "-"), st.get("status", ""), st.get("time_in", ""), st.get("time_out", "")])
                else:
                    writer.writerow(["Call", "Station Type", "Additional", "Info / Grid", "Status", "In", "Out"])
                    for st in self.stations:
                        writer.writerow([st.get("callsign", ""), st.get("type", ""), st.get("detail", "-"), st.get("info", "-"), st.get("status", ""), st.get("time_in", ""), st.get("time_out", "")])

                writer.writerow([])
                writer.writerow(["=== NET ACTIVITY & TRAFFIC LOG ==="])
                if self.current_mode == "SKYWARN":
                    writer.writerow(["Time", "Condition", "Location", "Source", "Narrative"])
                    for act in self.activities:
                        writer.writerow([act.get("timestamp", ""), act.get("condition", ""), act.get("location", ""), act.get("source", ""), act.get("narrative", "")])
                else:
                    writer.writerow(["Time", "Callsign", "Traffic Tag", "Details"])
                    for act in self.activities:
                        writer.writerow([act.get("timestamp", ""), act.get("callsign", ""), act.get("type", ""), act.get("note", "")])

            QMessageBox.information(self, "Export Successful", f"Net log successfully exported to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export CSV file:\n{str(e)}")

    def export_pdf(self):
        path, _ = QFileDialog.getSaveFileName(self, "Export PDF Summary", "net_report.pdf", "PDF Files (*.pdf)")
        if not path:
            return

        date_str = datetime.now().strftime("%Y-%m-%d %H:%M")

        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Tahoma, sans-serif; font-size: 11pt; color: #000; }}
                h1 {{ color: #000080; margin-bottom: 2px; }}
                .meta {{ font-size: 9pt; color: #555; margin-bottom: 20px; }}
                table {{ width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 20px; }}
                th {{ background-color: #000080; color: white; text-align: left; padding: 6px; font-size: 10pt; }}
                td {{ border: 1px solid #808080; padding: 6px; font-size: 9.5pt; }}
                tr.row-standby {{ background-color: #FFFF99; color: #000; }}
                tr.row-checkout {{ background-color: #CC0000; color: #FFF; }}
            </style>
        </head>
        <body>
            <h1>Amateur Radio Net Control Summary</h1>
            <div class="meta">Generated: {date_str} | Profile: {self.current_profile_name} ({self.current_mode} Mode)</div>

            <h3>Station Roster ({len(self.stations)} Stations)</h3>
            <table>
                <tr>
                    <th>Call</th>
                    <th>Station Type</th>
                    <th>{'Location' if self.current_mode == 'SKYWARN' else 'Additional'}</th>
                    {'' if self.current_mode == 'SKYWARN' else '<th>Info / Grid</th>'}
                    <th>Status</th>
                    <th>In</th>
                    <th>Out</th>
                </tr>
        """

        for st in self.stations:
            row_class = ""
            if st.get("status") == "Standby":
                row_class = ' class="row-standby"'
            elif st.get("status") == "Checked Out":
                row_class = ' class="row-checkout"'

            info_td = f"<td>{st.get('info', '-')}</td>" if self.current_mode != 'SKYWARN' else ""

            html += f"""
                <tr{row_class}>
                    <td><b>{st.get('callsign', '')}</b></td>
                    <td>{st.get('type', '')}</td>
                    <td>{st.get('detail', '-')}</td>
                    {info_td}
                    <td>{st.get('status', '')}</td>
                    <td>{st.get('time_in', '')}</td>
                    <td>{st.get('time_out', '')}</td>
                </tr>
            """

        html += f"""
            </table>

            <h3>Net Activity & Traffic Log</h3>
            <table>
                <tr>
                    {"<th>Time</th><th>Condition</th><th>Location</th><th>Source</th><th>Narrative</th>" if self.current_mode == "SKYWARN" else "<th>Time</th><th>Callsign</th><th>Traffic Tag</th><th>Details</th>"}
                </tr>
        """

        for act in self.activities:
            if self.current_mode == "SKYWARN":
                html += f"""
                    <tr>
                        <td>{act.get('timestamp', '')}</td>
                        <td><b>{act.get('condition', '')}</b></td>
                        <td>{act.get('location', '-')}</td>
                        <td>{act.get('source', '')}</td>
                        <td>{act.get('narrative', '')}</td>
                    </tr>
                """
            else:
                html += f"""
                    <tr>
                        <td>{act.get('timestamp', '')}</td>
                        <td><b>{act.get('callsign', '')}</b></td>
                        <td>{act.get('type', '')}</td>
                        <td>{act.get('note', '')}</td>
                    </tr>
                """

        html += """
            </table>
        </body>
        </html>
        """

        try:
            document = QTextDocument()
            document.setHtml(html)

            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(path)

            document.print(printer)
            QMessageBox.information(self, "Export Successful", f"PDF report successfully saved to:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export PDF file:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = NetControlApp()
    window.show()

    QTimer.singleShot(50, window.prompt_profile_on_startup)

    sys.exit(app.exec())
