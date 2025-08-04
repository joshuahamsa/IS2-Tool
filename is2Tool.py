import sys
import zipfile
import shutil
import openpyxl
from pathlib import Path
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QVBoxLayout, QHBoxLayout,
    QFileDialog, QComboBox, QCheckBox, QMessageBox, QGroupBox, QGridLayout, 
    QDialog, QLineEdit, QScrollArea, QMainWindow, QAction, QStackedLayout,
    QSpacerItem, QSizePolicy
)
from PyQt5.QtWidgets import QDateEdit
from PyQt5.QtCore import QDate
from datetime import datetime
import pywintypes
import win32file
import win32con
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))


def convert_and_unzip(is2_filepath):
    temp_zip = is2_filepath.with_suffix('.zip')
    shutil.copyfile(is2_filepath, temp_zip)
    extract_dir = temp_zip.parent / f"{temp_zip.stem}_unzipped"
    with zipfile.ZipFile(temp_zip, 'r') as zf:
        zf.extractall(extract_dir)
    temp_zip.unlink()
    return extract_dir  # ✅ Proper return


def set_file_created_to_modified(path: Path):
# v1.4 - add option to update the Date Created of the is2 file to match the Date Modified, which is the date the photo was taken
    """
    Sets a file's 'Date Created' to match its 'Date Modified'.
    Only works on Windows NTFS.
    """
    if not path.exists():
        return

    handle = win32file.CreateFileW(
        str(path),
        win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None
    )

    mod_time = pywintypes.Time(path.stat().st_mtime)

    win32file.SetFileTime(handle, mod_time, mod_time, mod_time)
    handle.close()

def set_windows_creation_time(target_file: Path, dt: datetime):
    # modified for v1.6 to use pywin32 instead of powershell for faster processing
    wintime = pywintypes.Time(dt)
    handle = win32file.CreateFile(
        str(target_file),
        win32con.GENERIC_WRITE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
        None,
        win32con.OPEN_EXISTING,
        win32con.FILE_ATTRIBUTE_NORMAL,
        None
    )
    win32file.SetFileTime(handle, wintime, None, None)  # Created, Accessed, Modified
    handle.close()


def get_visible_thumbnail(images_main_dir):
    jpgs = list(images_main_dir.glob('*.jpg'))
    return min(jpgs, key=lambda x: x.stat().st_size) if jpgs else None

def get_ir_thumbnail(thumbnails_dir):
    jpgs = list(thumbnails_dir.glob('*.jpg'))
    return min(jpgs, key=lambda x: x.stat().st_size) if jpgs else None

def get_photonotes_thumbnails(photo_notes_dir):
    thumbnails = []
    for i in range(3):
        sub = photo_notes_dir / str(i)
        if sub.is_dir():
            jpgs = list(sub.glob('*.jpg'))
            if jpgs:
                sorted_files = sorted(jpgs, key=lambda x: x.stat().st_size)
                thumb = sorted_files[0]
                full = sorted_files[-1] if len(sorted_files) > 1 else sorted_files[0]
                thumbnails.append((thumb, full))
    return thumbnails


class HomeScreen(QWidget):
    def __init__(self, on_start_callback):
        super().__init__()

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)

        title = QLabel("Welcome to IS2 Tool")
        title.setStyleSheet("font-size: 24px; font-weight: bold;")
        title.setAlignment(Qt.AlignHCenter)

        subtitle = QLabel("Process .is2 files from Fluke IR cameras.\nExport visible images and update metadata easily. \n\nClick Get Started to select a folder of .is2 files.")
        subtitle.setStyleSheet("font-size: 14px; padding: 10px;")
        subtitle.setAlignment(Qt.AlignHCenter)

        start_button = QPushButton("Get Started")
        start_button.setFixedWidth(200)
        start_button.setStyleSheet("background-color: #FFCC00; font-weight: bold;")
        start_button.clicked.connect(on_start_callback)
        start_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # Prevent weird stretching
        # Wrap in a horizontal box to center
        btn_container = QHBoxLayout()
        btn_container.addStretch()
        btn_container.addWidget(start_button)
        btn_container.addStretch()

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addLayout(btn_container)

        # Add vertical spacers to center vertically
        layout.insertStretch(0, 1)
        layout.addStretch(1)

        self.setLayout(layout)


class ZoomWindow(QDialog):
    def __init__(self, full_image_path):
        super().__init__()
        self.setWindowTitle("Photo - Full Resolution")

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setPixmap(QPixmap(str(full_image_path)))
        self.image_label.setScaledContents(True)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.image_label)

        layout = QVBoxLayout()
        layout.addWidget(self.scroll_area)
        self.setLayout(layout)

        self.scale_factor = 1.0
        self.original_pixmap = QPixmap(str(full_image_path))
        self._setup_events()
        self.reset_zoom(fit_to_window=True)

    def _setup_events(self):
        self.image_label.installEventFilter(self)
        self.scroll_area.viewport().installEventFilter(self)

        self.scroll_area.setFocusPolicy(Qt.StrongFocus)
        self.scroll_area.viewport().setFocusPolicy(Qt.StrongFocus)

        self.setMinimumSize(800, 600)
        self.image_label.setCursor(Qt.OpenHandCursor)
        self.drag_start_position = None
        self.dragging = False

    def eventFilter(self, source, event):
        if event.type() == event.MouseButtonDblClick:
            if abs(self.scale_factor - 1.0) > 0.01:
                self.reset_zoom()
            else:
                self.zoom(1.5)
            return True

        if event.type() == event.Wheel:
            if source is self.scroll_area.viewport() or source is self.image_label:
                delta = event.angleDelta().y()
                if delta > 0:
                    self.zoom(1.25)
                else:
                    self.zoom(0.8)
                return True

        if event.type() == event.MouseButtonPress:
            if event.button() == Qt.LeftButton:
                self.drag_start_position = event.globalPos()
                self.dragging = True
                self.image_label.setCursor(Qt.ClosedHandCursor)
                return True

        if event.type() == event.MouseMove and self.dragging:
            delta = event.globalPos() - self.drag_start_position
            self.scroll_area.horizontalScrollBar().setValue(self.scroll_area.horizontalScrollBar().value() - delta.x())
            self.scroll_area.verticalScrollBar().setValue(self.scroll_area.verticalScrollBar().value() - delta.y())
            self.drag_start_position = event.globalPos()
            return True

        if event.type() == event.MouseButtonRelease:
            self.dragging = False
            self.image_label.setCursor(Qt.OpenHandCursor)
            return True

        return super().eventFilter(source, event)

    def zoom(self, factor):
        new_scale = self.scale_factor * factor
        # Prevent zooming out too far or going crazy high
        if new_scale < 0.1 or new_scale > 20.0:
            return
        self.scale_factor = new_scale
        scaled_pixmap = self.original_pixmap.scaled(
            self.original_pixmap.size() * self.scale_factor,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )
        self.image_label.setPixmap(scaled_pixmap)
        self.image_label.resize(scaled_pixmap.size())

    def reset_zoom(self, fit_to_window=False):
        if fit_to_window:
            container_width = self.scroll_area.viewport().width()
            container_height = self.scroll_area.viewport().height()
            image_size = self.original_pixmap.size()

            scale_w = container_width / image_size.width()
            scale_h = container_height / image_size.height()
            self.scale_factor = min(scale_w, scale_h, 1.0)  # Never upscale by default
        else:
            self.scale_factor = 1.0

        scaled_pixmap = self.original_pixmap.scaled(
            self.original_pixmap.size() * self.scale_factor,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )
        self.image_label.setPixmap(scaled_pixmap)
        self.image_label.resize(scaled_pixmap.size())


class ImageReviewApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IS2 Tool")
        self.resize(800, 600)
        frame_geom = self.frameGeometry()
        screen_center = QApplication.primaryScreen().availableGeometry().center()
        frame_geom.moveCenter(screen_center)
        self.move(frame_geom.topLeft())
        self.setMinimumSize(600, 400)

        # --- Central widget and stacked layout ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.stacked_layout = QStackedLayout()
        self.central_widget.setLayout(self.stacked_layout)

        # --- Home screen ---
        self.home_screen = HomeScreen(self.show_main_tool)
        self.stacked_layout.addWidget(self.home_screen)

        # --- Main tool screen ---
        self.main_tool_widget = QWidget()
        self.tool_layout = QVBoxLayout()
        self.main_tool_widget.setLayout(self.tool_layout)
        self.stacked_layout.addWidget(self.main_tool_widget)

        # --- Build main tool UI using tool_layout ---
        self.filename_label = QLabel("")
        self.filename_label.setAlignment(Qt.AlignCenter)
        self.filename_label.setMaximumHeight(20)
        self.filename_label.setMinimumHeight(20)
        self.filename_label.setContentsMargins(0, 0, 0, 0)
        self.tool_layout.addWidget(self.filename_label)

        self.filter_checkbox = QCheckBox("Only show unrenamed files (start with 'IR_')")
        self.filter_checkbox.setChecked(False)

        self.export_visible_checkbox = QCheckBox("Export Visible Image")
        self.export_visible_checkbox.setChecked(True)

        self.images_group = QGroupBox("Thumbnails")
        self.images_layout = QGridLayout()
        self.images_group.setLayout(self.images_layout)
        self.tool_layout.addWidget(self.images_group)
        self.images_group.setVisible(False)

        self.label_ir = QLabel("IR Thumbnail")
        self.label_visible = QLabel("Visible Thumbnail")
        for label in [self.label_ir, self.label_visible]:
            label.setAlignment(Qt.AlignCenter)
        self.images_layout.addWidget(self.label_ir, 0, 0)
        self.images_layout.addWidget(self.label_visible, 0, 1)

        self.note_labels = []
        for i in range(3):
            label = QLabel(f"PhotoNote {i}")
            label.setAlignment(Qt.AlignCenter)
            self.images_layout.addWidget(label, 1, i)
            self.note_labels.append(label)

        self.tier_layout = QHBoxLayout()
        self.tier_combos = []
        self.tool_layout.addLayout(self.tier_layout)

        self.filename_field = QLineEdit()
        self.filename_field.setPlaceholderText("Enter custom filename or suffix")
        self.tool_layout.addWidget(self.filename_field)
        self.filename_field.setVisible(False)

        self.save_next_button = QPushButton("Save && Next")
        self.save_next_button.clicked.connect(self.save_and_next)
        self.tool_layout.addWidget(self.save_next_button)
        self.save_next_button.setVisible(False)

        self.nav_layout = QHBoxLayout()
        self.back_button = QPushButton("<")
        self.back_button.clicked.connect(self.go_back)
        self.nav_layout.addWidget(self.back_button)
        self.back_button.setVisible(False)

        self.next_button = QPushButton(">")
        self.next_button.clicked.connect(self.go_next)
        self.nav_layout.addWidget(self.next_button)
        self.next_button.setVisible(False)

        self.tool_layout.addLayout(self.nav_layout)

        # Menus and state
        self.create_menu_bar()
        self.is2_files = []
        self.current_index = 0
        self.used_names = {}
        self.extract_dir = None
        self.extract_dirs = []
        self.exported_images = {}  # v1.5 dictionary to track exported images

    def show_main_tool(self):
        self.stacked_layout.setCurrentWidget(self.main_tool_widget)
        self.select_folder()  # Prompt user to choose folder immediately

    def make_mouse_handler(self, full_path):
        return lambda e: self.handle_photonote_click(full_path)

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if not folder_path:
            return
        folder = Path(folder_path)
        all_files = [f for f in folder.iterdir() if f.suffix.lower() == ".is2"]
        if self.filter_checkbox.isChecked():
            all_files = [f for f in all_files if f.name.startswith("IR_")]
        # Auto-run date correction if all files start with "IR_"
        if self.is2_files and all(f.name.startswith("IR_") for f in self.is2_files):
            updated = 0
            for f in self.is2_files:
                try:
                    set_file_created_to_modified(f)
                    updated += 1
                except Exception as e:
                    print(f"[Auto] Failed to update {f.name}: {e}")
            print(f"[Auto] Updated Date Created on {updated} files.")
        self.is2_files = sorted(all_files, key=lambda x: x.stat().st_mtime) 
        self.current_index = 0
        self.show_current_file()

    def import_locations(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Locations Excel File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active

        self.tier_tree = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            non_empty = [str(cell).strip() for cell in row if cell and str(cell).strip()]
            if not non_empty:
                continue
            current = self.tier_tree
            for level in non_empty:
                current = current.setdefault(level, {})

        self.build_dynamic_tiers()
        QMessageBox.information(self, "Imported", "Locations file loaded successfully.")

    def build_dynamic_tiers(self):
        for combo in self.tier_combos:
            self.tier_layout.removeWidget(combo)
            combo.deleteLater()
        self.tier_combos.clear()

        def get_depth(tree, depth=0):
            return depth if not tree else max(get_depth(sub, depth + 1) for sub in tree.values())

        for i in range(get_depth(self.tier_tree)):
            combo = QComboBox()
            combo.setObjectName(f"Tier{i}")
            combo.currentIndexChanged.connect(lambda _, idx=i: self.update_dependent_combos(idx))
            self.tier_layout.addWidget(combo)
            self.tier_combos.append(combo)

        self.tier_combos[0].blockSignals(True)
        self.tier_combos[0].clear()
        self.tier_combos[0].addItems(sorted(self.tier_tree.keys()))
        self.tier_combos[0].setCurrentIndex(-1)
        self.tier_combos[0].blockSignals(False)
        self.update_dependent_combos(0)

    def update_dependent_combos(self, changed_index):
        current_tree = self.tier_tree
        for i in range(changed_index + 1):
            selected = self.tier_combos[i].currentText()
            if selected in current_tree:
                current_tree = current_tree[selected]
            else:
                return

        for j in range(changed_index + 1, len(self.tier_combos)):
            self.tier_combos[j].blockSignals(True)
            self.tier_combos[j].clear()
            self.tier_combos[j].blockSignals(False)

        if changed_index + 1 < len(self.tier_combos):
            next_combo = self.tier_combos[changed_index + 1]
            next_combo.blockSignals(True)
            next_combo.clear()
            next_combo.addItems(sorted(current_tree.keys()))
            next_combo.setCurrentIndex(-1)
            next_combo.blockSignals(False)

    def show_current_file(self):
        if self.extract_dir and self.extract_dir.exists():
            shutil.rmtree(self.extract_dir, ignore_errors=True)

        if self.current_index >= len(self.is2_files):
            QMessageBox.information(self, "Done", "No more files to process.")
            return

        # Make UI options visible after loading a folder
        self.images_group.setVisible(True)
        self.filename_field.setVisible(True)
        self.back_button.setVisible(True)
        self.next_button.setVisible(True)
        self.save_next_button.setVisible(True)

        self.label_ir.clear()
        self.label_visible.clear()
        for lab in self.note_labels:
            lab.clear()
            lab.setText("")

        is2_file = self.is2_files[self.current_index]
        self.filename_label.setText(f"Current File Name: {is2_file.name}")
        self.extract_dir = convert_and_unzip(is2_file)
        # Automatically update Date Created if file starts with "IR_"
        if is2_file.name.startswith("IR_"):
            try:
                set_file_created_to_modified(is2_file)
            except Exception as e:
                print(f"Failed to update created date for {is2_file.name}: {e}")

        self.extract_dirs.append(self.extract_dir)

        ir_thumb = get_ir_thumbnail(self.extract_dir / "Thumbnails")
        if ir_thumb:
            self.label_ir.setPixmap(QPixmap(str(ir_thumb)).scaled(250, 250, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self.label_ir.mousePressEvent = self.make_mouse_handler(ir_thumb)
            self.label_ir.setCursor(Qt.PointingHandCursor)
        else:
            self.label_ir.setText("No IR thumbnail found")
            self.label_ir.mousePressEvent = None

        visible_thumb = get_visible_thumbnail(self.extract_dir / "Images" / "Main")
        if visible_thumb:
            self.label_visible.setPixmap(QPixmap(str(visible_thumb)).scaled(250, 250, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            self.label_visible.mousePressEvent = self.make_mouse_handler(visible_thumb)
            self.label_visible.setCursor(Qt.PointingHandCursor)

        else:
            self.label_visible.setText("No visible thumbnail found")
            self.label_ir.mousePressEvent = None

        photo_notes_dir = self.extract_dir / "PhotoNotes"
        if photo_notes_dir.is_dir():
            for i, (thumb_file, full_file) in enumerate(get_photonotes_thumbnails(photo_notes_dir)):
                if i < 3:
                    pixmap = QPixmap(str(thumb_file))
                    self.note_labels[i].setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                    self.note_labels[i].mousePressEvent = self.make_mouse_handler(full_file)
                    self.note_labels[i].setCursor(Qt.PointingHandCursor)

    def save_and_next(self):
        if self.current_index >= len(self.is2_files):
            QMessageBox.information(self, "Done", "All files processed.")
            return

        location_parts = [c.currentText().strip()
                        for c in self.tier_combos if c.currentText().strip()]
        custom_input = self.filename_field.text().strip()

        if not location_parts and not custom_input:
            QMessageBox.warning(self, "Missing Input",
                                "Please select a location or enter a custom filename.")
            return

        base_name = " ".join(location_parts)
        if custom_input:
            base_name = f"{base_name} - {custom_input}" if base_name else custom_input

        original_file = self.is2_files[self.current_index]

        target_dir  = original_file.parent
        candidate   = target_dir / f"{base_name}.is2"
        suffix_num  = 2
        while candidate.exists():
            candidate = target_dir / f"{base_name}({suffix_num}).is2"
            suffix_num += 1
        new_file = candidate

        try:
            original_file.rename(new_file)

            # Export visible image if requested
            if self.export_visible_checkbox.isChecked():
                visible_dir = self.extract_dir / "Images" / "Main"
                visible_image = get_visible_thumbnail(visible_dir)
                if visible_image:
                    export_path = new_file.with_suffix(".jpg")
                    shutil.copyfile(visible_image, export_path)
                    self.exported_images[export_path] = datetime.fromtimestamp(new_file.stat().st_mtime)
                else:
                    QMessageBox.warning(self, "Export Failed", "No visible image found to export.")

        except Exception as e:
            QMessageBox.critical(self, "Rename Error", f"Failed to rename file:\n{e}")
            return

        if self.extract_dir:
            shutil.rmtree(self.extract_dir, ignore_errors=True)

        self.is2_files[self.current_index] = new_file
        self.current_index += 1

        if self.current_index < len(self.is2_files):
            self.show_current_file()
        else:
            QMessageBox.information(self, "Done", "All files processed.")

    def refresh_is2_list(self):
        if self.is2_files:
            folder = self.is2_files[0].parent
            all_files = [f for f in folder.iterdir() if f.suffix.lower() == ".is2"]
            if self.filter_checkbox.isChecked():
                all_files = [f for f in all_files if f.name.startswith("IR_")]
            self.is2_files = sorted(all_files, key=lambda x: x.name.lower())

    def go_next(self):
        self.refresh_is2_list()
        if self.current_index < len(self.is2_files) - 1:
            self.current_index += 1
            self.show_current_file()
        else:
            QMessageBox.information(self, "End", "This is the last file.")

    def go_back(self):
        self.refresh_is2_list()
        if self.current_index > 0:
            self.current_index -= 1
            self.show_current_file()
        else:
            QMessageBox.information(self, "Start", "This is the first file.")

    def handle_photonote_click(self, full_image_path):
        ZoomWindow(full_image_path).exec_()

    def closeEvent(self, event):
        # Clean up unzipped folders
        for d in getattr(self, 'extract_dirs', []):
            if d.exists():
                try:
                    shutil.rmtree(d, ignore_errors=True)
                except Exception:
                    pass  # Silently ignore cleanup errors for production use

        # v1.5 Only update creation dates if we exported any visible light images
        if getattr(self, 'exported_images', None):
            for jpg_path, desired_dt in self.exported_images.items():
                try:
                    set_windows_creation_time(jpg_path, desired_dt)
                except Exception:
                    pass  # Silently fail to avoid app crash on close

        event.accept()


    def set_created_dates_for_all(self):
        if not self.is2_files:
            QMessageBox.information(self, "No Files", "No .is2 files loaded. Select a folder first.")
            return

        updated = 0
        for f in self.is2_files:
            try:
                set_file_created_to_modified(f)
                updated += 1
            except Exception as e:
                print(f"Failed to update {f.name}: {e}")

        QMessageBox.information(self, "Done", f"Updated Date Created on {updated} .is2 files.")

    def create_menu_bar(self):
        menubar = self.menuBar()

        # --- File Menu ---
        file_menu = menubar.addMenu("File")

        open_folder_action = QAction("Open Folder...", self)
        open_folder_action.triggered.connect(self.select_folder)
        file_menu.addAction(open_folder_action)

        import_locations_action = QAction("Import Locations File", self)
        import_locations_action.triggered.connect(self.import_locations)
        file_menu.addAction(import_locations_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # --- Tools Menu ---
        tools_menu = menubar.addMenu("Tools")

        # Action: Set .is2 File Dates
        set_dates_action = QAction("Set Created Dates on .is2 Files", self)
        set_dates_action.triggered.connect(self.set_created_dates_for_all)
        tools_menu.addAction(set_dates_action)

        # --- Options Menu ---
        tools_menu = menubar.addMenu("Options")
        # Checkable: Only show unrenamed files
        self.filter_checkbox_action = QAction("Only show unrenamed files", self, checkable=True)
        self.filter_checkbox_action.setChecked(self.filter_checkbox.isChecked())
        self.filter_checkbox_action.toggled.connect(self.filter_checkbox.setChecked)
        self.filter_checkbox.toggled.connect(self.filter_checkbox_action.setChecked)
        tools_menu.addAction(self.filter_checkbox_action)

        # Checkable: Export visible image
        self.export_visible_action = QAction("Export Visible Image", self, checkable=True)
        self.export_visible_action.setChecked(self.export_visible_checkbox.isChecked())
        self.export_visible_action.toggled.connect(self.export_visible_checkbox.setChecked)
        self.export_visible_checkbox.toggled.connect(self.export_visible_action.setChecked)
        tools_menu.addAction(self.export_visible_action)

        # --- Help Menu ---
        help_menu = menubar.addMenu("Help")

        about_action = QAction("About", self)
        about_action.triggered.connect(lambda: QMessageBox.about(
            self, "About IS2 Tool",
            "IS2 Tool v2\nDeveloped with ❤️ by Joshua Hamsa\n\nProcesses .is2 files from Fluke IR cameras and exports visible images."
        ))
        help_menu.addAction(about_action)


def main():
    app = QApplication(sys.argv)

    # Load QSS
    with open("theme.qss", "r") as f:
        app.setStyleSheet(f.read())

    window = ImageReviewApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
