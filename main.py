import sys
import os
import fitz  # PyMuPDF
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QListWidget, QListWidgetItem, 
                               QFileDialog, QLabel, QMessageBox, QSplitter, QFrame,
                               QSlider, QSpinBox, QGroupBox, QAbstractItemView,
                               QMenu, QInputDialog, QLineEdit, QComboBox, QProgressBar,
                               QCheckBox)
from PySide6.QtCore import Qt, QSize, QThread, Signal, QMimeData
from PySide6.QtGui import QIcon, QPixmap, QImage, QAction, QFont, QDrag

# --- STYLING ---
DARK_THEME_QSS = """
QMainWindow {
    background-color: #1e1e1e;
    color: #ffffff;
}
QWidget {
    background-color: #1e1e1e;
    color: #ffffff;
    font-family: 'Segoe UI', 'Microsoft JhengHei';
    font-size: 14px;
}
QListWidget {
    background-color: #252526;
    border: 1px solid #3e3e42;
    border-radius: 5px;
    padding: 10px;
}
QListWidget::item {
    background-color: #2d2d30;
    margin: 5px;
    border-radius: 5px;
    color: #cccccc;
}
QListWidget::item:selected {
    background-color: #007acc;
    color: white;
}
QGroupBox {
    border: 1px solid #3e3e42;
    border-radius: 5px;
    margin-top: 20px;
    font-weight: bold;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 3px 0 3px;
}
QPushButton {
    background-color: #0e639c;
    color: white;
    border: none;
    padding: 8px 15px;
    border-radius: 4px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #1177bb;
}
QPushButton:pressed {
    background-color: #094770;
}
QPushButton:disabled {
    background-color: #3e3e42;
    color: #888888;
}
QLabel#StatusLabel {
    color: #aaaaaa;
    font-size: 12px;
}
QLabel#SectionHeader {
    color: #007acc;
    font-weight: bold;
    padding: 5px;
}
"""

class PDFWorker(QThread):
    """Background thread for heavy PDF operations"""
    finished = Signal(object)
    
    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit(e)

class SaveWorker(QThread):
    finished = Signal(bool, str) # Success, Message
    progress = Signal(int, int) # Current, Total

    def __init__(self, items_data, source_docs, out_path, overlays):
        super().__init__()
        self.items_data = items_data
        self.source_docs = source_docs
        self.out_path = out_path
        self.overlays = overlays # Dict: text, pos, color, size
        self.running = True

    def run(self):
        try:
            doc = fitz.open()
            total = len(self.items_data)
            
            for i, item_data in enumerate(self.items_data):
                if not self.running: return

                doc_id = item_data['doc_id']
                page_num = item_data['page_num']
                rotation = item_data['rotation']
                
                # Find source doc
                src_doc = None
                for entry in self.source_docs:
                    if entry['id'] == doc_id:
                        src_doc = entry['doc']
                        break
                
                if src_doc:
                    doc.insert_pdf(src_doc, from_page=page_num, to_page=page_num)
                    page = doc[-1]
                    
                    # Apply Rotation FIRST
                    if rotation != 0:
                        page.set_rotation((page.rotation + rotation) % 360)
                    
                    # Apply Overlay
                    # Pass the page name (from items_data)
                    page_name = item_data.get('text', '')
                    self._apply_overlay(page, i + 1, total, page_name)

                self.progress.emit(i + 1, total)
            
            # Save
            doc.save(self.out_path, garbage=4, deflate=True)
            doc.close()
            self.finished.emit(True, f"檔案已成功儲存至:\n{self.out_path}")
            
        except Exception as e:
            self.finished.emit(False, str(e))

    def _apply_overlay(self, page, current_num, total_pages, page_name):
        # 1. Check if enabled
        if not self.overlays.get('enabled', False):
            return

        text_templ = self.overlays.get('text', '')
        if not text_templ: return
        
        # Replace placeholders
        text = text_templ.replace('{n}', str(current_num))\
                         .replace('{total}', str(total_pages))\
                         .replace('{name}', str(page_name))
        
        pos = self.overlays.get('pos', 'Bottom-Right')
        color_name = self.overlays.get('color', 'Black')
        size = self.overlays.get('size', 12)
        
        # Color Map
        colors = {
            'Black': (0, 0, 0),
            'White': (1, 1, 1),
            'Red': (1, 0, 0),
            'Blue': (0, 0, 1),
            'Gray': (0.5, 0.5, 0.5)
        }
        rgb = colors.get(color_name, (0, 0, 0))

        # --- Use Unrotated Coordinates + Derotation ---
        rect = page.rect
        w = rect.width
        h = rect.height
        
        margin = 20
        vx, vy = 0, 0 # Visual coordinates
        align = 0 # 0=left, 1=center, 2=right
        
        # 2. Calculate Visual Position (vx, vy)
        # vy calculation (Vertical)
        if 'Top' in pos:
            vy = margin + size # Approx baseline
        elif 'Bottom' in pos:
            vy = h - margin
        else: # Middle
            vy = (h / 2) + (size * 0.35) # Approx vertical center adjustment
            
        # vx calculation (Horizontal)
        if 'Left' in pos:
            vx = margin
            align = 0
        elif 'Right' in pos:
            vx = w - margin
            align = 2
        elif 'Center' in pos:
            vx = w / 2
            align = 1

        # 3. Adjust vx for Alignment (Text Width)
        # Try finding a suitable Chinese font for calculation
        font_paths = [
            "C:/Windows/Fonts/msjh.ttc", # Microsoft JhengHei
            "C:/Windows/Fonts/msyh.ttc", # Microsoft YaHei
            "C:/Windows/Fonts/simsun.ttc", # SimSun
            "C:/Windows/Fonts/arial.ttf" # Fallback
        ]
        
        font = None
        font_file_used = None
        
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    font = fitz.Font(fontfile=fp)
                    font_file_used = fp
                    break
                except:
                    continue
        
        if font:
            width = font.text_length(text, fontsize=size)
        else:
            # Better Fallback Estimation: Chinese ~ size, ASCII ~ 0.5*size
            width = sum(size if ord(c) > 255 else size * 0.5 for c in text)

        if align == 2: # Right aligned
            vx -= width
        elif align == 1: # Center aligned
            vx -= (width / 2)
            
        # 4. Transform Visual Point (vx, vy) -> Physical Point (px, py)
        # Use derotation_matrix to map Visual -> Physical
        p_vis = fitz.Point(vx, vy)
        mat = page.derotation_matrix
        p_phys = p_vis * mat
        
        # 5. Calculate Text Rotation
        # Text needs to rotate WITH page rotation logic because insert_text is CCW and Page is CW
        # Logic: We want Visual Right -> Physical Transformed.
        # Analysis shows text_rot = page.rotation provides correct orientation (0->0, 90->90, 180->180, 270->270)
        text_rot = page.rotation
        
        try:
            if font_file_used:
                 page.insert_text(p_phys, text, fontsize=size, color=rgb, rotate=text_rot, fontfile=font_file_used)
            else:
                 page.insert_text(p_phys, text, fontsize=size, color=rgb, rotate=text_rot, fontname="china-ts")
        except Exception as e:
            print(f"Overlay Error: {e}")


class ThumbnailCache:
    """Cache for PDF page thumbnails to avoid reloading from disk constantly."""
    def __init__(self):
        self._cache = {} # Key: (doc_id, page_num), Value: QImage (base, 0 rotation)

    def get_image(self, doc_id, page_num):
        return self._cache.get((doc_id, page_num))

    def set_image(self, doc_id, page_num, image):
        self._cache[(doc_id, page_num)] = image
        
    def clear(self):
        self._cache.clear()

class HistoryManager:
    """Manages Undo & Redo History."""
    def __init__(self, max_stack=20):
        self.undo_stack = []
        self.redo_stack = []
        self.max_stack = max_stack
    
    def push_state(self, state):
        """Pushes a new state to undo stack and clears redo stack."""
        # If new state is same as last, ignore to prevent duplicate states
        if self.undo_stack and self.undo_stack[-1] == state:
            return

        self.undo_stack.append(state)
        self.redo_stack.clear() # Clear redo when a new action occurs
        
        if len(self.undo_stack) > self.max_stack:
            self.undo_stack.pop(0)

    def push_to_redo(self, state):
        """Pushes state to redo stack (used during Undo)."""
        self.redo_stack.append(state)
        
    def pop_undo(self):
        if not self.undo_stack:
            return None
        return self.undo_stack.pop()

    def pop_redo(self):
        if not self.redo_stack:
            return None
        return self.redo_stack.pop()
    
    def can_undo(self):
        return len(self.undo_stack) > 0
        
    def can_redo(self):
        return len(self.redo_stack) > 0


class PDFPageList(QListWidget):
    """Custom ListWidget to handle Drag & Drop of PDF Pages"""
    filesDropped = Signal(list) # Emitted when actual files are dropped
    aboutToChange = Signal()
    # contextMenuRequested = Signal(object) # Removed redundant signal

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setViewMode(QListWidget.IconMode)
        self.setIconSize(QSize(120, 160))
        self.setSpacing(10)
        self.setResizeMode(QListWidget.Adjust)
        
        # Drag & Drop Support
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QListWidget.DragDrop) 
        self.setSelectionMode(QListWidget.ExtendedSelection)
        self.setDefaultDropAction(Qt.MoveAction)
        
        # Context Menu
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        # self.customContextMenuRequested.connect(self.contextMenuRequested) # Removed redundant connection

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            # Handle File Drop (External)
            event.accept()
            urls = event.mimeData().urls()
            paths = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith('.pdf')]
            if paths:
                self.filesDropped.emit(paths)
        else:
            # Handle Internal/Cross-List Drop
            self.aboutToChange.emit()
            super().dropEvent(event)


class PDFEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF彙編器")
        self.resize(1300, 900)
        
        # Data Registry
        # source_docs: List of { 'doc': fitz.Document, 'path': str, 'id': int }
        self.source_docs = [] 
        self.doc_counter = 0

        # State & Cache
        self.thumbnail_cache = ThumbnailCache()
        self.history = HistoryManager()
        # self.clipboard_pages = [] # Removed clipboard, using direct duplicate

        # Keyboard Shortcuts
        self.undo_action = QAction("Undo", self)
        self.undo_action.setShortcut("Ctrl+Z")
        self.undo_action.triggered.connect(self.undo_operation)
        self.addAction(self.undo_action)

        self.redo_action = QAction("Redo", self)
        self.redo_action.setShortcut("Ctrl+Y")
        self.redo_action.triggered.connect(self.redo_operation)
        self.addAction(self.redo_action)
        
        self.copy_action = QAction("Duplicate", self)
        self.copy_action.setShortcut("Ctrl+D") # Changed from C to D for Duplicate commonly, or keep C? Let's use D or C. User accustomed to C maybe?
        # Standard copy is Ctrl+C, but here it's duplicate. Let's keep Ctrl+C for "Copy/Duplicate"
        self.copy_action.setShortcut("Ctrl+C") 
        self.copy_action.triggered.connect(self.duplicate_pages_op)
        self.addAction(self.copy_action)

        # Delete Key
        self.del_action = QAction("Delete", self)
        self.del_action.setShortcut("Delete")
        self.del_action.triggered.connect(self.delete_pages)
        self.addAction(self.del_action)

        # Setup UI
        self.setup_ui()
        self.apply_styles()
        
    def apply_styles(self):
        self.setStyleSheet(DARK_THEME_QSS)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # --- Sidebar ---
        sidebar = self.create_sidebar()
        main_layout.addWidget(sidebar)

        # --- Right Content (Splitter) ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(10, 10, 10, 10)

        self.splitter = QSplitter(Qt.Vertical)
        
        # 1. Main Composition Area
        self.main_area_widget = QWidget()
        vbox_main = QVBoxLayout(self.main_area_widget)
        vbox_main.setContentsMargins(0,0,0,0)
        
        lbl_main = QLabel("主要頁面 (Main Composition) - 最終產出的順序")
        lbl_main.setObjectName("SectionHeader")
        vbox_main.addWidget(lbl_main)
        
        self.main_list = PDFPageList()
        # Connect internal move signal handled by default, but we might want status
        self.main_list = PDFPageList()
        # Connect internal move signal handled by default, but we might want status
        self.main_list.aboutToChange.connect(self.capture_state)
        # Context Menu
        self.main_list.customContextMenuRequested.connect(self.show_context_menu)
        # DISABLE Drag Reordering in Main List (DropOnly allows drops from outside/Staging, but not Dragging items internally)
        self.main_list.setDragDropMode(QAbstractItemView.DragDrop)
        vbox_main.addWidget(self.main_list)
        
        self.splitter.addWidget(self.main_area_widget)

        # 2. Staging Area
        self.staging_area_widget = QWidget()
        vbox_staging = QVBoxLayout(self.staging_area_widget)
        vbox_staging.setContentsMargins(0,0,0,0)
        
        lbl_staging = QLabel("預備區 (Staging Area) - 拖曳檔案至此，再選取頁面加入上方")
        lbl_staging.setObjectName("SectionHeader")
        vbox_staging.addWidget(lbl_staging)
        
        self.staging_list = PDFPageList()
        # Handle file drops on both, but typically staging is for drops
        self.staging_list.filesDropped.connect(self.load_pdfs_to_staging)
        # Enable Drag from Staging (Default is DragDrop, which is fine, or DragOnly)
        self.staging_list.customContextMenuRequested.connect(self.show_context_menu)
        
        # Allow dropping files on main list too? Sure.
        self.main_list.filesDropped.connect(self.load_pdfs_to_staging) 

        vbox_staging.addWidget(self.staging_list)
        
        self.splitter.addWidget(self.staging_area_widget)
        self.splitter.setStretchFactor(0, 2)
        self.splitter.setStretchFactor(1, 1)

        right_layout.addWidget(self.splitter)
        
        # Progress Bar (Hidden by default)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("QProgressBar { border: 1px solid #3e3e42; border-radius: 5px; text-align: center; } QProgressBar::chunk { background-color: #007acc; }")
        right_layout.addWidget(self.progress_bar)
        
        # Status Bar
        self.status_label = QLabel("就緒 (Ready)")
        self.status_label.setObjectName("StatusLabel")
        right_layout.addWidget(self.status_label)

        main_layout.addWidget(right_panel)

    def create_sidebar(self):
        sidebar = QFrame()
        sidebar.setFixedWidth(280)
        sidebar.setStyleSheet("background-color: #252526; border-right: 1px solid #3e3e42;")
        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(15, 20, 15, 20)
        layout.setSpacing(15)
        
        lbl_tools = QLabel("工具箱 (Tools)")
        lbl_tools.setStyleSheet("font-size: 18px; font-weight: bold; color: #007acc; border: none;")
        layout.addWidget(lbl_tools)
        
        # File Group
        grp_file = QGroupBox("檔案 (File)")
        vbox = QVBoxLayout()
        btn_add = QPushButton("加入 PDF (Import)")
        btn_add.clicked.connect(self.open_pdf_dialog)
        vbox.addWidget(btn_add)
        
        btn_save = QPushButton("儲存 PDF (Save)")
        btn_save.clicked.connect(self.save_pdf)
        btn_save.setStyleSheet("background-color: #2ea043;")
        vbox.addWidget(btn_save)
        grp_file.setLayout(vbox)
        layout.addWidget(grp_file)

        # Output Settings Group
        grp_out = QGroupBox("輸出設定 (Export Settings)")
        vbox_out = QVBoxLayout()
        vbox_out.setSpacing(8)
        
        # Text
        vbox_out.addWidget(QLabel("頁碼/文字 (Text):"))
        
        # Checkbox Enabled
        self.chk_overlay_enable = QCheckBox("顯示頁碼文字 (Overlay Text)")
        self.chk_overlay_enable.setChecked(False) # Default Off
        vbox_out.addWidget(self.chk_overlay_enable)
        
        self.txt_overlay = QLineEdit("{n} / {total}")
        self.txt_overlay.setPlaceholderText("e.g. Page {n}")
        self.txt_overlay.setToolTip("可用變數: {n}=當前頁碼, {total}=總頁數, {name}=頁面名稱")
        vbox_out.addWidget(self.txt_overlay)
        
        # Enable text entry only if Checked (Optional, but nice)
        self.txt_overlay.setEnabled(False)
        self.chk_overlay_enable.toggled.connect(self.txt_overlay.setEnabled)
        
        # Helper Buttons for Text
        hbox_text_btns = QHBoxLayout()
        hbox_text_btns.setSpacing(2)
        
        btn_ins_n = QPushButton("插入頁碼")
        btn_ins_n.setToolTip("插入 {n}")
        btn_ins_n.setStyleSheet("padding: 2px 5px; font-size: 11px;")
        btn_ins_n.clicked.connect(lambda: self.insert_overlay_text("{n}"))
        hbox_text_btns.addWidget(btn_ins_n)

        btn_ins_total = QPushButton("插入總頁")
        btn_ins_total.setToolTip("插入 {total}")
        btn_ins_total.setStyleSheet("padding: 2px 5px; font-size: 11px;")
        btn_ins_total.clicked.connect(lambda: self.insert_overlay_text("{total}"))
        hbox_text_btns.addWidget(btn_ins_total)

        btn_ins_name = QPushButton("插入頁名")
        btn_ins_name.setToolTip("插入 {name}")
        btn_ins_name.setStyleSheet("padding: 2px 5px; font-size: 11px;")
        btn_ins_name.clicked.connect(lambda: self.insert_overlay_text("{name}"))
        hbox_text_btns.addWidget(btn_ins_name)
        
        vbox_out.addLayout(hbox_text_btns)
        
        # Position
        vbox_out.addWidget(QLabel("位置 (Position):"))
        self.combo_pos = QComboBox()
        self.combo_pos.addItems([
            "Bottom-Right", "Bottom-Center", "Bottom-Left",
            "Top-Right", "Top-Center", "Top-Left",
            "Middle-Right", "Middle-Left" # Removed Middle-Center/Middle to keep it simple or full?
        ])
        # Add remaining if needed or just keep standard
        vbox_out.addWidget(self.combo_pos)
        
        # Color & Size
        hbox_style = QHBoxLayout()
        self.combo_color = QComboBox()
        self.combo_color.addItems(["Black", "White", "Red", "Blue", "Gray"])
        hbox_style.addWidget(self.combo_color)
        
        self.spin_size = QSpinBox()
        self.spin_size.setRange(8, 72)
        self.spin_size.setValue(12)
        self.spin_size.setSuffix(" pt")
        hbox_style.addWidget(self.spin_size)
        
        vbox_out.addLayout(hbox_style)
        
        grp_out.setLayout(vbox_out)
        layout.addWidget(grp_out)
        
        # Edit Group
        grp_edit = QGroupBox("編輯 (Edit)")
        vbox = QVBoxLayout()
        
        btn_cw = QPushButton("向右旋轉 (Rotate CW)")
        btn_cw.clicked.connect(lambda: self.rotate_pages(90))
        vbox.addWidget(btn_cw)
        
        btn_ccw = QPushButton("向左旋轉 (Rotate CCW)")
        btn_ccw.clicked.connect(lambda: self.rotate_pages(-90))
        vbox.addWidget(btn_ccw)
        
        # Move Buttons
        hbox_move = QHBoxLayout()
        btn_left = QPushButton("<< 前移")
        btn_left.clicked.connect(self.move_page_left)
        btn_left.setStyleSheet("background-color: #444444;")
        
        btn_right = QPushButton("後移 >>")
        btn_right.clicked.connect(self.move_page_right)
        btn_right.setStyleSheet("background-color: #444444;")
        
        hbox_move.addWidget(btn_left)
        hbox_move.addWidget(btn_right)
        vbox.addLayout(hbox_move)
        
        # Undo / Redo Row
        hbox_undo = QHBoxLayout()
        btn_undo = QPushButton("↩ 復原")
        btn_undo.clicked.connect(self.undo_operation)
        btn_undo.setStyleSheet("background-color: #555555;")
        
        btn_redo = QPushButton("重做 ↪")
        btn_redo.clicked.connect(self.redo_operation)
        btn_redo.setStyleSheet("background-color: #555555;")
        
        hbox_undo.addWidget(btn_undo)
        hbox_undo.addWidget(btn_redo)
        vbox.addLayout(hbox_undo)
        
        # New Buttons
        btn_copy = QPushButton("複製 (Duplicate)")
        btn_copy.clicked.connect(self.duplicate_pages_op)
        vbox.addWidget(btn_copy)

        btn_rename = QPushButton("重新命名 (Rename)")
        btn_rename.clicked.connect(lambda: self.rename_page_op(None))
        vbox.addWidget(btn_rename)
        
        btn_del = QPushButton("刪除頁面 (Delete)")
        btn_del.clicked.connect(self.delete_pages)
        btn_del.setStyleSheet("background-color: #d73a49;")
        vbox.addWidget(btn_del)
        
        grp_edit.setLayout(vbox)
        layout.addWidget(grp_edit)



        layout.addStretch()
        layout.addWidget(QLabel("Copyright © Liyuchiutiger Gongminshen"))
        return sidebar

    # --- Logic ---

    def insert_overlay_text(self, text):
        self.txt_overlay.insert(text)
        self.txt_overlay.setFocus()

    def open_pdf_dialog(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "加入檔案", "", 
                                                "Supported Files (*.pdf *.png *.jpg *.jpeg *.bmp);;PDF Files (*.pdf);;Images (*.png *.jpg *.jpeg *.bmp)")
        if paths:
            self.load_pdfs_to_staging(paths)

    def load_pdfs_to_staging(self, paths):
        # We start a worker for each, or one worker for all.
        self.status_label.setText(f"正在載入 {len(paths)} 個檔案...")
        
        for path in paths:
            self._load_single_pdf(path)

    def _load_single_pdf(self, path):
        # 1. Register Doc
        try:
            # Read into memory to avoid file lock on Windows which prevents saving/overwriting
            with open(path, "rb") as f:
                file_bytes = f.read()
            
            # Get extension logic
            ext = os.path.splitext(path)[1].lower().strip(".")
            if not ext:
                ext = "pdf" # Default assumption
                
            doc = fitz.open(ext, file_bytes)
            
            # Handle Images by converting to PDF in-memory
            if not doc.is_pdf:
                try:
                    pdf_bytes = doc.convert_to_pdf()
                    doc = fitz.open("pdf", pdf_bytes)
                    doc.set_metadata({'title': os.path.basename(path)}) # Set title from original filename
                except Exception as img_err:
                    print(f"Conversion failed for {path}: {img_err}")
                    # Try to continue if possible, matches fitz logic
            
            doc_id = self.doc_counter
            self.doc_counter += 1
            
            entry = {'doc': doc, 'path': path, 'id': doc_id}
            self.source_docs.append(entry)
            
            # 2. Worker to generate thumbnails
            worker = PDFWorker(self._gen_thumbnails, doc, doc_id)
            worker.finished.connect(self._on_thumbnails_ready)
            worker.run() 
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load {path}: {e}")

    def _gen_thumbnails(self, doc, doc_id):
        items_data = []
        for i in range(len(doc)):
            page = doc.load_page(i)
            # Low res for thumbnail
            pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
            img_bytes = pix.tobytes("png")
            items_data.append((doc_id, i, img_bytes))
        return items_data

    def _on_thumbnails_ready(self, items_data):
        if isinstance(items_data, Exception):
            return 
        
        # Add to Staging List
        for doc_id, page_num, img_bytes in items_data:
            # Cache the base image
            img = QImage.fromData(img_bytes)
            self.thumbnail_cache.set_image(doc_id, page_num, img)
            
            icon = QIcon(QPixmap.fromImage(img))
            item = QListWidgetItem(icon, f"P{page_num + 1}")
            
            # STORE DATA
            # UserRole: Page Index
            # UserRole+1: Rotation
            # UserRole+2: Doc ID
            # UserRole+3: Original Label (if any) - mostly just "P{num}"
            item.setData(Qt.UserRole, page_num)
            item.setData(Qt.UserRole + 1, 0)
            item.setData(Qt.UserRole + 2, doc_id)
            # Tooltip
            item.setToolTip(f"Doc ID: {doc_id} | Page: {page_num + 1}")
            
            self.staging_list.addItem(item)
            
        self.status_label.setText("已將檔案加入預備區 (Added files to Staging Area)")

    def get_doc_by_id(self, doc_id):
        for entry in self.source_docs:
            if entry['id'] == doc_id:
                return entry['doc']
        return None

    # --- History & State ---
    
    def capture_state(self):
        """Captures the current state of Main List only"""
        state = []
        for i in range(self.main_list.count()):
            item = self.main_list.item(i)
            data = {
                'doc_id': item.data(Qt.UserRole + 2),
                'page_num': item.data(Qt.UserRole),
                'rotation': item.data(Qt.UserRole + 1),
                'text': item.text()
            }
            state.append(data)
        self.history.push_state(state)

    def undo_operation(self):
        if not self.history.can_undo():
            self.status_label.setText("沒有動作可復原 (Nothing to Undo)")
            return
        
        # Capture current state to Redo Stack before restoring old state
        current_state = self.get_current_state_data()
        self.history.push_to_redo(current_state)
            
        state = self.history.pop_undo()
        self.restore_state(state)
        self.status_label.setText("已復原 (Undone)")

    def redo_operation(self):
        if not self.history.can_redo():
            self.status_label.setText("沒有動作可重做 (Nothing to Redo)")
            return
            
        # Capture current state to Undo Stack (but don't clear Redo)
        current_state = self.get_current_state_data()
        self.history.undo_stack.append(current_state) # Manually append to avoid clearing redo
            
        state = self.history.pop_redo()
        self.restore_state(state)
        self.status_label.setText("已重做 (Redone)")

    def get_current_state_data(self):
        """Helper to get state data without pushing to stack"""
        state = []
        for i in range(self.main_list.count()):
            item = self.main_list.item(i)
            data = {
                'doc_id': item.data(Qt.UserRole + 2),
                'page_num': item.data(Qt.UserRole),
                'rotation': item.data(Qt.UserRole + 1),
                'text': item.text()
            }
            state.append(data)
        return state

    def restore_state(self, state):
        self.main_list.clear() # This might trigger signals? No, programmatic changes don't usually invoke drag signals.
        for data in state:
            doc_id = data['doc_id']
            page_num = data['page_num']
            rotation = data['rotation']
            text = data.get('text', f"P{page_num+1}")
            
            item = QListWidgetItem()
            item.setText(text)
            item.setData(Qt.UserRole, page_num)
            item.setData(Qt.UserRole + 1, rotation)
            item.setData(Qt.UserRole + 2, doc_id)
            
            self.main_list.addItem(item)
            self.update_item_thumbnail(item)

    def rotate_pages(self, angle):
        # Check which list is focused or has selection
        target_list = None
        if self.main_list.hasFocus():
            target_list = self.main_list
        elif self.staging_list.hasFocus():
            target_list = self.staging_list
        else:
            # Fallback: check selection count
            if len(self.main_list.selectedItems()) > 0:
                target_list = self.main_list
            elif len(self.staging_list.selectedItems()) > 0:
                target_list = self.staging_list
        
        if not target_list: 
            return

        # Capture State only if modifying Main List
        if target_list == self.main_list:
            self.capture_state()

        items = target_list.selectedItems()
        for item in items:
            current_rot = item.data(Qt.UserRole + 1) or 0
            new_rot = (current_rot + angle) % 360
            item.setData(Qt.UserRole + 1, new_rot)
            
            # UPDATE VISUAL
            self.update_item_thumbnail(item)

    def update_item_thumbnail(self, item):
        doc_id = item.data(Qt.UserRole + 2)
        page_num = item.data(Qt.UserRole)
        rotation = item.data(Qt.UserRole + 1)
        
        # Get from Cache
        base_img = self.thumbnail_cache.get_image(doc_id, page_num)
        
        # If not in cache (e.g. from saved project? or cleared?), try to reload?
        if not base_img:
            doc = self.get_doc_by_id(doc_id)
            if doc:
                 page = doc.load_page(page_num)
                 pix = page.get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
                 base_img = QImage.fromData(pix.tobytes("png"))
                 self.thumbnail_cache.set_image(doc_id, page_num, base_img)
        
        if base_img:
            # Rotation Preview
            if rotation != 0:
                import PySide6.QtGui
                tr = PySide6.QtGui.QTransform()
                tr.rotate(rotation)
                final_img = base_img.transformed(tr)
            else:
                final_img = base_img
                
            icon = QIcon(QPixmap.fromImage(final_img))
            item.setIcon(icon)
            
            # Update Tooltip
            item.setToolTip(f"Doc: {doc_id} | Page: {page_num+1} | Rot: {rotation}°")

    def delete_pages(self):
        # Delete from whichever list is active
        # If Main List, capture state
        target_list = None
        if self.main_list.hasFocus():
            target_list = self.main_list
        elif self.staging_list.hasFocus():
            target_list = self.staging_list
        else:
             if len(self.main_list.selectedItems()) > 0:
                 target_list = self.main_list
        
        if not target_list: return
        
        if target_list == self.main_list and len(target_list.selectedItems()) > 0:
             self.capture_state()

        items = target_list.selectedItems()
        if not items: return
        
        # Sort reverse to avoid index shifting issues
        rows = sorted([target_list.row(item) for item in items], reverse=True)
        for row in rows:
            target_list.takeItem(row)

    def move_page_left(self):
        self._move_page_selection(-1)

    def move_page_right(self):
        self._move_page_selection(1)

    def _move_page_selection(self, delta):
        # Only for Main List? User said "Modify order... click Left/Right"
        # Usually implies Main List. Staging supports reorder? Maybe.
        # Let's target Main List for now as that's the primary output.
        target_list = self.main_list
        if not target_list.hasFocus() and len(target_list.selectedItems()) == 0:
            return

        items = target_list.selectedItems()
        if not items: return
        
        # Capture Sort Order
        self.capture_state()
        
        # Get rows
        rows = sorted([target_list.row(item) for item in items])
        
        if delta < 0: # Move Up/Left
            if rows[0] == 0: return # Already at top
            
            # Move items one by one from top to bottom index
            for row in rows:
                item = target_list.takeItem(row)
                target_list.insertItem(row - 1, item)
                item.setSelected(True)
                
        else: # Move Down/Right
            if rows[-1] == target_list.count() - 1: return # Already at bottom
            
            # Move items one by one from bottom to top index
            for row in reversed(rows):
                item = target_list.takeItem(row)
                target_list.insertItem(row + 1, item)
                item.setSelected(True)
                
        # Scroll to ensure visible
        target_list.scrollToItem(items[0])

    # --- Copy / Paste / Rename / Menu ---
    
    def show_context_menu(self, pos):
        sender = self.sender()
        if not isinstance(sender, QListWidget):
             sender = self.main_list # Fallback?
             
        item = sender.itemAt(pos)
        
        menu = QMenu(self)
        
        qt_copy = QAction("複製 (Copy/Add)", self) # Duplicate if Main, Add if Staging
        qt_copy.triggered.connect(self.duplicate_pages_op)
        menu.addAction(qt_copy)
        
        # Paste removed as per request
        
        menu.addSeparator()
        
        if item:
            qt_rename = QAction("重新命名 (Rename)", self)
            qt_rename.triggered.connect(lambda: self.rename_page_op(item))
            menu.addAction(qt_rename)
            
            qt_del = QAction("刪除 (Delete)", self)
            qt_del.triggered.connect(self.delete_pages)
            menu.addAction(qt_del)
            
            menu.addSeparator()
            
            qt_cw = QAction("向右旋轉 (Rotate CW)", self)
            qt_cw.triggered.connect(lambda: self.rotate_pages(90))
            menu.addAction(qt_cw)
            
            qt_ccw = QAction("向左旋轉 (Rotate CCW)", self)
            qt_ccw.triggered.connect(lambda: self.rotate_pages(-90))
            menu.addAction(qt_ccw)

        menu.exec_(sender.mapToGlobal(pos))

    def duplicate_pages_op(self):
        """Duplicates selected pages immediately."""
        # Determine source
        source_list = None
        if self.main_list.hasFocus() or len(self.main_list.selectedItems()) > 0:
            source_list = self.main_list
            target_list = self.main_list # Duplicate into main
        elif self.staging_list.hasFocus() or len(self.staging_list.selectedItems()) > 0:
            source_list = self.staging_list
            target_list = self.main_list # Add to main
        
        if not source_list: return
        
        items = source_list.selectedItems()
        if not items: return
        
        # If modifying main list, capture state first
        if target_list == self.main_list:
            self.capture_state()
            
        # Collect Data
        new_items_data = []
        for item in items:
            data = {
                'doc_id': item.data(Qt.UserRole + 2),
                'page_num': item.data(Qt.UserRole),
                'rotation': item.data(Qt.UserRole + 1),
                'text': item.text()
            }
            new_items_data.append(data)
            
        # Insertion Point
        if source_list == self.main_list:
            # Insert after last selected item
            last_sel = items[-1]
            row = self.main_list.row(last_sel) + 1
        else:
            # Append to end
            row = self.main_list.count()
            
        # Create Items
        for data in new_items_data:
            doc_id = data['doc_id']
            page_num = data['page_num']
            rotation = data['rotation']
            text = data.get('text', f"P{page_num+1}")
            
            # If coming from Staging, use clean text. If from Main, maybe append Copy?
            # User said "Just duplicate", usually implies Copy of...
            # But from Staging it means "Add".
            if source_list == self.staging_list:
                final_text = text
            else:
                 # If it already says (Copy), maybe don't stack it infinitely or do?
                 final_text = text + " (Copy)"
            
            item = QListWidgetItem()
            item.setText(final_text)
            item.setData(Qt.UserRole, page_num)
            item.setData(Qt.UserRole + 1, rotation)
            item.setData(Qt.UserRole + 2, doc_id)
            
            self.main_list.insertItem(row, item)
            self.update_item_thumbnail(item)
            row += 1
            
        self.status_label.setText(f"已複製/加入 {len(new_items_data)} 頁 (Duplicated/Added)")
        
        # Paste Op Removed

    def rename_page_op(self, item_arg=None):
        item = item_arg
        if not item:
            # Check focus
            if self.main_list.hasFocus() or len(self.main_list.selectedItems()) > 0:
                 selected = self.main_list.selectedItems()
                 if selected: item = selected[0]
            elif self.staging_list.hasFocus() or len(self.staging_list.selectedItems()) > 0:
                 selected = self.staging_list.selectedItems()
                 if selected: item = selected[0]
        
        if not item: return
        
        old_text = item.text()
        new_text, ok = QInputDialog.getText(self, "重新命名", "輸入新名稱:", text=old_text)
        if ok and new_text:
            # If changing Main List, capture state
            if item.listWidget() == self.main_list:
                self.capture_state()
                
            item.setText(new_text)
            self.status_label.setText(f"已重新命名: {old_text} -> {new_text}")

    def save_pdf(self):
        if self.main_list.count() == 0:
            QMessageBox.warning(self, "Warning", "主要頁面是空的 (Main composition is empty)")
            return
            
        out_path, _ = QFileDialog.getSaveFileName(self, "儲存 PDF", "", "PDF Files (*.pdf)")
        if not out_path: return
        
        # UI LOCK
        self.status_label.setText("儲存中... (Saving...)")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, self.main_list.count())
        self.progress_bar.setValue(0)
        self.setEnabled(False) # Disable UI
        
        # GATHER DATA
        items_data = []
        for i in range(self.main_list.count()):
            item = self.main_list.item(i)
            items_data.append({
                'doc_id': item.data(Qt.UserRole + 2),
                'page_num': item.data(Qt.UserRole),
                'rotation': item.data(Qt.UserRole + 1),
                'text': item.text() # Pass current name
            })
            
        overlays = {
            'enabled': self.chk_overlay_enable.isChecked(),
            'text': self.txt_overlay.text(),
            'pos': self.combo_pos.currentText(),
            'color': self.combo_color.currentText(),
            'size': self.spin_size.value()
        }
        
        # START WORKER
        self.save_worker = SaveWorker(items_data, self.source_docs, out_path, overlays)
        self.save_worker.progress.connect(self.on_save_progress)
        self.save_worker.finished.connect(self.on_save_finished)
        self.save_worker.start()

    def on_save_progress(self, current, total):
        self.progress_bar.setValue(current)
        self.status_label.setText(f"儲存中... {current}/{total}")

    def on_save_finished(self, success, msg):
        self.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_label.setText(msg if success else "儲存失敗 (Save Failed)")
        
        if success:
            QMessageBox.information(self, "成功 (Success)", msg)
        else:
            QMessageBox.critical(self, "錯誤 (Error)", f"儲存失敗:\n{msg}")
            




if __name__ == "__main__":
    app = QApplication(sys.argv)
    font = QFont("Microsoft JhengHei", 10)
    app.setFont(font)
    window = PDFEditor()
    window.show()
    sys.exit(app.exec())
