# widgets.py
import sys
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QFrame, QFormLayout, QTextEdit, QCheckBox
)
from PyQt5.QtGui import QPixmap, QCursor
from PyQt5.QtCore import Qt, QEvent, QPoint

class FormItemWidget(QWidget):
    def __init__(self, item_data=None, parent=None):
        super().__init__(parent)
        self.init_ui(item_data)

    def init_ui(self, item_data):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # 使用 QFrame 包覆所有元件並設定邊框
        frame = QFrame()
        frame.setObjectName("outerFrame")
        frame.setStyleSheet("#outerFrame { border: 2px solid black; border-radius: 15px; }")
        frame_layout = QHBoxLayout(frame)
        frame_layout.setContentsMargins(5, 5, 5, 5)
        frame_layout.setSpacing(10)
        
        # 左側：表單輸入欄位
        inputs_widget = QWidget()
        form_layout = QFormLayout(inputs_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(5)
        
        self.description_input = QTextEdit()
        self.description_input.setFixedHeight(80)
        label = QLabel("施工說明：")
        label.setStyleSheet("font-weight: bold;")
        form_layout.addRow(label, self.description_input)
        
        self.time_input = QLineEdit()
        time_label = QLabel("時間：")
        time_label.setStyleSheet("font-weight: bold;")
        form_layout.addRow(time_label, self.time_input)
        
        # 圖片路徑與瀏覽按鈕
        self.image_path_input = QLineEdit()
        self.image_browse_button = QPushButton("瀏覽")
        self.image_browse_button.clicked.connect(self.browse_image)
        image_layout = QHBoxLayout()
        image_layout.addWidget(self.image_path_input)
        image_layout.addWidget(self.image_browse_button)
        image_label = QLabel("圖片：")
        image_label.setStyleSheet("font-weight: bold;")
        form_layout.addRow(image_label, image_layout)
        
        self.time_checkbox = QCheckBox("是否標註時間")
        form_layout.addRow(self.time_checkbox)
        
        frame_layout.addWidget(inputs_widget)
        
        # 右側：圖片預覽框
        self.image_preview_label = QLabel("圖片預覽")
        self.image_preview_label.setFixedSize(500, 300)
        self.image_preview_label.setAlignment(Qt.AlignCenter)
        frame_layout.addWidget(self.image_preview_label, stretch=1, alignment=Qt.AlignTop)
        self.image_preview_label.installEventFilter(self)
        
        main_layout.addWidget(frame)
        
        if item_data:
            self.set_data(item_data)

    def set_data(self, item_data):
        self.description_input.setPlainText(item_data.get('施工說明', ''))
        self.time_input.setText(item_data.get('時間', ''))
        self.image_path_input.setText(item_data.get('圖片路徑', ''))
        path = item_data.get('圖片路徑', '')
        if path:
            pixmap = QPixmap(path).scaled(500, 300, Qt.KeepAspectRatio)
            self.image_preview_label.setPixmap(pixmap)
            self.image_preview_label.setProperty("original_pixmap", QPixmap(path))
        self.time_checkbox.setChecked(item_data.get('標註時間', False))

    def get_data(self):
        return {
            '施工說明': self.description_input.toPlainText(),
            '時間': self.time_input.text(),
            '圖片路徑': self.image_path_input.text(),
            '標註時間': self.time_checkbox.isChecked()
        }

    def browse_image(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "選擇圖片文件", "",
            "Images (*.png *.jpg *.jpeg *.bmp);;All Files (*)", options=options
        )
        if file_path:
            self.image_path_input.setText(file_path)
            pixmap = QPixmap(file_path).scaled(500, 300, Qt.KeepAspectRatio)
            self.image_preview_label.setPixmap(pixmap)
            self.image_preview_label.setProperty("original_pixmap", QPixmap(file_path))

    def eventFilter(self, source, event):
        if source == self.image_preview_label and self.image_preview_label.property("original_pixmap") is not None:
            if event.type() == QEvent.Enter:
                if not hasattr(self, 'zoom_label') or self.zoom_label is None:
                    self.zoom_label = QLabel(self)
                    self.zoom_label.setWindowFlags(Qt.ToolTip)
                original_pixmap = self.image_preview_label.property("original_pixmap")
                self.zoom_label.setPixmap(original_pixmap.scaled(1600, 900, Qt.KeepAspectRatio))
                cursor_pos = QCursor.pos()
                self.zoom_label.move(cursor_pos + QPoint(20, 20))
                self.zoom_label.show()
            elif event.type() == QEvent.Leave:
                if hasattr(self, 'zoom_label') and self.zoom_label is not None:
                    self.zoom_label.hide()
        return super().eventFilter(source, event)