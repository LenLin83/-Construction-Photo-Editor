import sys
import os
import json
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog,
    QMessageBox, QTextEdit, QCheckBox, QVBoxLayout, QHBoxLayout, QFormLayout,
    QComboBox, QListWidget, QListWidgetItem, QAbstractItemView, QFrame,
    QGraphicsDropShadowEffect
)
from PyQt5.QtGui import QPixmap, QCursor
from PyQt5.QtCore import Qt, QEvent, QPoint
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

#----------------------------------
# 自訂單一案件項目的 widget
# 將所有控制項包覆在一個 QFrame 中，並設定 2px 黑色邊框
#----------------------------------
class FormItemWidget(QWidget):
    
    def __init__(self, item_data=None, parent=None):
        super().__init__(parent)
        self.init_ui(item_data)

    def init_ui(self, item_data):
        # 主佈局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # 用 QFrame 包覆所有元件，並只對這個 QFrame 設定 2px 黑色邊框
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
        # self.image_browse_button.setFlat(True)  # 平坦風格，不顯示邊框
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
        # self.image_preview_label.setStyleSheet("border: 2px solid gray; border-radius: 5px;")
        # 移除圖片預覽框的額外邊框設定
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
        # 當滑鼠進入圖片預覽框時顯示放大圖片
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

#----------------------------------
# 主程式
#----------------------------------
class ConstructionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.image_bytes_list = []      # 用來保存 BytesIO 物件，方便釋放資源
        self.saved_data_file = "saved_data.json"  # 保存案件資料的檔案

        if hasattr(sys, '_MEIPASS'):
            self.template_path = os.path.join(sys._MEIPASS, "施工照片.docx")
        else:
            self.template_path = "施工照片.docx"
        self.doc = DocxTemplate(self.template_path)

        self.init_ui()
        self.load_saved_projects()

    def init_ui(self):
        self.setWindowTitle("施工照片生成器")
        self.setGeometry(100, 100, 1200, 800)
        main_layout = QVBoxLayout(self)
        shadow = QGraphicsDropShadowEffect()


        # 案件基本資訊區
        project_info_layout = QHBoxLayout()
        self.project_selector = QComboBox()
        self.project_selector.addItem("選擇或創建新案子")
        self.project_selector.currentIndexChanged.connect(self.load_selected_project)
        project_info_layout.addWidget(self.project_selector)

        self.case_id_label = QLabel("案件編號：")
        self.case_id_label.setStyleSheet("font-weight: bold;")
        self.case_id_input = QLineEdit()
        project_info_layout.addWidget(self.case_id_label)
        project_info_layout.addWidget(self.case_id_input)

        self.case_address_label = QLabel("案件地址：")
        self.case_address_label.setStyleSheet("font-weight: bold;")
        self.case_address_input = QLineEdit()
        project_info_layout.addWidget(self.case_address_label)
        project_info_layout.addWidget(self.case_address_input)

        main_layout.addLayout(project_info_layout)

        # 僅保留 QListWidget（取消上下箭頭移動功能）
        self.item_list = QListWidget()
        self.item_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.item_list.setSpacing(20)
        self.item_list.setStyleSheet("""
    QListWidget::item:hover {
        background-color: #CCE5FF;  /* 淺藍色，可自行調整 */
        border-radius: 15px;
        padding: 5px;
    }
    QListWidget::item:selected {
        # background-color: #A8D4FF;  /* 更亮的藍色，可自行調整 */
        border-radius: 15px;
        padding: 5px;
    }
""")
        main_layout.addWidget(self.item_list)

        # 按鈕區：新增項目、刪除選取項目、移除案子
        btn_layout = QHBoxLayout()
        
        self.add_item_button = QPushButton("新增內容項目")
        self.add_item_button.setStyleSheet(
            """
            QPushButton {
                background-color: #CCF1FF;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #B3E5FF;
            }
            QPushButton:pressed {
                background-color: #99D6FF;
            }
            """
        )
        self.add_item_button.clicked.connect(lambda: self.add_form_item())

        # 建立陰影效果並設定到按鈕上
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)       # 模糊半徑
        shadow.setOffset(0, 0)        # X, Y 方向的偏移
        shadow.setColor(Qt.gray)      # 陰影顏色
        self.add_item_button.setGraphicsEffect(shadow)
        
        btn_layout.addWidget(self.add_item_button)

        
        self.delete_selected_button = QPushButton("刪除選取項目")
        self.delete_selected_button.setStyleSheet(
            """
            QPushButton {
                background-color: #FCFFCC;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #F0FFB2;
            }
            QPushButton:pressed {
                background-color: #E6FF99;
            }
            """
        )
        self.delete_selected_button.clicked.connect(self.delete_selected_items)
        
        # 建立陰影效果並設定到按鈕上
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)       # 模糊半徑
        shadow.setOffset(0, 0)        # X, Y 方向的偏移
        shadow.setColor(Qt.gray)      # 陰影顏色
        self.delete_selected_button.setGraphicsEffect(shadow)
        
        btn_layout.addWidget(self.delete_selected_button)
        
        self.remove_project_button = QPushButton("移除案子")
        self.remove_project_button.setStyleSheet(
            """
            QPushButton {
                background-color: #FFCCCC;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FFB3B3;
            }
            QPushButton:pressed {
                background-color: #FF9999;
            }
            """
        )
        self.remove_project_button.clicked.connect(self.remove_project)
        
        # 建立陰影效果並設定到按鈕上
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(10)       # 模糊半徑
        shadow.setOffset(0, 0)        # X, Y 方向的偏移
        shadow.setColor(Qt.gray)      # 陰影顏色
        self.remove_project_button.setGraphicsEffect(shadow)
        
        btn_layout.addWidget(self.remove_project_button)
        
        main_layout.addLayout(btn_layout)
        
        # 燒光碟選項與生成文檔按鈕
        self.burn_disc_checkbox = QCheckBox("是否燒光碟")
        main_layout.addWidget(self.burn_disc_checkbox)
        
        self.generate_button = QPushButton("生成文檔")
        self.generate_button.setStyleSheet(
            """
            QPushButton {
                background-color: #9EFFB5;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7DFF99;
            }
            QPushButton:pressed {
                background-color: #66FF80;
            }
            """
        )
        self.generate_button.clicked.connect(self.generate_document)
        main_layout.addWidget(self.generate_button)
        
        self.setLayout(main_layout)

    def add_form_item(self, item_data=None):
        widget = FormItemWidget(item_data)
        list_item = QListWidgetItem()
        list_item.setSizeHint(widget.sizeHint())
        self.item_list.addItem(list_item)
        self.item_list.setItemWidget(list_item, widget)

    def delete_selected_items(self):
        for item in self.item_list.selectedItems():
            row = self.item_list.row(item)
            self.item_list.takeItem(row)

    def generate_document(self):
        if self.item_list.count() == 0:
            QMessageBox.warning(self, "警告", "請至少添加一組內容")
            return

        case_id = self.case_id_input.text()
        case_address = self.case_address_input.text()
        if not case_id or not case_address:
            QMessageBox.warning(self, "警告", "請輸入案件編號和地址")
            return

        try:
            processed_items = []
            folder_name = f"照片-{case_id}-{case_address}"
            if self.burn_disc_checkbox.isChecked():
                if not os.path.exists(folder_name):
                    os.makedirs(folder_name)

            for i in range(self.item_list.count()):
                item = self.item_list.item(i)
                widget = self.item_list.itemWidget(item)
                data = widget.get_data()

                description = data['施工說明']
                time = data['時間']
                image_path = data['圖片路徑']
                show_time = data['標註時間']

                if not description or not time or not image_path:
                    QMessageBox.warning(self, "警告", "請填寫所有欄位並選擇圖片")
                    return

                if self.burn_disc_checkbox.isChecked():
                    original_image_save_path = os.path.join(
                        folder_name, f"{i + 1:02d}-{description}-{case_id}-{case_address}.jpg"
                    )
                    original_image = Image.open(image_path)
                    if original_image.mode == 'RGBA':
                        original_image = original_image.convert('RGB')
                    original_image.save(original_image_save_path)

                if show_time:
                    try:
                        dpi = 96
                        width_px = int(10.3 * dpi / 2.54)
                        height_px = int(5.4 * dpi / 2.54)
                        image = Image.open(image_path)
                        image = image.resize((width_px, height_px), Image.LANCZOS)
                        if image.mode == 'RGBA':
                            image = image.convert('RGB')
                        draw = ImageDraw.Draw(image)
                        font_size = 36
                        try:
                            font = ImageFont.truetype("arial.ttf", font_size)
                        except Exception:
                            font = ImageFont.load_default()
                        bbox = draw.textbbox((0, 0), time, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                        text_position = (image.width - text_width - 20, image.height - text_height - 20)
                        draw.text(text_position, time, font=font, fill="red")
                        image_bytes = BytesIO()
                        image.save(image_bytes, format='JPEG')
                        image_bytes.seek(0)
                        self.image_bytes_list.append(image_bytes)
                        inline_image = InlineImage(self.doc, image_bytes, width=Cm(10.3), height=Cm(5.4))
                    except Exception as e:
                        QMessageBox.critical(self, "錯誤", f"在圖片上標註時間時出錯：{e}")
                        return

                else:
                    inline_image = InlineImage(self.doc, image_path, width=Cm(10.3), height=Cm(5.4))

                item_context = {
                    '案件編號': case_id,
                    '內容': description,
                    '時間': time,
                    '圖片': inline_image
                }
                processed_items.append(item_context)

            context = {'items': processed_items}
            output_path = f"{case_id}.docx"
            self.doc.render(context)
            self.doc.save(output_path)
            QMessageBox.information(self, "成功", f"文檔已生成：{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"生成文檔時出錯：{e}")

    def load_saved_projects(self):
        if os.path.exists(self.saved_data_file):
            try:
                with open(self.saved_data_file, 'r', encoding='utf-8') as file:
                    saved_data = json.load(file)
                    for project_name in saved_data.keys():
                        self.project_selector.addItem(project_name)
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"加載保存的項目時出錯：{e}")

    def load_selected_project(self):
        selected_index = self.project_selector.currentIndex()
        if selected_index == 0:
            return

        try:
            with open(self.saved_data_file, 'r', encoding='utf-8') as file:
                saved_data = json.load(file)
                project_name = self.project_selector.currentText()
                if project_name in saved_data:
                    project_data = saved_data[project_name]
                    self.case_id_input.setText(project_data['案件編號'])
                    self.case_address_input.setText(project_data['案件地址'])
                    self.clear_form_items()
                    for item_data in project_data['items']:
                        self.add_form_item(item_data)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"加載項目時出錯：{e}")

    def clear_form_items(self):
        self.item_list.clear()

    def save_current_project(self):
        case_id = self.case_id_input.text()
        case_address = self.case_address_input.text()
        if not case_id or not case_address:
            return

        project_name = f"{case_id}-{case_address}"
        project_data = {
            '案件編號': case_id,
            '案件地址': case_address,
            'items': []
        }

        for i in range(self.item_list.count()):
            item = self.item_list.item(i)
            widget = self.item_list.itemWidget(item)
            project_data['items'].append(widget.get_data())

        try:
            if os.path.exists(self.saved_data_file):
                with open(self.saved_data_file, 'r', encoding='utf-8') as file:
                    saved_data = json.load(file)
            else:
                saved_data = {}
            saved_data[project_name] = project_data
            with open(self.saved_data_file, 'w', encoding='utf-8') as file:
                json.dump(saved_data, file, indent=4, ensure_ascii=False)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"保存項目時出錯：{e}")

    def remove_project(self):
        selected_index = self.project_selector.currentIndex()
        if selected_index == 0:
            QMessageBox.warning(self, "警告", "請選擇要移除的案子")
            return

        project_name = self.project_selector.currentText()
        try:
            if os.path.exists(self.saved_data_file):
                with open(self.saved_data_file, 'r', encoding='utf-8') as file:
                    saved_data = json.load(file)
                if project_name in saved_data:
                    del saved_data[project_name]
                    with open(self.saved_data_file, 'w', encoding='utf-8') as file:
                        json.dump(saved_data, file, indent=4, ensure_ascii=False)
                    self.project_selector.removeItem(selected_index)
                    QMessageBox.information(self, "成功", f"案子 '{project_name}' 已被移除")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"移除項目時出錯：{e}")

    def closeEvent(self, event):
        self.save_current_project()
        for image_bytes in self.image_bytes_list:
            image_bytes.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ConstructionApp()
    window.show()
    sys.exit(app.exec_())
