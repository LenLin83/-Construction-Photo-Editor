# base_construction.py
import sys, os, json
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QLabel, QLineEdit,
    QListWidget, QPushButton, QMessageBox, QAbstractItemView, QCheckBox,
    QListWidgetItem
)
from PyQt5.QtCore import QSettings, Qt, QPoint
from PyQt5.QtGui import QPixmap
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

from widgets import FormItemWidget

class BaseConstructionApp(QWidget):
    def __init__(self, template_path, saved_data_file, info_fields):
        """
        :param template_path: 模板檔案路徑
        :param saved_data_file: 儲存資料的檔案名稱（僅作為區分用，不直接存檔於檔案中）
        :param info_fields: 字典，定義基本資訊欄位，例如：
                            {"id": "案件編號", "address": "案件地址"} 或 {"id": "站別", "address": "施工日期"}
        """
        super().__init__()
        self.template_path = template_path
        self.saved_data_file = saved_data_file  # 此參數僅用於區分不同版本
        self.info_fields = info_fields
        self.doc = DocxTemplate(self.template_path)
        self.image_bytes_list = []  # 用來保存圖片 BytesIO 物件
        self.init_ui()
        self.load_saved_projects()

    def get_output_filename(self, id_value, address_value):
        # 預設邏輯（案件版）：輸出檔名為 "{id_value}.docx"
        return f"{id_value}.docx"

    def get_photo_dimensions(self):
        # 預設尺寸 (案件版)：寬 6.5 cm, 高 10 cm
        return (10, 6.5)

    def init_ui(self):
        self.setGeometry(100, 100, 1200, 800)
        self.setWindowTitle("施工照片生成器")
        layout = QVBoxLayout(self)
        
        # 基本資訊區：根據 info_fields 設定欄位標籤
        info_layout = QHBoxLayout()
        self.project_selector = QComboBox()
        default_text = "選擇或創建新案子" if self.info_fields["id"] == "案件編號" else "選擇或創建新站"
        self.project_selector.addItem(default_text)
        self.project_selector.currentIndexChanged.connect(self.load_selected_project)
        info_layout.addWidget(self.project_selector)
        
        self.id_label = QLabel(f"{self.info_fields['id']}：")
        self.id_label.setStyleSheet("font-weight: bold;")
        self.id_input = QLineEdit()
        info_layout.addWidget(self.id_label)
        info_layout.addWidget(self.id_input)
        
        self.address_label = QLabel(f"{self.info_fields['address']}：")
        self.address_label.setStyleSheet("font-weight: bold;")
        self.address_input = QLineEdit()
        info_layout.addWidget(self.address_label)
        info_layout.addWidget(self.address_input)
        
        layout.addLayout(info_layout)
        
        # 列表顯示施工項目
        self.item_list = QListWidget()
        self.item_list.setDragDropMode(QAbstractItemView.InternalMove)
        self.item_list.setStyleSheet("""
            QListWidget::item:hover {
                background-color: #CCE5FF;
                border-radius: 15px;
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #A8D4FF;
                border-radius: 15px;
                padding: 5px;
            }
        """)
        layout.addWidget(self.item_list)
        
        # 按鈕區：新增、刪除、移除專案（移除下拉選單的專案功能）
        btn_layout = QHBoxLayout()
        self.add_item_button = QPushButton("新增內容項目")
        self.add_item_button.setStyleSheet("""
            QPushButton {
                background-color: #CCF1FF;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #B3E5FF; }
            QPushButton:pressed { background-color: #99D6FF; }
        """)
        self.add_item_button.clicked.connect(self.add_form_item)
        btn_layout.addWidget(self.add_item_button)
        
        self.delete_selected_button = QPushButton("刪除選取項目")
        self.delete_selected_button.setStyleSheet("""
            QPushButton {
                background-color: #FCFFCC;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #F0FFB2; }
            QPushButton:pressed { background-color: #E6FF99; }
        """)
        self.delete_selected_button.clicked.connect(self.delete_selected_items)
        btn_layout.addWidget(self.delete_selected_button)
        
        self.remove_project_button = QPushButton("移除案子" if self.info_fields["id"]=="案件編號" else "移除站")
        self.remove_project_button.setStyleSheet("""
            QPushButton {
                background-color: #FFCCCC;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #FFB3B3; }
            QPushButton:pressed { background-color: #FF9999; }
        """)
        self.remove_project_button.clicked.connect(self.remove_project)
        btn_layout.addWidget(self.remove_project_button)
        layout.addLayout(btn_layout)
        
        # 燒光碟選項與生成文檔按鈕
        self.burn_disc_checkbox = QCheckBox("是否燒光碟")
        layout.addWidget(self.burn_disc_checkbox)
        
        self.generate_button = QPushButton("生成文檔")
        self.generate_button.setStyleSheet("""
            QPushButton {
                background-color: #9EFFB5;
                border: none;
                border-radius: 10px;
                padding: 10px;
                font-size: 25px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #7DFF99; }
            QPushButton:pressed { background-color: #66FF80; }
        """)
        self.generate_button.clicked.connect(self.generate_document)
        layout.addWidget(self.generate_button)
        
        self.setLayout(layout)

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
        from docxtpl import InlineImage
        id_value = self.id_input.text()
        address_value = self.address_input.text()
        if not id_value or not address_value:
            QMessageBox.warning(self, "警告", f"請輸入{self.info_fields['id']}和{self.info_fields['address']}")
            return
        if self.item_list.count() == 0:
            QMessageBox.warning(self, "警告", "請至少添加一組內容")
            return
        try:
            processed_items = []
            folder_name = f"照片-{id_value}-{address_value}"
            if self.burn_disc_checkbox.isChecked() and not os.path.exists(folder_name):
                os.makedirs(folder_name)
            width_val, height_val = self.get_photo_dimensions()
            for i in range(self.item_list.count()):
                item = self.item_list.item(i)
                widget = self.item_list.itemWidget(item)
                data = widget.get_data()
                description = data['施工說明']
                time_val = data['時間']
                image_path = data['圖片路徑']
                show_time = data['標註時間']
                if not description or not time_val or not image_path:
                    QMessageBox.warning(self, "警告", "請填寫所有欄位並選擇圖片")
                    return
                if self.burn_disc_checkbox.isChecked():
                    save_path = os.path.join(folder_name, f"{i+1:02d}-{description}-{id_value}-{address_value}.jpg")
                    original_image = Image.open(image_path)
                    if original_image.mode == 'RGBA':
                        original_image = original_image.convert('RGB')
                    original_image.save(save_path)
                if show_time:
                    try:
                        dpi = 1024
                        width_px = int(width_val * dpi / 2.54)
                        height_px = int(height_val * dpi / 2.54)
                        image = Image.open(image_path)
                        image = image.resize((width_px, height_px), Image.LANCZOS)
                        if image.mode == 'RGBA':
                            image = image.convert('RGB')
                        draw = ImageDraw.Draw(image)
                        font_size = dpi // 6
                        try:
                            font = ImageFont.truetype("arial.ttf", font_size)
                        except Exception:
                            font = ImageFont.load_default()
                        bbox = draw.textbbox((0, 0), time_val, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                        text_position = (image.width - text_width - 50, image.height - text_height - 120)
                        draw.text(text_position, time_val, font=font, fill="red")
                        image_bytes = BytesIO()
                        image.save(image_bytes, format='JPEG')
                        image_bytes.seek(0)
                        self.image_bytes_list.append(image_bytes)
                        inline_image = InlineImage(self.doc, image_bytes, width=Cm(width_val), height=Cm(height_val))
                    except Exception as e:
                        QMessageBox.critical(self, "錯誤", f"在圖片上標註時間時出錯：{e}")
                        return
                else:
                    inline_image = InlineImage(self.doc, image_path, width=Cm(width_val), height=Cm(height_val))
                item_context = {
                    self.info_fields["id"]: id_value,
                    '內容': description,
                    '時間': time_val,
                    '圖片': inline_image
                }
                processed_items.append(item_context)
            context = {'items': processed_items}
            output_path = self.get_output_filename(id_value, address_value)
            self.doc.render(context)
            self.doc.save(output_path)
            QMessageBox.information(self, "成功", f"文檔已生成：{output_path}")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"生成文檔時出錯：{e}")
    
    def clear_form_items(self):
        self.item_list.clear()

    
    # 以下使用 QSettings 存取專案資料
    def save_current_project(self):
        from PyQt5.QtCore import QSettings
        import json
        
        settings = QSettings("MyCompany", f"ConstructionPhotoEditor_{self.info_fields['id']}")
        
        id_value = self.id_input.text().strip()
        address_value = self.address_input.text().strip()
        
        # 若沒輸入文字，就不進行儲存
        if not id_value or not address_value:
            # print("未輸入任何文字，不儲存")
            return
        
        project_name = f"{id_value}-{address_value}"
        project_data = {
            self.info_fields["id"]: id_value,
            self.info_fields["address"]: address_value,
            'items': []
        }
        
        # 收集 item_list 中每個項目的資料
        for i in range(self.item_list.count()):
            item = self.item_list.item(i)
            widget = self.item_list.itemWidget(item)
            project_data['items'].append(widget.get_data())
        
        # 寫入 QSettings
        settings.setValue(project_name, json.dumps(project_data, ensure_ascii=False, indent=4))
        print(f"Project saved: {project_name}")
    
    def load_saved_projects(self):
        from PyQt5.QtCore import QSettings
        settings = QSettings("MyCompany", f"ConstructionPhotoEditor_{self.info_fields['id']}")
        keys = settings.allKeys()
        for key in keys:
            self.project_selector.addItem(key)
    
    def load_selected_project(self):
        from PyQt5.QtCore import QSettings
        import json
        settings = QSettings("MyCompany", f"ConstructionPhotoEditor_{self.info_fields['id']}")
        selected_index = self.project_selector.currentIndex()
        if selected_index == 0:
            return
        project_name = self.project_selector.currentText()
        project_json = settings.value(project_name)
        if project_json:
            try:
                project_data = json.loads(project_json)
                self.id_input.setText(project_data[self.info_fields["id"]])
                self.address_input.setText(project_data[self.info_fields["address"]])
                self.clear_form_items()
                for item_data in project_data['items']:
                    self.add_form_item(item_data)
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"加載項目時出錯：{e}")
    
    def remove_project(self):
        from PyQt5.QtCore import QSettings
        settings = QSettings("MyCompany", f"ConstructionPhotoEditor_{self.info_fields['id']}")
        selected_index = self.project_selector.currentIndex()
        if selected_index == 0:
            QMessageBox.warning(self, "警告", f"請選擇要移除的{self.info_fields['id']}")
            return
        project_name = self.project_selector.currentText()
        settings.remove(project_name)
        self.project_selector.removeItem(selected_index)
        QMessageBox.information(self, "成功", f"{self.info_fields['id']} '{project_name}' 已被移除")
    
    def closeEvent(self, event):
        # 不自動儲存
        for image_bytes in self.image_bytes_list:
            image_bytes.close()
        event.accept()
