import sys
import os
import json
import re
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox, QTextEdit, QCheckBox, QScrollArea
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QFormLayout, QFrame, QComboBox
from PyQt5.QtGui import QPixmap, QCursor
from PyQt5.QtCore import Qt, QEvent, QPoint
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO

class ConstructionApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.items = []  # 用於存放多組輸入的內容
        self.image_bytes_list = []  # 用於保存 BytesIO 對象，便於釋放內存
        self.zoom_label = None  # 用於顯示放大圖片的 QLabel
        self.saved_data_file = "saved_data.json"  # 保存數據的文件名
        self.saved_photos_folder = "saved_photos"  # 保存圖片的資料夾

        if hasattr(sys, '_MEIPASS'):
            # 當程式被打包成 EXE 時，使用 sys._MEIPASS 來獲取臨時目錄
            self.template_path = os.path.join(sys._MEIPASS, "施工照片.docx")
        else:
            # 當程式作為腳本運行時，使用用戶手動選擇的範本文件
            self.template_path = "施工照片.docx"

        # 加載範本
        self.doc = DocxTemplate(self.template_path)

        # 加載保存的項目
        self.load_saved_projects()

    def init_ui(self):
        # 設置視窗標題和大小
        self.setWindowTitle("施工照片生成器")
        self.setGeometry(100, 100, 1200, 800)  # 調整視窗大小以適應預覽圖片

        # 主佈局
        self.main_layout = QVBoxLayout(self)

        # 案件選擇和基本信息輸入
        project_info_layout = QHBoxLayout()

        self.project_selector = QComboBox()
        self.project_selector.addItem("選擇或創建新案子")
        self.project_selector.currentIndexChanged.connect(self.load_selected_project)
        project_info_layout.addWidget(self.project_selector)

        self.case_id_label = QLabel("案件編號：")
        self.case_id_input = QLineEdit()
        project_info_layout.addWidget(self.case_id_label)
        project_info_layout.addWidget(self.case_id_input)

        self.case_address_label = QLabel("案件地址：")
        self.case_address_input = QLineEdit()
        project_info_layout.addWidget(self.case_address_label)
        project_info_layout.addWidget(self.case_address_input)

        self.main_layout.addLayout(project_info_layout)

        # # 添加 "選擇範本文件" 按鈕
        # self.browse_template_button = QPushButton("選擇範本文件")
        # self.browse_template_button.clicked.connect(self.browse_template)
        # self.main_layout.addWidget(self.browse_template_button)

        # 添加動態表單的區域
        self.form_area = QVBoxLayout()

        # 滾動區域，當有多個表單時可滾動查看
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        form_content = QWidget()
        form_content.setLayout(self.form_area)
        scroll.setWidget(form_content)
        self.main_layout.addWidget(scroll)

        # “新增內容項目”按鈕和“刪除項目”按鈕
        add_delete_layout = QHBoxLayout()
        self.add_item_button = QPushButton("新增內容項目")
        self.add_item_button.clicked.connect(self.add_form_item)
        add_delete_layout.addWidget(self.add_item_button)

        self.delete_item_button = QPushButton("刪除項目")
        self.delete_item_button.clicked.connect(self.delete_form_item)
        add_delete_layout.addWidget(self.delete_item_button)

        self.remove_project_button = QPushButton("移除案子")
        self.remove_project_button.clicked.connect(self.remove_project)
        add_delete_layout.addWidget(self.remove_project_button)

        self.main_layout.addLayout(add_delete_layout)

        # 增加燒光碟選項
        self.burn_disc_checkbox = QCheckBox("是否燒光碟")
        self.main_layout.addWidget(self.burn_disc_checkbox)

        # 生成文档按鈕
        self.generate_button = QPushButton("生成文檔")
        self.generate_button.clicked.connect(self.generate_document)
        self.main_layout.addWidget(self.generate_button)

        self.setLayout(self.main_layout)

    def add_form_item(self, item_data=None):
        # 每次添加一組新的表單項目
        form_layout = QHBoxLayout()  # 使用 QHBoxLayout 將每組表單橫向排列
        form_widget = QFrame()  # 使用 QFrame 來封裝每組項目，方便刪除
        form_widget.setLayout(form_layout)

        # 左側表單輸入區域
        input_layout = QFormLayout()

        # 施工說明
        description_input = QTextEdit()
        description_input.setFixedHeight(80)  # 設定高度，使得與預覽圖片齊平
        input_layout.addRow(QLabel("施工說明："), description_input)

        # 時間
        time_input = QLineEdit()
        input_layout.addRow(QLabel("時間："), time_input)

        # 圖片選擇
        image_path_input = QLineEdit()
        image_browse_button = QPushButton("瀏覽")
        image_preview_label = QLabel("圖片預覽")
        image_preview_label.setFixedSize(500, 300)
        image_preview_label.setAlignment(Qt.AlignCenter)
        image_preview_label.setStyleSheet("border: 1px solid black;")

        if item_data:
            description_input.setPlainText(item_data['施工說明'])
            time_input.setText(item_data['時間'])
            image_path_input.setText(item_data['圖片路徑'])
            original_pixmap = QPixmap(item_data['圖片路徑']).scaled(500, 300, Qt.KeepAspectRatio)
            image_preview_label.setPixmap(original_pixmap)
            image_preview_label.setProperty("original_pixmap", QPixmap(item_data['圖片路徑']))

        image_browse_button.clicked.connect(lambda: self.browse_image(image_path_input, image_preview_label))
        image_layout = QHBoxLayout()
        image_layout.addWidget(image_path_input)
        image_layout.addWidget(image_browse_button)
        input_layout.addRow(QLabel("圖片："), image_layout)

        # 是否標註時間
        time_checkbox = QCheckBox("是否標註時間")
        input_layout.addRow(time_checkbox)
        if item_data:
            time_checkbox.setChecked(item_data['標註時間'])

        # 將左側輸入佈局添加到表單佈局中
        form_layout.addLayout(input_layout)

        # 右側圖片預覽
        form_layout.addWidget(image_preview_label, alignment=Qt.AlignTop)  # 將圖片預覽對齊到頂部

        # 事件：滑鼠懸停時顯示放大圖片
        image_preview_label.installEventFilter(self)
        self.current_image_preview_label = None  # 用於追蹤當前懸停的圖片預覽

        # 將整組表單放到主表單區域
        self.form_area.addWidget(form_widget)

        # 保存每個表單控件，方便後續處理
        form_item = {
            'widget': form_widget,  # 保存 QFrame 以便刪除
            '施工說明': description_input,
            '時間': time_input,
            '圖片路徑': image_path_input,
            '圖片預覽': image_preview_label,
            '標註時間': time_checkbox
        }
        self.items.append(form_item)

    def delete_form_item(self):
        # 刪除最後一個添加的內容項目
        if self.items:
            form_item = self.items.pop()
            form_item['widget'].deleteLater()  # 刪除 UI 控件
        else:
            QMessageBox.warning(self, "警告", "沒有更多的項目可以刪除")

    def browse_image(self, image_path_input, image_preview_label):
        # 打開文件選擇對話框
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "選擇圖片文件", "", "Images (*.png *.jpg *.jpeg *.bmp);;All Files (*)", options=options)
        if file_path:
            image_path_input.setText(file_path)
            original_pixmap = QPixmap(file_path).scaled(500, 300, Qt.KeepAspectRatio)
            image_preview_label.setPixmap(original_pixmap)
            image_preview_label.setProperty("original_pixmap", QPixmap(file_path))  # 保存原始解析度的圖片

    def eventFilter(self, source, event):
        if isinstance(source, QLabel) and source.property("original_pixmap") is not None:
            if event.type() == QEvent.Enter:
                # 當滑鼠進入圖片範圍時顯示放大圖片
                original_pixmap = source.property("original_pixmap")
                if original_pixmap is not None:
                    if self.zoom_label is None:
                        self.zoom_label = QLabel(self)
                        self.zoom_label.setWindowFlags(Qt.ToolTip)
                    self.zoom_label.setPixmap(original_pixmap.scaled(1600, 900, Qt.KeepAspectRatio))
                    cursor_pos = QCursor.pos()
                    self.zoom_label.move(cursor_pos + QPoint(20, 20))  # 放大圖片顯示在滑鼠右下角
                    self.zoom_label.show()
            elif event.type() == QEvent.Leave:
                # 當滑鼠離開圖片範圍時隱藏放大圖片
                if self.zoom_label is not None:
                    self.zoom_label.hide()
        return super().eventFilter(source, event)

    # def browse_template(self):
    #     # 打開文件選擇對話框，讓用戶選擇要使用的範本文件
    #     options = QFileDialog.Options()
    #     file_path, _ = QFileDialog.getOpenFileName(self, "選擇範本文件", "", "Word Documents (*.docx);;All Files (*)", options=options)
    #     if file_path:
    #         self.template_path = file_path
    #         QMessageBox.information(self, "範本文件選擇", f"已選擇範本文件：{file_path}")

    def generate_document(self):
        # 確認是否有添加任何內容
        if not self.items:
            QMessageBox.warning(self, "警告", "請至少添加一組內容")
            return

        # 獲取案件編號和地址
        case_id = self.case_id_input.text()
        case_address = self.case_address_input.text()
        if not case_id or not case_address:
            QMessageBox.warning(self, "警告", "請輸入案件編號和地址")
            return

        # 使用 docxtpl 生成 Word 文檔
        try:
            processed_items = []

            # 創建資料夾，如果燒光碟選項已勾選
            folder_name = f"照片-{case_id}-{case_address}"
            if self.burn_disc_checkbox.isChecked():
                if not os.path.exists(folder_name):
                    os.makedirs(folder_name)

            for index, form in enumerate(self.items):
                # 獲取每組表單的資料
                description = form['施工說明'].toPlainText()
                time = form['時間'].text()
                image_path = form['圖片路徑'].text()
                show_time = form['標註時間'].isChecked()

                # 檢查必填欄位
                if not description or not time or not image_path:
                    QMessageBox.warning(self, "警告", "請填寫所有欄位並選擇圖片")
                    return

                # 保存原始圖片至資料夾
                if self.burn_disc_checkbox.isChecked():
                    original_image_save_path = os.path.join(folder_name, f"{index + 1:02d}-{description}-{case_id}-{case_address}.jpg")
                    original_image = Image.open(image_path)
                    if original_image.mode == 'RGBA':
                        original_image = original_image.convert('RGB')
                    original_image.save(original_image_save_path)

                # 如果選擇顯示時間，則在圖片上標註時間
                if show_time:
                    try:
                        # 打開原始圖片並將其調整為固定尺寸（10.3 x 5.4 厘米）
                        dpi = 96  # 假設 DPI 為 96
                        width_px = int(10.3 * dpi / 2.54)  # 10.3 厘米轉換為像素
                        height_px = int(5.4 * dpi / 2.54)  # 5.4 厘米轉換為像素

                        image = Image.open(image_path)
                        # 使用 LANCZOS 濾鏡調整圖片大小（替代 ANTIALIAS）
                        image = image.resize((width_px, height_px), Image.LANCZOS)

                        # 如果圖片是 RGBA 模式，轉換為 RGB 模式
                        if image.mode == 'RGBA':
                            image = image.convert('RGB')

                        # 在圖片上標註時間
                        draw = ImageDraw.Draw(image)

                        # 設置字體（注意這裡需要本地有 TrueType 字體文件 .ttf）
                        font_size = 36  # 設置固定的字體大小
                        font = ImageFont.truetype("arial.ttf", font_size)

                        # 使用 textbbox 來計算文本邊界框，獲得文本的寬和高
                        bbox = draw.textbbox((0, 0), time, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]

                        # 設置文本的位置（右下角），顏色為紅色
                        text_position = (image.width - text_width - 20, image.height - text_height - 20)  # 距離右下角 20 像素
                        draw.text(text_position, time, font=font, fill="red")

                        # 將修改後的圖片保存在記憶體中，而不是保存到文件
                        image_bytes = BytesIO()
                        image.save(image_bytes, format='JPEG')
                        image_bytes.seek(0)
                        self.image_bytes_list.append(image_bytes)  # 保存 BytesIO 對象以便後續釋放

                        # 將修改後的圖片添加到 items 列表中
                        inline_image = InlineImage(self.doc, image_bytes, width=Cm(10.3))

                    except Exception as e:
                        QMessageBox.critical(self, "錯誤", f"在圖片上標註時間時出錯：{e}")
                        return
                else:
                    # 未選擇標註時間的情況下，直接使用原始圖片
                    inline_image = InlineImage(self.doc, image_path, width=Cm(10.3))

                # 將這組資料添加到處理列表
                item = {
                    '案件編號': case_id,
                    '內容': description,
                    '時間': time,
                    '圖片': inline_image
                }
                processed_items.append(item)

            # 準備上下文
            context = {'items': processed_items}

            # 渲染範本並保存
            output_path = f"{case_id}.docx"
            self.doc.render(context)
            self.doc.save(output_path)

            QMessageBox.information(self, "成功", f"文檔已生成：{output_path}")

        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"生成文檔時出錯：{e}")

    def closeEvent(self, event):
        # 在關閉窗口時釋放記憶體中的 BytesIO 對象
        for image_bytes in self.image_bytes_list:
            image_bytes.close()
        event.accept()

    def load_saved_projects(self):
        # 加載保存的項目，並更新到下拉列表中
        if os.path.exists(self.saved_data_file):
            try:
                with open(self.saved_data_file, 'r') as file:
                    saved_data = json.load(file)
                    for project_name in saved_data.keys():
                        self.project_selector.addItem(project_name)
            except Exception as e:
                QMessageBox.critical(self, "錯誤", f"加載保存的項目時出錯：{e}")

    def load_selected_project(self):
        # 加載選擇的項目，根據 project_selector 的選擇
        selected_index = self.project_selector.currentIndex()
        if selected_index == 0:
            return

        # 從保存的數據中加載項目
        try:
            with open(self.saved_data_file, 'r') as file:
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
        # 清除所有的表單項目
        while self.items:
            form_item = self.items.pop()
            form_item['widget'].deleteLater()

    def save_current_project(self):
        # 保存當前項目的數據
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

        for form in self.items:
            item_data = {
                '施工說明': form['施工說明'].toPlainText(),
                '時間': form['時間'].text(),
                '圖片路徑': form['圖片路徑'].text(),
                '標註時間': form['標註時間'].isChecked()
            }
            project_data['items'].append(item_data)

        # 將項目數據保存到文件中
        try:
            if os.path.exists(self.saved_data_file):
                with open(self.saved_data_file, 'r') as file:
                    saved_data = json.load(file)
            else:
                saved_data = {}

            saved_data[project_name] = project_data

            with open(self.saved_data_file, 'w') as file:
                json.dump(saved_data, file, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"保存項目時出錯：{e}")

    def remove_project(self):
        # 移除選定的項目
        selected_index = self.project_selector.currentIndex()
        if selected_index == 0:
            QMessageBox.warning(self, "警告", "請選擇要移除的案子")
            return

        project_name = self.project_selector.currentText()
        try:
            if os.path.exists(self.saved_data_file):
                with open(self.saved_data_file, 'r') as file:
                    saved_data = json.load(file)

                if project_name in saved_data:
                    del saved_data[project_name]

                    with open(self.saved_data_file, 'w') as file:
                        json.dump(saved_data, file, indent=4)

                    self.project_selector.removeItem(selected_index)
                    QMessageBox.information(self, "成功", f"案子 '{project_name}' 已被移除")
        except Exception as e:
            QMessageBox.critical(self, "錯誤", f"移除項目時出錯：{e}")

    def closeEvent(self, event):
        # 在關閉窗口時保存當前項目
        self.save_current_project()
        for image_bytes in self.image_bytes_list:
            image_bytes.close()
        event.accept()

# 主程序
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ConstructionApp()
    window.show()
    sys.exit(app.exec_())
