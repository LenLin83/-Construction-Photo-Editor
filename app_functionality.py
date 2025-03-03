# app_functionality.py
from base_construction import BaseConstructionApp

class ConstructionApp(BaseConstructionApp):
    def __init__(self):
        super().__init__(
            template_path="施工照片.docx",
            saved_data_file="saved_data.json",
            info_fields={"id": "案件編號", "address": "案件地址"}
        )
    # 使用基底類別預設：輸出檔名為 "{id_value}.docx"，尺寸為 (6.5,10) cm
