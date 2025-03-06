# MRT_project.py
from base_construction import BaseConstructionApp

class ConstructionAppStation(BaseConstructionApp):
    def __init__(self):
        super().__init__(
            template_path="照片.docx",
            saved_data_file="saved_data_station.json",
            info_fields={"id": "站別", "address": "施工日期"}
        )

    def get_output_filename(self, id_value, address_value):
        # 輸出格式為 "施工日期(站別).docx"
        return f"{address_value}({id_value}).docx"
    
    def get_photo_dimensions(self):
        # 捷運版尺寸：寬 5.4 cm, 高 8.3 cm
        return (8.3, 5.4)