# main.py
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTabWidget

from app_functionality import ConstructionApp       # 案件版
from MRT_project import ConstructionAppStation        # 捷運版

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        self.setGeometry(100, 100, 1200, 1000)
        self.tabs = QTabWidget()
        
        
        self.case_tab = ConstructionApp()
        self.tabs.addTab(self.case_tab, "案件版")
        
        self.mrt_tab = ConstructionAppStation()
        self.tabs.addTab(self.mrt_tab, "捷運版")
        
        layout.addWidget(self.tabs)
        self.setLayout(layout)
        self.setWindowTitle("施工照片生成器")
    
    def closeEvent(self, event):
        current_index = self.tabs.currentIndex()
        # 0 表示第一個分頁（案件版），1 表示第二個分頁（捷運版）
        if current_index == 0:
            # 儲存案件版
            try:
                self.case_tab.save_current_project()
            except Exception as e:
                print("案件版存檔錯誤：", e)
        else:
            # 儲存捷運版
            try:
                self.mrt_tab.save_current_project()
            except Exception as e:
                print("捷運版存檔錯誤：", e)
    
        event.accept()

if __name__ == '__main__':
    from PyQt5.QtWidgets import QApplication
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(lambda: print("Application is about to quit", flush=True))
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
