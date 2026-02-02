import sys
import pandas as pd
import cv2  # 导入 OpenCV 库
from PyQt5.QtWidgets import (QApplication, QWidget, QHBoxLayout, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QMessageBox, QFrame)
from PyQt5.QtGui import QPixmap, QFont
from PyQt5.QtCore import Qt

# --- 定义高级感 QSS 样式表 ---
STYLESHEET = """
QWidget {
    background-color: #2b2b2b;
    color: #ffffff;
    font-family: "Segoe UI", "Microsoft YaHei";
}

QLabel {
    font-size: 14px;
    color: #dcdcdc;
}

QPushButton {
    background-color: #3d3d3d;
    border: 1px solid #555555;
    border-radius: 5px;
    padding: 10px;
    font-size: 13px;
    min-height: 20px;
}

QPushButton:hover {
    background-color: #505050;
    border: 1px solid #00aaff;
}

QPushButton#primary_btn {
    background-color: #007acc;
    font-weight: bold;
    border: none;
}

QPushButton#primary_btn:hover {
    background-color: #0098ff;
}

QTableWidget {
    background-color: #1e1e1e;
    alternate-background-color: #252525;
    gridline-color: #404040;
    border: 1px solid #404040;
    border-radius: 5px;
}

QHeaderView::section {
    background-color: #333333;
    padding: 4px;
    border: 1px solid #404040;
    font-weight: bold;
}

QTableWidgetItem {
    padding: 5px;
}
"""

class InvoiceSystem(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setStyleSheet(STYLESHEET) # 应用美化样式
        
    def initUI(self):
        # 1. 窗口基本属性
        self.setWindowTitle('发票智能录入系统 - 智屏版')
        self.resize(1200, 800)
        
        main_layout = QHBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)
        
        # --- 左侧布局：图像处理模块 ---
        left_panel = QFrame()
        left_layout = QVBoxLayout(left_panel)
        
        title_left = QLabel("📷 图像输入与处理")
        title_left.setStyleSheet("font-size: 18px; font-weight: bold; color: #00aaff; margin-bottom: 10px;")
        left_layout.addWidget(title_left)

        self.image_label = QLabel('请点击下方按钮上传发票原图')
        self.image_label.setAlignment(Qt.AlignCenter) 
        self.image_label.setStyleSheet("""
            border: 2px dashed #555555; 
            border-radius: 10px;
            background: #1e1e1e;
        """)
        left_layout.addWidget(self.image_label, 1) 
        
        self.btn_load = QPushButton('📂 1. 导入发票图像')
        self.btn_load.clicked.connect(self.load_image)
        left_layout.addWidget(self.btn_load)

        self.btn_pre = QPushButton('✨ 2. 图像预处理 (OpenCV 增强)')
        self.btn_pre.clicked.connect(self.run_preprocess)
        left_layout.addWidget(self.btn_pre)
        
        # --- 右侧布局：结果与控制 ---
        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        
        title_right = QLabel("📊 结构化识别结果")
        title_right.setStyleSheet("font-size: 18px; font-weight: bold; color: #00ffaa; margin-bottom: 10px;")
        right_layout.addWidget(title_right)
        
        self.result_table = QTableWidget(5, 2)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setHorizontalHeaderLabels(['字段名称', '识别内容'])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        self.fields = ['发票代码', '发票号码', '开票日期', '合计金额', '校验码']
        for i, field in enumerate(self.fields):
            self.result_table.setItem(i, 0, QTableWidgetItem(field))
            self.result_table.setItem(i, 1, QTableWidgetItem("等待识别..."))
            
        right_layout.addWidget(self.result_table)
        
        self.btn_ocr = QPushButton('🚀 3. 开始智能识别')
        self.btn_ocr.setObjectName("primary_btn") # 使用特殊蓝色样式
        self.btn_ocr.clicked.connect(self.mock_ocr)
        right_layout.addWidget(self.btn_ocr)

        self.btn_export = QPushButton('📥 4. 导出识别结果 (Excel)')
        self.btn_export.clicked.connect(self.export_to_excel)
        right_layout.addWidget(self.btn_export)
        
        main_layout.addWidget(left_panel, 2) 
        main_layout.addWidget(right_panel, 1) 
        self.setLayout(main_layout)
        
    def load_image(self):
        fname, _ = QFileDialog.getOpenFileName(self, '选择发票图片', '', '图片文件 (*.jpg *.png *.jpeg)')
        if fname:
            self.current_file = fname
            pixmap = QPixmap(fname)
            self.image_label.setPixmap(pixmap.scaled(self.image_label.width()-20, 
                                                   self.image_label.height()-20, 
                                                   Qt.KeepAspectRatio, 
                                                   Qt.SmoothTransformation))

    def run_preprocess(self):
        if hasattr(self, 'current_file'):
            img = cv2.imread(self.current_file)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            # 对应开题报告中的二值化处理
            _, binary = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            preview_path = "preprocessed_temp.jpg"
            cv2.imwrite(preview_path, binary)
            pixmap = QPixmap(preview_path)
            self.image_label.setPixmap(pixmap.scaled(self.image_label.width()-20, 
                                                   self.image_label.height()-20, 
                                                   Qt.KeepAspectRatio, 
                                                   Qt.SmoothTransformation))
            QMessageBox.information(self, "预处理完成", "图像已优化。")
        else:
            QMessageBox.warning(self, "警告", "请先导入发票图像！")

    def mock_ocr(self):
        results = ["011002200111", "88776655", "2026-02-02", "￥520.00", "12345678901234567890"]
        for i, val in enumerate(results):
            self.result_table.setItem(i, 1, QTableWidgetItem(val))
        QMessageBox.information(self, "识别完成", "OCR 处理已结束。")

    def export_to_excel(self):
        data = {self.result_table.item(i, 0).text(): [self.result_table.item(i, 1).text()] 
                for i in range(self.result_table.rowCount())}
        try:
            df = pd.DataFrame(data)
            save_path, _ = QFileDialog.getSaveFileName(self, "保存 Excel", "发票数据.xlsx", "Excel Files (*.xlsx)")
            if save_path:
                df.to_excel(save_path, index=False)
                QMessageBox.information(self, "成功", "文件保存成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")     

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = InvoiceSystem()
    ex.show()
    sys.exit(app.exec_())