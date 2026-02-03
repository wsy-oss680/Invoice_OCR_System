import sys
import pandas as pd
import cv2  # 导入 OpenCV 库
import numpy as np
from PyQt5.QtWidgets import (QApplication, QWidget, QHBoxLayout, QVBoxLayout, 
                             QPushButton, QLabel, QFileDialog, QTableWidget, 
                             QTableWidgetItem, QHeaderView, QMessageBox, QFrame)
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import os
import PyQt5
import tempfile
# 这一行是核心，它会告诉程序去哪里找 windows 插件
os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins')

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
        self.ocr_engine = None  # OCR 引擎只初始化一次
        self.current_file = None  # 当前文件路径
        self.preprocessed_file = None  # 预处理后的文件路径
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
            self.preprocessed_file = None  # 重置预处理状态
            
            # 使用 OpenCV 读取以支持中文路径
            try:
                img_array = np.fromfile(fname, dtype=np.uint8)
                img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
                if img is None:
                    QMessageBox.critical(self, "错误", "无法读取图像文件，请检查文件是否损坏！")
                    return
                
                pixmap = QPixmap(fname)
                self.image_label.setPixmap(pixmap.scaled(self.image_label.width()-20, 
                                                       self.image_label.height()-20, 
                                                       Qt.KeepAspectRatio, 
                                                       Qt.SmoothTransformation))
            except Exception as e:
                QMessageBox.critical(self, "错误", f"加载图像失败: {str(e)}")

    def run_preprocess(self):
        if not self.current_file:
            QMessageBox.warning(self, "警告", "请先导入发票图像！")
            return
            
        try:
            # 使用 numpy 读取支持中文路径
            img_array = np.fromfile(self.current_file, dtype=np.uint8)
            img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            
            if img is None:
                QMessageBox.critical(self, "错误", "无法读取图像文件！")
                return
            
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            # 使用自适应阈值，效果更好
            binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                                          cv2.THRESH_BINARY, 11, 2)
            
            # 使用临时文件，避免路径问题
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.jpg')
            self.preprocessed_file = temp_file.name
            temp_file.close()
            
            cv2.imwrite(self.preprocessed_file, binary)
            pixmap = QPixmap(self.preprocessed_file)
            self.image_label.setPixmap(pixmap.scaled(self.image_label.width()-20, 
                                                   self.image_label.height()-20, 
                                                   Qt.KeepAspectRatio, 
                                                   Qt.SmoothTransformation))
            QMessageBox.information(self, "预处理完成", "图像已使用自适应阈值优化。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"预处理失败: {str(e)}")

    def mock_ocr(self):
        if not self.current_file:
            QMessageBox.warning(self, "警告", "请先导入发票图像！")
            return
            
        # 禁用按钮，防止重复点击
        self.btn_ocr.setEnabled(False)
        self.btn_ocr.setText("🔄 识别中...")
        QApplication.processEvents()  # 刷新界面
        
        try:
            # 初始化 OCR 引擎（只初始化一次）
            if self.ocr_engine is None:
                from paddleocr import PaddleOCR
                # 设置环境变量，避免版本兼容性问题
                os.environ['FLAGS_use_mkldnn'] = '0'
                self.ocr_engine = PaddleOCR(lang="ch", use_angle_cls=True)
            
            # 优先使用预处理后的图像
            target_file = self.preprocessed_file if self.preprocessed_file else self.current_file
            
            # 执行识别
            result = self.ocr_engine.ocr(target_file, cls=True)
            
            # 检查结果是否为空
            if not result or not result[0]:
                QMessageBox.warning(self, "提示", "未识别到文字内容，请尝试预处理图像后再识别。")
                return
            
            # 提取识别出的所有文字
            raw_text = ""
            all_texts = []
            for line in result[0]:
                text = line[1][0]
                raw_text += text + " "
                all_texts.append(text)
            
            # 简单的字段提取逻辑（基于关键词匹配）
            self.extract_fields(all_texts)
            
            QMessageBox.information(self, "识别成功", 
                                   f"成功识别 {len(all_texts)} 个文字区域！\n\n完整文本:\n{raw_text[:100]}...")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"识别失败: {str(e)}")
        finally:
            # 恢复按钮状态
            self.btn_ocr.setEnabled(True)
            self.btn_ocr.setText("🚀 3. 开始智能识别")
    
    def extract_fields(self, texts):
        """从识别文本中提取发票字段"""
        import re
        
        # 将所有文本合并用于搜索
        full_text = " ".join(texts)
        
        # 发票代码：10位数字（广东增值税发票格式）
        invoice_code = ""
        for text in texts:
            # 匹配10位数字
            if re.match(r'^\d{10}$', text):
                invoice_code = text
                break
        
        # 发票号码：8位数字（No 后面的数字）
        invoice_number = ""
        for i, text in enumerate(texts):
            # 查找 "No" 关键词后面的数字
            if 'No' in text or 'NO' in text or '号码' in text:
                # 尝试从后续文本中查找8位数字
                for j in range(i, min(i+3, len(texts))):
                    match = re.search(r'\d{8}', texts[j])
                    if match:
                        invoice_number = match.group()
                        break
                if invoice_number:
                    break
        
        # 如果上面未找到，直接查找8位数字
        if not invoice_number:
            for text in texts:
                if re.match(r'^\d{8}$', text) and text != invoice_code:
                    invoice_number = text
                    break
        
        # 开票日期：查找多种日期格式
        date_pattern = re.compile(r'\d{4}年\d{1,2}月\d{1,2}日|\d{4}-\d{2}-\d{2}|\d{4}/\d{2}/\d{2}|\d{4}\.\d{2}\.\d{2}')
        invoice_date = ""
        for text in texts:
            match = date_pattern.search(text)
            if match:
                invoice_date = match.group()
                break
        
        # 合计金额：查找多种金额格式
        amount_pattern = re.compile(r'¥\s*\d+[,\d]*\.\d{2}|\d+[,\d]*\.\d{2}|\d+元\d+角\d+分')
        total_amount = ""
        max_amount = 0.0
        
        # 查找最大的金额（通常价税合计是最大的）
        for text in texts:
            matches = amount_pattern.findall(text)
            for match in matches:
                # 提取数值
                num_str = re.sub(r'[¥,元角分\s]', '', match)
                try:
                    amount = float(num_str)
                    if amount > max_amount:
                        max_amount = amount
                        total_amount = match
                except:
                    continue
        
        # 校验码/密码区：查找包含特殊字符的长字符串
        check_code = ""
        for text in texts:
            # 密码区通常包含 <> * + - / 等符号
            if len(text) > 20 and re.search(r'[<>*+\-/]', text):
                check_code = text
                break
        
        # 如果未找到密码区，查找20位数字/字母组合
        if not check_code:
            for text in texts:
                if re.match(r'^[0-9A-Z]{20,}$', text) or re.match(r'^\d{20,}$', text):
                    check_code = text
                    break
        
        # 填充表格
        self.result_table.setItem(0, 1, QTableWidgetItem(invoice_code or "未识别"))
        self.result_table.setItem(1, 1, QTableWidgetItem(invoice_number or "未识别"))
        self.result_table.setItem(2, 1, QTableWidgetItem(invoice_date or "未识别"))
        self.result_table.setItem(3, 1, QTableWidgetItem(total_amount or "未识别"))
        self.result_table.setItem(4, 1, QTableWidgetItem(check_code or "未识别"))

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
