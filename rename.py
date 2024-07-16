import sys

import os
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QLineEdit, QPushButton, QFileDialog, QTextEdit)

class RenameApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Batch Rename Files')
        self.setGeometry(300, 300, 1000, 600)
        # self.setStyleSheet("QWidget { background-color: #121212; }")
        # 创建布局
        layout = QVBoxLayout()

        # Excel文件路径输入和选择按钮
        excelHBoxLayout = QHBoxLayout()
        self.excel_label = QLabel('Excel File Path:', self)
        self.excel_path_input = QLineEdit(self)
        self.excel_browse_button = QPushButton('Browse', self)
        excelHBoxLayout.addWidget(self.excel_label)
        excelHBoxLayout.addWidget(self.excel_path_input)
        excelHBoxLayout.addWidget(self.excel_browse_button)

        # 文件存储路径输入和选择按钮
        directoryHBoxLayout = QHBoxLayout()
        self.directory_label = QLabel('Directory Path:', self)
        self.directory_path_input = QLineEdit(self)
        self.directory_browse_button = QPushButton('Browse', self)
        directoryHBoxLayout.addWidget(self.directory_label)
        directoryHBoxLayout.addWidget(self.directory_path_input)
        directoryHBoxLayout.addWidget(self.directory_browse_button)

        # 重命名按钮
        self.rename_button = QPushButton('Rename Files', self)
        self.rename_button.clicked.connect(self.rename_files)

        # 消息显示文本框
        self.message_text_edit = QTextEdit(self)
        self.message_text_edit.setReadOnly(True)
        layout.addWidget(self.message_text_edit)
        self.message_text_edit.setStyleSheet("QTextEdit { background-color: #121212; color: #FFFFFF; }")
        # 添加到主布局
        layout.addLayout(excelHBoxLayout)
        layout.addLayout(directoryHBoxLayout)
        layout.addWidget(self.rename_button)

        self.setLayout(layout)

        # 设置按钮的槽函数
        self.excel_browse_button.clicked.connect(self.select_excel_file)
        self.directory_browse_button.clicked.connect(self.select_directory)

    def select_excel_file(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "",
                                                  "Excel Files (*.xlsx *.xls);;All Files (*)", options=options)
        if fileName:
            self.excel_path_input.setText(fileName)

    def select_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.directory_path_input.setText(directory)

    def rename_files(self):
        excel_path = self.excel_path_input.text()
        directory_path = self.directory_path_input.text()

        # 显示正在处理的消息
        self.append_message(f'Processing files in directory "{directory_path}" using data from "{excel_path}"')

        # 这里是你的批量重命名逻辑
        # 例如，你可以在这里调用之前提供的函数
        # 确保在函数中打印消息，并使用append_message来更新界面
        # 这里调用你的批量重命名函数
        df = pd.read_excel(excel_path, header=None)
        
        # 获取所有列的索引
        columns_to_use = list(range(df.shape[1]))
        attributes = [df.iloc[:, i] for i in columns_to_use]
        
        # 获取指定目录下的所有文件
        current_files = os.listdir(directory_path)

        # 确保Excel文件中的行数与指定目录下的文件数量相同
        if len(attributes[0]) != len(current_files):
            self.append_message("Error: The number of rows in the Excel does not match the number of files in the directory.")
            return

        # 迭代并重命名文件
        for index, row in enumerate(zip(*attributes)):
            # 使用列索引作为文件名的一部分，假设每个属性之间用"-"连接，并添加.docx扩展名
            new_filename = "-".join(str(item) for item in row) + ".docx"
            old_file = current_files[index]
            old_file_path = os.path.join(directory_path, old_file)
            new_file_path = os.path.join(directory_path, new_filename)

            # 检查新文件名是否已经存在，如果不存在则重命名
            if not os.path.exists(new_file_path):
                os.rename(old_file_path, new_file_path)
                self.append_message(f'Renamed "{old_file}" to "{new_filename}"')
            else:
                self.append_message(f'Error: The file "{new_filename}" already exists. Skipping rename for "{old_file}".')
        # 假设重命名操作成功完成
        self.append_message('Renaming operation completed successfully.')

    def append_message(self, message):
        # 将消息添加到文本框中
        current_text = self.message_text_edit.toPlainText()
        self.message_text_edit.setPlainText(current_text + '\n' + message)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = RenameApp()
    ex.show()
    sys.exit(app.exec_())


