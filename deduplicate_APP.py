import sys
from functools import partial
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import load_workbook
from PySide6.QtCore import QRunnable, Qt, QThreadPool, Signal
from PySide6.QtGui import QAction, QDragEnterEvent, QDropEvent, QScreen
from PySide6.QtWidgets import (QApplication, QCheckBox, QHBoxLayout, QLabel,
                               QMenu, QMessageBox, QProgressBar, QPushButton,
                               QTreeWidget, QTreeWidgetItem, QVBoxLayout,
                               QWidget)


class MyTask(QRunnable):
    def __init__(self, function, *args):
        super().__init__()

    def run(self):
        """
        重写 run 方法
        :return: 无
        """


class MainWidget(QWidget):
    def __init__(self):
        """
        初始化
        :return: 无
        """
        super().__init__()

        self.file_infos = {}  # 保存文件信息
        self.checkbox_infos = {}  # 保存 QCheckBox 的信息
        self.progressBar_value = 0  # 进度条的值

        self.max_cols = 0  # 最大列数

        # 线程池
        self.thread_pool = QThreadPool()
        self.thread_pool.setMaxThreadCount(4)

        self.initUI()

    def initUI(self):
        """
        初始化UI
        :return: 无
        """
        self.setAcceptDrops(True)

        layout = QVBoxLayout()
        checkbox_layout = QHBoxLayout()

        self.label = QLabel("将 Excel 文件或文件夹拖放到框内:", self)
        self.treeWidget = QTreeWidget()
        self.btn_clear = QPushButton('清空列表')

        self.label_tips = QLabel('请选择要检测重复的列号:')

        self.widget = QWidget()

        for i in range(1, 14):
            # 新建 QCheckBox
            checkBox = QCheckBox()
            checkBox.setText(f'第{i}列')

            # checkBox 状态改变信号处理
            checkBox.stateChanged.connect(partial(self.on_checkBox_state_changed, index=i))

            # 将 checkBox 添加到水平布局
            checkbox_layout.addWidget(checkBox)

            # 将编号与 QCheckBox 对应起来
            self.checkbox_infos[i] = {
                'obj': checkBox,
                'checkState': False
            }

        self.widget.setLayout(checkbox_layout)

        self.btn_cancel = QPushButton('停止执行')
        self.btn = QPushButton('开始去重')

        self.label_result = QLabel('...')
        self.progressBar = QProgressBar()

        self.contex_menu = QMenu(self)  # 创建菜单

        layout.addWidget(self.label)
        layout.addWidget(self.treeWidget)
        layout.addWidget(self.btn_clear)

        layout.addWidget(self.label_tips)
        layout.addWidget(self.widget)

        layout.addWidget(self.btn_cancel)
        layout.addWidget(self.btn)
        layout.addWidget(self.label_result)

        layout.addWidget(self.progressBar)

        self.setLayout(layout)

        # 为 TreeWidget 添加列表头
        self.treeWidget.setHeaderLabels(['文件名', '是否有重复项', '总行数', '总列数', '重复数据行数'])
        self.treeWidget.header().setDefaultAlignment(Qt.AlignCenter)  # 居中显示

        # 设置 TreeWidget 可拖拽
        self.treeWidget.setAcceptDrops(True)
        self.treeWidget.setDragEnabled(True)

        # 为右键菜单添加移除选项
        self.remove_action = QAction("移除选项", self, triggered=self.removeItem)
        self.contex_menu.addAction(self.remove_action)

        # 为 QTreeWidget 绑定右键菜单
        self.treeWidget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.treeWidget.customContextMenuRequested.connect(self.showContexMenu)

        self.btn_clear.clicked.connect(self.clearList)  # 清除列表的按键点击信号绑定
        self.btn_cancel.clicked.connect(self.cancelProcess_excel)
        self.btn.clicked.connect(self.process_excel)  # 开始去重按钮点击信号绑定

        self.setWindowTitle('Excel 去重软件')  # 窗体标题
        self.setWindowOpacity(0.95)

        self.setGeometry(800, 500, 890, 550)  # x, y, w, h

        self.centerOnScreen()  # 居中显示窗体

    def centerOnScreen(self):
        """
        将窗体移动到屏幕中央
        :return: 无
        """

    def dragEnterEvent(self, event: QDragEnterEvent):
        """
        拖拽进入事件
        :param event: 事件
        :return: 无
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        """
        拖拽事件
        :param event: 事件
        :return: 无
        """

    def addFilePath_to_TreeView(self, filepath: str):
        """
        添加文件路径到 TreeView 中
        :param filepath: 文件路径
        :return: 无
        """

    def showContexMenu(self, point):
        """
        将右键菜单显示在屏幕中
        :param point: 坐标
        :return: 无
        """

    def getExcelFileRowsCols(self, filepath) -> Tuple:
        """
        获取 Excel 文件的总行数
        :param filepath: 文件路径
        :return: 无
        """

    def updateTreeView(self, filepath: str, only_rowcol: bool):
        """
        更新 QTreeWidget
        :param filepath: 文件路径
        :param max_rows: 文件总行数
        :return: 无
        """

    def updateProgressBar(self):
        """
        更新进度条
        :return: 无
        """

    def removeItem(self):
        """
        移除 QTreeWidget 中的项目
        :return: 无
        """

    def setLable(self, label: QLabel, text, color='#000'):
        """
        设定 QLabel 的文本和样式
        :param label: QLabel 对象
        :param text: label 上要显示的文本
        :param color: label 文字上的颜色
        :return: 无
        """

    def on_checkBox_state_changed(self, checked, index):
        """
        QCheckBox 状态改变时要执行的槽函数
        :param checked: > 0 为选中; 其余为未选中
        :param index: 编号
        :return: 无
        """

    def getAllCheckBoxState(self) -> Dict:
        """
        获取所有 checkBox 的状态, 判断用户是否勾选了 checkBox
        :return:  返回一个字典. {1: checked, 1: checked, ...}
        """

    def clearList(self):
        """
        清空 文件列表
        :return: 无
        """

    def cancelProcess_excel(self):
        """
        取消处理 Excel 文件
        :return: 无
        """

    def process_excel(self):
        """
        处理 Excel 文件
        :return: 无
        """

    def deduplicate_excel(self, filepath: Path, cols: List[int]) -> bool:
        """
        Excel 文件去重的核心函数
        :param filepath: Excel 文件的路径
        :param cols: 要用哪几列来检查重复
        :return: 返回 bool 值. 如果有重复的就返回 True, 没有重复数据就返回 False
        """


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWidget()
    window.show()
    sys.exit(app.exec())
