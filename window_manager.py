import sys
import win32gui
import win32con
import win32api
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                           QHBoxLayout, QPushButton, QCheckBox, 
                           QLabel, QFrame, QScrollArea, QButtonGroup, QRadioButton,
                           QSizeGrip)
from PyQt6.QtCore import Qt, QSize, QPoint, QRect
from PyQt6.QtGui import QIcon, QShortcut, QKeySequence, QPalette, QColor, QFont, QMouseEvent

STYLE_SHEET = """
QMainWindow {
    background-color: #ffffff;
}

QWidget#centralWidget {
    background-color: #ffffff;
    border: 1px solid #e1e4e8;
    border-radius: 12px;
}

QWidget#titleBar {
    background-color: #f8f9fa;
    border-top-left-radius: 12px;
    border-top-right-radius: 12px;
    border-bottom: 1px solid #e1e4e8;
}

QPushButton {
    background-color: #ffffff;
    border: 1px solid #d0d7de;
    border-radius: 6px;
    padding: 8px 16px;
    color: #24292f;
    font-weight: 500;
    min-width: 80px;
}

QPushButton:hover {
    background-color: #f3f4f6;
    border-color: #1a7f37;
    color: #1a7f37;
}

QPushButton:pressed {
    background-color: #e7e8ea;
    border-color: #1a7f37;
}

QPushButton#arrangeButton {
    background-color: #2da44e;
    color: white;
    border: none;
    padding: 12px 20px;
    font-weight: bold;
    font-size: 14px;
}

QPushButton#arrangeButton:hover {
    background-color: #2c974b;
}

QPushButton#arrangeButton:pressed {
    background-color: #298e46;
}

QPushButton#hideButton {
    background-color: #0969da;
    color: white;
    border: none;
    padding: 12px 20px;
    font-weight: bold;
    font-size: 14px;
}

QPushButton#hideButton:hover {
    background-color: #0552b5;
}

QPushButton#hideButton:pressed {
    background-color: #033d8b;
}

QPushButton#refreshButton {
    background-color: transparent;
    border: 1px solid #d0d7de;
    color: #57606a;
    padding: 8px 12px;
}

QPushButton#refreshButton:hover {
    background-color: #f6f8fa;
    color: #24292f;
    border-color: #1a7f37;
}

QPushButton#showButton {
    background-color: #0969da;
    color: white;
    border: none;
    padding: 12px;
    font-weight: bold;
    font-size: 14px;
}

QPushButton#showButton:hover {
    background-color: #0552b5;
}

QPushButton#showButton:pressed {
    background-color: #033d8b;
}

QCheckBox {
    spacing: 8px;
    color: #24292f;
    padding: 8px;
    border-radius: 6px;
    font-size: 13px;
}

QCheckBox:hover {
    background-color: #f6f8fa;
}

QCheckBox::indicator {
    width: 22px;
    height: 22px;
    border: 2px solid #d0d7de;
    border-radius: 6px;
    background-color: white;
}

QCheckBox::indicator:hover {
    border-color: #2da44e;
    background-color: #f6f8fa;
}

QCheckBox::indicator:checked {
    background-color: #2da44e;
    border-color: #2da44e;
    image: url(check.png);
}

QCheckBox::indicator:checked:hover {
    background-color: #2c974b;
    border-color: #2c974b;
}

QRadioButton {
    spacing: 8px;
    padding: 8px 12px;
    color: #24292f;
    border-radius: 6px;
    font-size: 13px;
}

QRadioButton:hover {
    background-color: #f6f8fa;
}

QRadioButton::indicator {
    width: 20px;
    height: 20px;
    border: 2px solid #d0d7de;
    border-radius: 10px;
    background-color: white;
}

QRadioButton::indicator:checked {
    background-color: #2da44e;
    border: 6px solid #2da44e;
}

QRadioButton::indicator:hover {
    border-color: #2da44e;
}

QFrame {
    background-color: white;
    border: 1px solid #d0d7de;
    border-radius: 8px;
}

QFrame#windowListFrame {
    background-color: white;
    padding: 16px;
    margin: 8px;
}

QFrame#layoutFrame {
    padding: 20px;
    margin: 8px;
    background-color: #f6f8fa;
}

QScrollArea {
    border: none;
    background-color: transparent;
}

QScrollBar:vertical {
    border: none;
    background-color: #f6f8fa;
    width: 8px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background-color: #d0d7de;
    border-radius: 4px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #afb8c1;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QLabel {
    color: #24292f;
}

QLabel#headerLabel {
    font-size: 16px;
    font-weight: bold;
    color: #24292f;
    padding: 4px;
}

QLabel#titleLabel {
    font-size: 14px;
    font-weight: bold;
    color: #24292f;
}

QSizeGrip {
    background-color: transparent;
    width: 16px;
    height: 16px;
}

/* 窗口边缘拉伸区域样式 */
QWidget#resizeLeft, QWidget#resizeRight {
    background: transparent;
    width: 5px;
}

QWidget#resizeTop, QWidget#resizeBottom {
    background: transparent;
    height: 5px;
}

QWidget#resizeTopLeft, QWidget#resizeTopRight, 
QWidget#resizeBottomLeft, QWidget#resizeBottomRight {
    background: transparent;
    width: 10px;
    height: 10px;
}
"""

class WindowListItem(QWidget):
    """窗口列表项"""
    def __init__(self, hwnd, title):
        super().__init__()
        self.hwnd = hwnd
        self.title = title
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        
        self.checkbox = QCheckBox(title)
        self.checkbox.setChecked(True)
        layout.addWidget(self.checkbox)
        
        self.setAutoFillBackground(True)
        self.normalPalette = self.palette()
        self.hoverPalette = QPalette(self.normalPalette)
        self.hoverPalette.setColor(QPalette.ColorRole.Window, QColor("#f6f8fa"))
        
        # 添加鼠标点击事件
        self.mousePressEvent = self.on_mouse_press
    
    def on_mouse_press(self, event):
        """处理鼠标点击事件"""
        if event.button() == Qt.MouseButton.LeftButton:
            # 切换复选框状态
            self.checkbox.setChecked(not self.checkbox.isChecked())
            event.accept()
    
    def enterEvent(self, event):
        self.setPalette(self.hoverPalette)
        
    def leaveEvent(self, event):
        self.setPalette(self.normalPalette)

class ResizeHandle(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.edge = None
        self.setMouseTracking(True)
        
    def setEdge(self, edge):
        """设置拉伸方向"""
        self.edge = edge
        # 根据拉伸方向设置鼠标样式
        if edge in ["left", "right"]:
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        elif edge in ["top", "bottom"]:
            self.setCursor(Qt.CursorShape.SizeVerCursor)
        elif edge in ["topleft", "bottomright"]:
            self.setCursor(Qt.CursorShape.SizeFDiagCursor)
        elif edge in ["topright", "bottomleft"]:
            self.setCursor(Qt.CursorShape.SizeBDiagCursor)
        
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.parent.resizing = True
            self.parent.resizeStartPos = event.globalPosition().toPoint()
            self.parent.resizeStartGeometry = self.parent.geometry()
            event.accept()
            
    def mouseMoveEvent(self, event):
        if hasattr(self.parent, 'resizing') and self.parent.resizing:
            delta = event.globalPosition().toPoint() - self.parent.resizeStartPos
            newGeometry = QRect(self.parent.resizeStartGeometry)
            
            # 获取屏幕尺寸
            screen = QApplication.primaryScreen()
            screenGeometry = screen.availableGeometry()
            
            # 限制最小和最大尺寸
            MIN_WIDTH = 360
            MIN_HEIGHT = 400
            MAX_WIDTH = screenGeometry.width()
            MAX_HEIGHT = screenGeometry.height()
            
            if self.edge in ["left", "topleft", "bottomleft"]:
                newWidth = self.parent.resizeStartGeometry.width() - delta.x()
                if MIN_WIDTH <= newWidth <= MAX_WIDTH:
                    newGeometry.setLeft(self.parent.resizeStartGeometry.left() + delta.x())
            
            if self.edge in ["right", "topright", "bottomright"]:
                newWidth = self.parent.resizeStartGeometry.width() + delta.x()
                if MIN_WIDTH <= newWidth <= MAX_WIDTH:
                    newGeometry.setRight(self.parent.resizeStartGeometry.right() + delta.x())
            
            if self.edge in ["top", "topleft", "topright"]:
                newHeight = self.parent.resizeStartGeometry.height() - delta.y()
                if MIN_HEIGHT <= newHeight <= MAX_HEIGHT:
                    newGeometry.setTop(self.parent.resizeStartGeometry.top() + delta.y())
            
            if self.edge in ["bottom", "bottomleft", "bottomright"]:
                newHeight = self.parent.resizeStartGeometry.height() + delta.y()
                if MIN_HEIGHT <= newHeight <= MAX_HEIGHT:
                    newGeometry.setBottom(self.parent.resizeStartGeometry.bottom() + delta.y())
            
            # 确保窗口不会超出屏幕范围
            if newGeometry.left() < screenGeometry.left():
                newGeometry.moveLeft(screenGeometry.left())
            if newGeometry.right() > screenGeometry.right():
                newGeometry.moveRight(screenGeometry.right())
            if newGeometry.top() < screenGeometry.top():
                newGeometry.moveTop(screenGeometry.top())
            if newGeometry.bottom() > screenGeometry.bottom():
                newGeometry.moveBottom(screenGeometry.bottom())
            
            self.parent.setGeometry(newGeometry)
            event.accept()
            
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.parent.resizing = False
            event.accept()

class WindowManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("窗口管理工具")
        self.setMinimumSize(360, 600)
        self.setStyleSheet(STYLE_SHEET)
        
        # 设置窗口无边框和保持在最上层
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        
        # Store our window handle
        self.hwnd = None
        self.windowCreated = False
        self.dragging = False
        self.dragPos = QPoint()
        self.resizing = False
        self.resizeStartPos = QPoint()
        self.resizeStartGeometry = self.geometry()
        self.updating_checkboxes = False  # 防止循环触发
        
        # Create main widget and layout
        main_widget = QWidget()
        main_widget.setObjectName("centralWidget")
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(16, 16, 16, 16)
        
        # 添加拉伸控件
        self.setupResizeHandles()
        
        # 添加标题栏
        title_bar = QWidget()
        title_bar.setObjectName("titleBar")
        title_bar_layout = QHBoxLayout(title_bar)
        title_bar_layout.setContentsMargins(8, 8, 8, 8)
        
        title_label = QLabel("窗口管理工具")
        title_label.setObjectName("titleLabel")
        title_bar_layout.addWidget(title_label)
        
        # 添加最小化和关闭按钮
        title_bar_layout.addStretch()
        
        minimize_btn = QPushButton("—")
        minimize_btn.setObjectName("minimizeButton")
        minimize_btn.clicked.connect(self.showMinimized)
        title_bar_layout.addWidget(minimize_btn)
        
        close_btn = QPushButton("×")
        close_btn.setObjectName("closeButton")
        close_btn.clicked.connect(self.close)
        title_bar_layout.addWidget(close_btn)
        
        main_layout.addWidget(title_bar)
        
        # Create window list frame
        window_list_frame = QFrame()
        window_list_frame.setObjectName("windowListFrame")
        window_list_layout = QVBoxLayout(window_list_frame)
        
        # Add header with select all checkbox
        header_layout = QHBoxLayout()
        self.select_all_checkbox = QCheckBox("全选")
        self.select_all_checkbox.setObjectName("headerLabel")
        self.select_all_checkbox.clicked.connect(self.toggle_all_windows)  # 使用clicked而不是stateChanged
        header_layout.addWidget(self.select_all_checkbox)
        
        refresh_btn = QPushButton("刷新")
        refresh_btn.setObjectName("refreshButton")
        refresh_btn.clicked.connect(self.refresh_window_list)
        header_layout.addWidget(refresh_btn)
        
        window_list_layout.addLayout(header_layout)
        
        # Create scroll area for window list
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        scroll_content = QWidget()
        self.window_list_layout = QVBoxLayout(scroll_content)
        self.window_list_layout.setSpacing(2)
        self.window_list_layout.addStretch()
        
        scroll_area.setWidget(scroll_content)
        window_list_layout.addWidget(scroll_area)
        
        main_layout.addWidget(window_list_frame)
        
        # Create layout options
        layout_frame = QFrame()
        layout_frame.setObjectName("layoutFrame")
        layout_frame_layout = QVBoxLayout(layout_frame)
        
        layout_label = QLabel("布局选项")
        layout_label.setObjectName("headerLabel")
        layout_frame_layout.addWidget(layout_label)
        
        # Add radio buttons for layout options
        self.layout_group = QButtonGroup()
        
        vertical_radio = QRadioButton("垂直排列")
        vertical_radio.setChecked(True)
        self.layout_group.addButton(vertical_radio)
        layout_frame_layout.addWidget(vertical_radio)
        
        cascade_radio = QRadioButton("层叠排列")
        self.layout_group.addButton(cascade_radio)
        layout_frame_layout.addWidget(cascade_radio)
        
        tile_radio = QRadioButton("网格排列")
        self.layout_group.addButton(tile_radio)
        layout_frame_layout.addWidget(tile_radio)
        
        main_layout.addWidget(layout_frame)
        
        # Create arrange button
        button_layout = QHBoxLayout()
        
        show_btn = QPushButton("显示窗口")
        show_btn.setObjectName("showButton")
        show_btn.clicked.connect(self.show_windows)
        button_layout.addWidget(show_btn)
        
        hide_btn = QPushButton("隐藏窗口")
        hide_btn.setObjectName("hideButton")
        hide_btn.clicked.connect(self.hide_windows)
        button_layout.addWidget(hide_btn)
        
        arrange_btn = QPushButton("一键排列 (Ctrl+Alt+Z)")
        arrange_btn.setObjectName("arrangeButton")
        arrange_btn.clicked.connect(self.arrange_windows)
        button_layout.addWidget(arrange_btn)
        
        main_layout.addLayout(button_layout)
        
        # Add shortcut
        self.shortcut = QShortcut(QKeySequence("Ctrl+Alt+Z"), self)
        self.shortcut.activated.connect(self.arrange_windows)
        
        # Add size grip
        size_grip = QSizeGrip(self)
        size_grip.setFixedSize(16, 16)
        
        # Create bottom layout for size grip
        bottom_layout = QHBoxLayout()
        bottom_layout.addStretch()
        bottom_layout.addWidget(size_grip)
        main_layout.addLayout(bottom_layout)
        
        # Initialize window list
        self.refresh_window_list()

    def setupResizeHandles(self):
        # 添加拉伸控件
        resize_left = ResizeHandle(self)
        resize_left.setObjectName("resizeLeft")
        resize_left.setFixedSize(5, self.height())
        resize_left.move(0, 0)
        resize_left.setEdge("left")
        
        resize_right = ResizeHandle(self)
        resize_right.setObjectName("resizeRight")
        resize_right.setFixedSize(5, self.height())
        resize_right.move(self.width() - 5, 0)
        resize_right.setEdge("right")
        
        resize_top = ResizeHandle(self)
        resize_top.setObjectName("resizeTop")
        resize_top.setFixedSize(self.width(), 5)
        resize_top.move(0, 0)
        resize_top.setEdge("top")
        
        resize_bottom = ResizeHandle(self)
        resize_bottom.setObjectName("resizeBottom")
        resize_bottom.setFixedSize(self.width(), 5)
        resize_bottom.move(0, self.height() - 5)
        resize_bottom.setEdge("bottom")
        
        resize_top_left = ResizeHandle(self)
        resize_top_left.setObjectName("resizeTopLeft")
        resize_top_left.setFixedSize(10, 10)
        resize_top_left.move(0, 0)
        resize_top_left.setEdge("topleft")
        
        resize_top_right = ResizeHandle(self)
        resize_top_right.setObjectName("resizeTopRight")
        resize_top_right.setFixedSize(10, 10)
        resize_top_right.move(self.width() - 10, 0)
        resize_top_right.setEdge("topright")
        
        resize_bottom_left = ResizeHandle(self)
        resize_bottom_left.setObjectName("resizeBottomLeft")
        resize_bottom_left.setFixedSize(10, 10)
        resize_bottom_left.move(0, self.height() - 10)
        resize_bottom_left.setEdge("bottomleft")
        
        resize_bottom_right = ResizeHandle(self)
        resize_bottom_right.setObjectName("resizeBottomRight")
        resize_bottom_right.setFixedSize(10, 10)
        resize_bottom_right.move(self.width() - 10, self.height() - 10)
        resize_bottom_right.setEdge("bottomright")

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            self.dragging = True
            self.dragPos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
            
    def mouseReleaseEvent(self, event: QMouseEvent):
        self.dragging = False
        event.accept()
        
    def mouseMoveEvent(self, event: QMouseEvent):
        if self.dragging and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self.dragPos)
            event.accept()

    def get_window_list(self):
        """获取所有可见窗口的列表"""
        def enum_window_callback(hwnd, windows):
            """枚举窗口回调函数"""
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and hwnd != self.hwnd:  # 排除自身窗口
                    # 获取窗口样式
                    style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                    
                    # 检查是否为主应用窗口
                    if ((style & win32con.WS_VISIBLE) and  # 窗口可见
                        (style & win32con.WS_OVERLAPPEDWINDOW) and  # 是标准窗口
                        not (ex_style & win32con.WS_EX_TOOLWINDOW)):  # 不是工具窗口
                        
                        # 排除特定的系统窗口
                        class_name = win32gui.GetClassName(hwnd)
                        if class_name not in ['Shell_TrayWnd', 'Progman', 'WorkerW']:
                            try:
                                # 获取窗口位置
                                rect = win32gui.GetWindowRect(hwnd)
                                if rect[2] - rect[0] > 0 and rect[3] - rect[1] > 0:  # 确保窗口有大小
                                    windows.append((hwnd, title))
                            except:
                                pass
            return True

        windows = []
        try:
            # 枚举所有顶级窗口
            win32gui.EnumWindows(enum_window_callback, windows)
            # 按标题排序
            return sorted(windows, key=lambda x: x[1].lower())
        except Exception as e:
            print(f"Error enumerating windows: {str(e)}")
            return []

    def toggle_all_windows(self, checked):
        """切换所有窗口的选中状态"""
        try:
            # 遍历所有窗口项并设置状态
            for i in range(self.window_list_layout.count()):
                item = self.window_list_layout.itemAt(i)
                if item and isinstance(item.widget(), WindowListItem):
                    item.widget().checkbox.setChecked(checked)
        
            # 更新选中计数
            self.update_selected_count()
                    
        except Exception as e:
            print(f"Error toggling windows: {str(e)}")

    def closeEvent(self, event):
        """Restore stdout when closing"""
        super().closeEvent(event)

    def showEvent(self, event):
        """Capture our window handle when window is shown"""
        super().showEvent(event)
        if self.windowCreated and not self.hwnd:
            self.hwnd = win32gui.FindWindow(None, self.windowTitle())

    def get_screen_info(self):
        """Get information about the screen"""
        try:
            # 获取主显示器信息
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            return screen_width, screen_height
        except Exception as e:
            print(f"Error getting screen information: {str(e)}")
            return 1920, 1080  # 默认分辨率

    def refresh_window_list(self):
        """Refresh the list of windows"""
        try:
            # 清除现有列表
            while self.window_list_layout.count():
                item = self.window_list_layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            # 获取窗口列表
            windows = self.get_window_list()
            window_count = len(windows)
            
            # 创建标签栏
            header_widget = QWidget()
            header_layout = QHBoxLayout(header_widget)
            header_layout.setContentsMargins(8, 4, 8, 4)
            
            # 窗口计数标签
            count_label = QLabel(f"窗口总数: {window_count}")
            count_label.setStyleSheet("""
                QLabel {
                    color: #24292f;
                    font-size: 12px;
                    padding: 4px 8px;
                    background: #f6f8fa;
                    border-radius: 6px;
                    font-weight: bold;
                }
            """)
            header_layout.addWidget(count_label)
            
            # 添加分隔符
            separator = QLabel("|")
            separator.setStyleSheet("color: #d0d7de; margin: 0 8px;")
            header_layout.addWidget(separator)
            
            # 当前选中数量标签
            self.selected_count_label = QLabel(f"已选择: {window_count}")
            self.selected_count_label.setStyleSheet("""
                QLabel {
                    color: #24292f;
                    font-size: 12px;
                    padding: 4px 8px;
                    background: #ddf4ff;
                    border-radius: 6px;
                }
            """)
            header_layout.addWidget(self.selected_count_label)
            
            header_layout.addStretch()
            self.window_list_layout.addWidget(header_widget)
            
            # 添加分割线
            line = QFrame()
            line.setFrameShape(QFrame.Shape.HLine)
            line.setStyleSheet("background-color: #d0d7de;")
            self.window_list_layout.addWidget(line)
            
            # 添加窗口列表
            for hwnd, title in windows:
                window_item = WindowListItem(hwnd, title)
                window_item.checkbox.stateChanged.connect(self.update_selected_count)
                self.window_list_layout.addWidget(window_item)
            
            self.window_list_layout.addStretch()
            
            # 初始化全选状态
            self.select_all_checkbox.setChecked(True)
            self.update_selected_count()
            
        except Exception as e:
            print(f"Error refreshing window list: {str(e)}")

    def update_selected_count(self):
        """更新选中窗口数量"""
        try:
            selected_count = 0
            total_count = 0
        
            # 计算选中数量
            for i in range(self.window_list_layout.count()):
                item = self.window_list_layout.itemAt(i)
                if item and isinstance(item.widget(), WindowListItem):
                    total_count += 1
                    if item.widget().checkbox.isChecked():
                        selected_count += 1
        
            # 更新选中数量标签
            if hasattr(self, 'selected_count_label'):
                self.selected_count_label.setText(f"已选择: {selected_count}")
        
            # 更新全选复选框状态
            if hasattr(self, 'select_all_checkbox'):
                self.select_all_checkbox.setChecked(selected_count == total_count)
                
        except Exception as e:
            print(f"Error updating selected count: {str(e)}")

    def get_selected_windows(self):
        """Get list of selected windows"""
        selected_windows = []
        for i in range(self.window_list_layout.count()):
            item = self.window_list_layout.itemAt(i)
            if item and isinstance(item.widget(), WindowListItem):
                window_item = item.widget()
                if window_item.checkbox.isChecked():
                    selected_windows.append(window_item.hwnd)
        return selected_windows

    def arrange_windows(self):
        """Arrange the selected windows"""
        try:
            selected_windows = []
            for i in range(self.window_list_layout.count()):
                item = self.window_list_layout.itemAt(i)
                if item and isinstance(item.widget(), WindowListItem):
                    window_item = item.widget()
                    if window_item.checkbox.isChecked():
                        selected_windows.append(window_item.hwnd)

            if not selected_windows:
                return

            # Get screen dimensions
            screen_width, screen_height = self.get_screen_info()
            taskbar_height = 40
            margin = 8

            # 恢复并激活所有选中的窗口
            for hwnd in selected_windows:
                # 强制结束最小化状态
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # 移除最大化状态
                placement = win32gui.GetWindowPlacement(hwnd)
                if placement[1] == win32con.SW_SHOWMAXIMIZED:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # 将窗口带到前台
                win32gui.SetForegroundWindow(hwnd)

            if self.layout_group.checkedButton().text() == "垂直排列":  # 垂直平铺
                num_windows = len(selected_windows)
                if num_windows == 1:
                    # 单个窗口占据整个屏幕
                    win32gui.SetWindowPos(selected_windows[0], win32con.HWND_TOPMOST,
                                        margin, margin,
                                        screen_width - 2 * margin,
                                        screen_height - taskbar_height - 2 * margin,
                                        win32con.SWP_SHOWWINDOW)
                    win32gui.SetWindowPos(selected_windows[0], win32con.HWND_NOTOPMOST,
                                        0, 0, 0, 0,
                                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
                else:
                    # 计算每列窗口数
                    windows_per_col = (num_windows + 1) // 2
                    window_width = (screen_width - margin * 3) // 2
                    window_height = (screen_height - taskbar_height - margin * (windows_per_col + 1)) // windows_per_col

                    for i, hwnd in enumerate(selected_windows):
                        col = i % 2  # 0 for left column, 1 for right column
                        row = i // 2  # Row within the column
                        
                        x = margin + col * (window_width + margin)
                        y = margin + row * (window_height + margin)

                        # 设置窗口位置和大小
                        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                            x, y,
                                            window_width, window_height,
                                            win32con.SWP_SHOWWINDOW)
                        # 重置窗口Z序
                        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                            0, 0, 0, 0,
                                            win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)

            elif self.layout_group.checkedButton().text() == "层叠排列":  # 重叠平铺
                base_width = int(screen_width * 0.8)
                base_height = int((screen_height - taskbar_height) * 0.8)
                offset = 30

                for i, hwnd in enumerate(selected_windows):
                    x = margin + (i * offset)
                    y = margin + (i * offset)
                    
                    # 如果超出屏幕范围，重置位置
                    if x + base_width > screen_width - margin:
                        x = margin
                    if y + base_height > screen_height - taskbar_height - margin:
                        y = margin

                    # 设置窗口位置和大小
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                        x, y,
                                        base_width, base_height,
                                        win32con.SWP_SHOWWINDOW)
                    # 重置窗口Z序
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                        0, 0, 0, 0,
                                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)

            else:  # 网格排列
                num_windows = len(selected_windows)
                
                # 计算最佳网格大小
                cols = int(num_windows ** 0.5)
                if cols * cols < num_windows:
                    cols += 1
                rows = (num_windows + cols - 1) // cols
                
                # 计算单个窗口大小，考虑边距
                window_width = (screen_width - margin * (cols + 1)) // cols
                window_height = (screen_height - taskbar_height - margin * (rows + 1)) // rows

                for i, hwnd in enumerate(selected_windows):
                    col = i % cols
                    row = i // cols
                    
                    x = margin + col * (window_width + margin)
                    y = margin + row * (window_height + margin)

                    # 设置窗口位置和大小
                    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST,
                                        x, y,
                                        window_width, window_height,
                                        win32con.SWP_SHOWWINDOW)
                    # 重置窗口Z序
                    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST,
                                        0, 0, 0, 0,
                                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
        except Exception as e:
            print(f"Error arranging windows: {str(e)}")

    def show_windows(self):
        """显示选中的窗口"""
        try:
            for i in range(self.window_list_layout.count()):
                item = self.window_list_layout.itemAt(i)
                if item and isinstance(item.widget(), WindowListItem):
                    window_item = item.widget()
                    if window_item.checkbox.isChecked():
                        hwnd = window_item.hwnd
                        # 恢复窗口
                        if win32gui.IsIconic(hwnd):  # 如果窗口最小化
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        # 激活窗口
                        win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            print(f"Error showing windows: {str(e)}")

    def hide_windows(self):
        """隐藏选中的窗口"""
        try:
            for i in range(self.window_list_layout.count()):
                item = self.window_list_layout.itemAt(i)
                if item and isinstance(item.widget(), WindowListItem):
                    window_item = item.widget()
                    if window_item.checkbox.isChecked():
                        hwnd = window_item.hwnd
                        # 最小化窗口
                        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        except Exception as e:
            print(f"Error hiding windows: {str(e)}")

    def resizeEvent(self, event):
        """处理窗口大小改变事件"""
        super().resizeEvent(event)
        # 更新拉伸控件位置
        for child in self.children():
            if isinstance(child, ResizeHandle):
                if child.objectName() == "resizeLeft":
                    child.setFixedSize(5, self.height())
                    child.move(0, 0)
                elif child.objectName() == "resizeRight":
                    child.setFixedSize(5, self.height())
                    child.move(self.width() - 5, 0)
                elif child.objectName() == "resizeTop":
                    child.setFixedSize(self.width(), 5)
                    child.move(0, 0)
                elif child.objectName() == "resizeBottom":
                    child.setFixedSize(self.width(), 5)
                    child.move(0, self.height() - 5)
                elif child.objectName() == "resizeTopLeft":
                    child.move(0, 0)
                elif child.objectName() == "resizeTopRight":
                    child.move(self.width() - 10, 0)
                elif child.objectName() == "resizeBottomLeft":
                    child.move(0, self.height() - 10)
                elif child.objectName() == "resizeBottomRight":
                    child.move(self.width() - 10, self.height() - 10)


if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        window = WindowManager()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"Error running application: {str(e)}")
