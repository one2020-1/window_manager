# Windows窗口管理工具

一个基于PyQt6开发的Windows窗口管理工具，提供直观的图形界面和便捷的窗口布局功能。

## 功能特点

### 窗口管理
- 📋 显示所有可见窗口列表
- ✅ 支持多选窗口进行批量操作
- 🔄 实时刷新窗口列表
- 🖱️ 点击列表项可直接激活对应窗口

### 布局选项
- 📊 垂直排列：将选中窗口垂直排列
- 📚 层叠排列：将选中窗口以层叠方式排列
- 🎯 网格排列：将选中窗口以网格方式排列

### 窗口控制
- 👁️ 显示/隐藏选中窗口
- ⚡ 一键排列（快捷键：Ctrl+Alt+Z）
- 📌 窗口置顶功能

### 界面特性
- 🎨 现代化UI设计
- 🖼️ 无边框窗口设计
- ✨ 美观的动画和悬停效果
- 📱 支持调整窗口大小
- 🔝 始终保持在最上层

## 系统要求
- Windows操作系统
- Python 3.6+

## 安装依赖
```bash
pip install -r requirements.txt
```

## 依赖项
- PyQt6：用于创建图形用户界面
- pywin32：用于Windows系统交互

## 使用方法
1. 运行程序：
```bash
python window_manager.py


2. 使用步骤：
   - 在窗口列表中选择要管理的窗口
   - 选择所需的布局方式（垂直/层叠/网格）
   - 点击"一键排列"或使用快捷键(Ctrl+Alt+Z)进行排列
   - 可以随时显示/隐藏选中的窗口

## 特别说明
- 程序会保持在桌面最上层，方便随时调用
- 支持通过鼠标拖拽调整工具窗口位置和大小
- 提供窗口刷新功能，确保列表始终最新
