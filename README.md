# 时迹 · TimeTrace

> 一个悬浮在桌面的轻量级工作计时器

![GitHub release (latest by date)](https://img.shields.io/github/v/release/jsmbjsm/TimeTrace)
![GitHub](https://img.shields.io/github/license/jsmbjsm/TimeTrace)

## 简介

**时迹** 是一个专为「不想看表、不想被打扰」的用户设计的桌面计时工具。

- 🪟 **悬浮置顶**：始终停留在屏幕最上层，半透明不遮挡视线
- ⏱️ **专注计时**：正计时 / 暂停 / 结束，操作一气呵成
- 📋 **任务管理**：新建、选择、删除任务名，随时切换
- 💾 **一键导出 Excel**：自动记录任务名称、起止时间、总时长
- 🔇 **静默运行**：无弹窗、无提示音，仅有时间标签的视觉反馈
- 🖱️ **系统托盘**：最小化到托盘，点击图标切换显隐
- 🔒 **单实例运行**：双击 exe 只会启动一个实例

## 截图

<img width="650" height="111" alt="e5485cc5bb46214b2a96c5a9856d5bee" src="https://github.com/user-attachments/assets/e3235e4d-4bbc-4667-91ca-82dc12b32e14" />


## 快速开始

### 直接运行源码（需 Python 环境）

```bash
pip install PyQt5 openpyxl
python ShiJi.py
```

### 打包成 exe（无需 Python 环境）

```bash
pip install pyinstaller
pyinstaller -F -w -i shiji.ico --add-data "shiji.ico;." --name "ShiJi" ShiJi.py
```

打包完成后，`dist/ShiJi.exe` 即为最终产物，**双击即可运行**。

## 使用指南

| 操作 | 效果 |
|------|------|
| 点击「选择任务」 | 弹出对话框，选择已有任务或输入新任务名 |
| 点击绿色 ▶ | 开始计时（按钮变为橙色 ⏸） |
| 点击橙色 ⏸ | 暂停计时（按钮变回绿色 ▶） |
| 点击红色 ⏹ | 结束当前任务，自动记录一条工时 |
| 点击蓝色 💾 | 导出 Excel 到桌面（文件名带时间戳，自动防覆盖） |
| 点击窗口右上角 ✕ | 隐藏到系统托盘（进程继续运行） |
| 左键单击托盘图标 | 显示 / 隐藏主窗口 |
| 右键单击托盘图标 | 弹出托盘菜单（开机自启 / 退出） |

## 数据存储

时迹不会在你的安装目录留下任何文件。

| 平台 | 存储路径 |
|------|----------|
| Windows | `%APPDATA%\ShiJi\tasks.json` |
| macOS | `~/Library/Application Support/ShiJi/tasks.json` |
| Linux | `~/.config/ShiJi/tasks.json` |

首次启动时，会自动将旧版（v1.0.x）的 `tasks.json` 迁移至新目录。

## 右键菜单

在主窗口任意位置右键，可以：

- 拖动滑块调整窗口透明度（10%~100%）
- 切换「始终置顶」开关

## 技术栈

- Python 3.12
- PyQt5（图形界面）
- openpyxl（Excel 导出）
- PyInstaller（打包 exe）

## 开发与贡献

本项目使用 MIT 许可证，欢迎 Fork、修改、提 Issue。

如果你发现了 Bug，或者有好的功能建议，欢迎在 [Issues](https://github.com/jsmbjsm/TimeTrace/issues) 中提出。

## License

[MIT](./LICENSE) © 2026 记神名不祭神明
