# JSON → Excel 转换工具 (json2excel-gui)

一个基于 **Tauri** + **Rust** 开发的高性能 JSON 转 Excel 桌面应用。提供极简的交互体验，支持多种数据导入方式，并能智能识别嵌套的 JSON 数组结构。

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Version](https://img.shields.io/badge/version-0.1.0-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)

![App Mockup](docs/images/app_mockup.png)

## 📖 使用说明

### 1. 导入数据
工具提供三种灵活的导入方式：
- **文件模式**：点击“浏览”选择本地 `.json` 文件。
- **Curl / Fetch 模式**：直接粘贴开发工具中的网络请求命令，点击“执行请求”即可获取实时数据。
- **JSON 模式**：直接在编辑器中粘贴或编写 JSON 字符串，点击“解析 JSON”。

### 2. 字段选择（高级功能）
如果 JSON 中包含多个数组（例如一个响应中有 `users` 和 `orders`），工具会弹出字段选择面板。你可以勾选需要导出的字段，每个选中的字段将对应 Excel 中的一个 **Sheet**。

### 3. 开始导出
确认输出路径后，点击“开始转换”。转换完成后，点击“在文件夹中显示”即可直接跳转到生成的 Excel 文件。

## ✨ 特性

- **🚀 多种导入方式**:
  - **文件模式**: 直接选择本地 `.json` 文件。
  - **粘贴模式**: 直接粘贴 JSON 文本。
  - **请求模式**: 支持粘贴 `curl` 或 `fetch` 命令，自动执行并解析接口返回的 JSON。
- **🔍 智能结构分析**:
  - 自动递归扫描 JSON，发现所有潜在的数组字段。
  - 支持 **Sheet 勾选**: 发现多个数组时，你可以选择将哪些字段导出为 Excel 的不同工作表（Sheet）。
  - 支持 **矩阵模式**: 自动识别 `keys` + `values` 的二维数据格式。
- **🎨 现代交互 UI**:
  - 丝滑的 **深色/浅色/跟随系统** 主题切换。
  - 转换完成后支持 **一键打开目录**。
  - 实时转换进度与错误提示。
- **⚡ 极速性能**: 后端采用 Rust 编写，使用 `rust_xlsxwriter` 库，处理大数据量时依然轻量迅捷。

## 🛠️ 技术栈

- **后端**: [Rust](https://www.rust-lang.org/) + [Tauri v2](https://tauri.app/)
- **Excel 生成**: [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter)
- **前端**: 原生 HTML / CSS / JavaScript (Vanilla)
- **JSON 解析**: [serde_json](https://github.com/serde-rs/json)

## 📦 快速开始

### 环境依赖

- [Rust](https://www.rust-lang.org/tools/install)
- [Node.js](https://nodejs.org/) (建议使用 LTS 版本)

### 开发与运行

1. **安装依赖**:
   ```bash
   npm install
   ```

2. **启动开发环境**:
   ```bash
   npm run tauri dev
   ```

3. **构建安装包**:
   ```bash
   npm run tauri build
   ```

## 📂 项目结构

```text
├── src/               # 前端代码 (HTML/JS/CSS)
├── src-tauri/         # Rust 后端代码
│   ├── src/
│   │   ├── main.rs    # 程序入口 (移动端/桌面端入口)
│   │   └── lib.rs     # 核心命令逻辑 (JSON 解析与 Excel 生成)
│   └── tauri.conf.json # Tauri 配置文件
└── package.json       # 项目元数据与脚本
```

## 📄 开源协议

本项目基于 [MIT](LICENSE) 协议开源。

---
*Created with ❤️ by YouDing*
