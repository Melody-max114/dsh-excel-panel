# 📊 dsh-excel-panel

> DSH 右侧栏 Excel 编辑插件：在 `dsh-better-sidebar` 中直接预览、编辑和保存 `.xlsx` 文件。

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-DSH%20Web-blue)
![Version](https://img.shields.io/badge/version-0.6.0-orange)

## 简介

`dsh-excel-panel` 是一个面向 DSH Web 的右侧栏插件，让你不用离开聊天界面就能：

- 打开 `.xlsx` 文件并直接编辑单元格
- 像 Excel 一样操作工作表、合并单元格、调整字体样式
- 实时预览公式计算结果
- 一键保存回原文件

## ✨ 功能特性

### 交互

- 单击选中单元格
- 双击编辑文字
- 点击空白处取消选中
- 拖动 / Shift+点击多选范围
- 多选范围使用淡蓝色半透明叠加，不遮挡单元格底色

### 格式工具栏

- 字体大小下拉框
  - 多选时如果字体大小不一致，显示空白“字号”
- 加粗 `B` / 斜体 `I` / 下划线 `U`
  - 多选时能识别统一状态
- 左对齐 / 居中 / 右对齐
- 字体颜色菜单
  - 20 个标准色
  - 自定义颜色：5 个自定义色块 + RGB 滑块 + HEX 输入
- 背景颜色菜单
  - 20 个标准色
  - 自定义颜色：5 个自定义色块 + RGB 滑块 + HEX 输入

### 自定义颜色

- 左键点击自定义色块：直接使用该颜色
- 右键点击自定义色块：打开内嵌色盘设置颜色
- 色盘内支持：
  - 40 个快速选色
  - RGB 三通道滑块
  - HEX 输入框
  - 预览色块
  - `✓` 确认按钮

### 单元格与工作表

- 公式栏 `fx`
- 公式实时推导：
  - `+ - * / ^`、单元格/区域引用
  - `SUM`、`ROUND`、`IF`、`MAX`、`MIN`、`AVERAGE`、`COUNT`、`ABS`、`INT`、`LEN`、`UPPER`、`LOWER`
- 合并 / 取消合并
- 插入 / 删除行和列
- 列宽 / 行高拖拽调整
- 底部 Sheet 栏 + 新建工作表
- 右键菜单：复制 / 粘贴 / 清除 / 合并 / 插入删除行列

### 保存

- 每次保存都会**重建整个 xlsx**，避免样式表逐步损坏
- 多次保存格式不丢失
- 保存失败时提示具体原因
- 文件被占用时提供“关闭占用程序并重试”

### 撤销 / 重做

- `Ctrl+Z` 撤销
- `Ctrl+Y` / `Ctrl+Shift+Z` 重做

## 📦 安装

### 方式一：npm 安装（推荐）

```bash
dsh plugin --profile web add dsh-excel-panel
```

或者直接安装到项目：

```bash
npm install dsh-excel-panel
```

### 方式二：本地 file 依赖

将插件目录放到本地，例如：

```text
D:\apps\DP专武\dsh-excel-panel
```

在 DSH profile 目录（通常为 `~/.dsh/profiles/web`）执行：

```bash
dsh plugin --profile web add file:D:/apps/DP专武/dsh-excel-panel
```

### 方式三：zip 包安装

从 GitHub Releases 下载 `dsh-excel-panel-v0.6.0.zip`，解压到任意目录，再按方式一添加 file 依赖即可。

## 🚀 使用

1. 重启 DSH Web。
2. 在右侧栏打开任意 `.xlsx` 文件。
3. 单击选中单元格，双击编辑文字；也可以在上方 `fx` 栏输入公式，例如 `=D2*0.15`。
4. 多选范围后，使用格式工具栏批量调整字体、颜色、对齐等。
5. 点击右上角「保存」写回磁盘原文件。

## ⌨️ 快捷键

| 快捷键 | 功能 |
| --- | --- |
| `Ctrl+Z` | 撤销 |
| `Ctrl+Y` / `Ctrl+Shift+Z` | 重做 |
| 按住 Shift 点击 | 选择范围 |
| 鼠标拖动 | 选择范围 |
| 双击 | 编辑单元格 |
| `Enter` | 编辑时确认并下移 |
| `Tab` | 编辑时右移 |

## 🧩 插件结构

```text
dsh-excel-panel/
├── package.json
├── cordis.patch.yml
├── README.md
├── CHANGELOG.md
├── DEBUG_V2.md
└── lib/
    ├── index.js      # 后端：读取/写入 xlsx
    └── client.js     # 前端：Excel 编辑器 UI
```

## 🔌 后端接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `POST` | `/excel-panel/read` | 读取 xlsx，返回单元格数据、公式、样式、合并区域 |
| `POST` | `/excel-panel/write` | 重建并写回 xlsx |
| `POST` | `/excel-panel/unlock` | 关闭占用目标文件的进程 |
| `POST` | `/excel-panel/log` | 写入操作日志 |

## ⚙️ 开发

```bash
# 修改后检查语法
node --check lib/index.js
node --check lib/client.js

# 同步到 profile 运行目录
cp -r lib ~/.dsh/profiles/web/node_modules/dsh-excel-panel/
```

## 📄 License

当前仓库使用 [MIT](LICENSE)，允许自由使用、修改和分发。
