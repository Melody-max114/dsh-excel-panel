# 📊 dsh-excel-panel

> DSH 右侧栏 Excel 编辑插件：在 `dsh-better-sidebar` 中直接预览、编辑和保存 `.xlsx` 文件。

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
![Platform](https://img.shields.io/badge/platform-DSH%20Web-blue)
![Version](https://img.shields.io/badge/version-0.3.0-orange)

## 简介

`dsh-excel-panel` 是一个面向 DSH Web 的右侧栏插件，让你不用离开聊天界面就能：

- 打开 `.xlsx` 文件并直接编辑单元格
- 像 Excel 一样操作工作表、合并单元格、调整字体样式
- 实时预览公式计算结果
- 一键保存回原文件

## ✨ 功能特性

- **可编辑 Excel 网格**
  - 行号 / 列标 / 工作表标签
  - 底部 Sheet 栏 + `+` 新建工作表
  - 单元格直接编辑，支持公式栏 `fx`

- **公式实时推导**
  - 支持 `+ - * / ^`、单元格/区域引用
  - 支持常用函数：`SUM`、`ROUND`、`IF`、`MAX`、`MIN`、`AVERAGE`、`COUNT`、`ABS`、`INT`、`LEN`、`UPPER`、`LOWER`
  - 编辑时自动重算相关公式

- **格式化**
  - 字体大小
  - 加粗 / 斜体 / 下划线
  - 字体颜色
  - 背景填充色

- **合并单元格**
  - 拖动选择范围
  - 合并 / 取消合并
  - 保存合并区域到 xlsx

- **右键菜单**
  - 复制 / 粘贴 / 清除内容
  - 合并 / 取消合并
  - 插入 / 删除行和列

- **键盘操作**
  - 方向键移动选中
  - Enter 编辑并下移
  - Tab 右移
  - Ctrl+Z / Ctrl+Y 撤销重做

- **撤销 / 重做**
  - `Ctrl+Z` 撤销
  - `Ctrl+Y` / `Ctrl+Shift+Z` 重做

- **实时刷新**
  - 自动检测外部文件修改
  - 手动「刷新」按钮


## 🙏 致谢 / 基于

本插件基于 [dsh-better-sidebar](https://github.com/omdsh-dev/DSH-better-sidebar) 的文件查看器扩展机制开发，使用其 `registerFileViewer` 能力在右侧栏注册可编辑的 Excel 预览。

同时参考了 Excel 的交互方式，以及 DSH 右侧面板/Office 预览类插件的设计思路。

> 如果原插件作者认为需要额外声明或调整归属，请联系本仓库作者补充。

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

或者手动在 `package.json` 添加：

```json
"dependencies": {
  "dsh-excel-panel": "file:D:/apps/DP专武/dsh-excel-panel"
}
```

并在 `dsh.profile.bundles` 中加入：

```json
"dsh-excel-panel"
```

然后执行：

```bash
pnpm install
```

### 方式三：zip 包安装

从 GitHub Releases 下载 `dsh-excel-panel-v0.3.0.zip`，解压到任意目录，再按方式一添加 file 依赖即可。

## 🚀 使用

1. 重启 DSH Web。
2. 在右侧栏打开任意 `.xlsx` 文件。
3. 单击选中单元格，双击编辑文字；也可以在上方 `fx` 栏输入公式，例如 `=D2*0.15`。
4. 多选范围后，使用格式工具栏批量调整字体、下划线、颜色等。
5. 点击右上角「保存」写回磁盘原文件。

## ⌨️ 快捷键

| 快捷键 | 功能 |
| --- | --- |
| `Ctrl+Z` | 撤销 |
| `Ctrl+Y` / `Ctrl+Shift+Z` | 重做 |
| 按住 Shift 点击 | 选择范围 |
| 鼠标拖动 | 选择范围 / 填充（右下角小方块） |

## 🧩 插件结构

```text
dsh-excel-panel/
├── package.json
├── cordis.patch.yml
├── README.md
└── lib/
    ├── index.js      # 后端：读取/写入 xlsx
    └── client.js     # 前端：Excel 编辑器 UI
```

## 🔌 后端接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `POST` | `/excel-panel/read` | 读取 xlsx，返回单元格数据、公式、样式、合并区域 |
| `POST` | `/excel-panel/write` | 将编辑后的数据写回原 xlsx |

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

