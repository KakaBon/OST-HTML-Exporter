# OST-HTML-Exporter

一个把 Excel 工作簿导出为 **纯观看 HTML 文件**的 Windows 小工具。

它适合这样的场景：Excel 原件继续保留公式、宏和可编辑内容；需要给别人查看时，再生成一个可以直接用浏览器打开的 HTML 观看版。宽表不会像 PDF / XPS 那样受到纸张宽度和分页限制，超出浏览器窗口时可以直接横向滚动。

## 能做什么

- 支持 `.xlsx` 和 `.xlsm` 工作簿。
- 默认优先导出 `OST Timeline` 和 `OST Collection`。
- 如果工作簿中不存在这两个工作表，则导出名称中 **不包含 `Backup`** 的工作表。
- 隐藏行和隐藏列不会出现在导出的 HTML 中。
- Excel 中的宏按钮等操作控件不会作为 HTML 内容导出。
- 尽量保留当前工作表的主要显示效果，包括：
  - 列宽；
  - 行高；
  - 字体、字号、粗体、斜体、下划线；
  - 文字颜色；
  - 单元格背景填充色；
  - 对齐方式；
  - 自动换行；
  - 边框；
  - 合并单元格。
- 公式单元格优先显示工作簿中已经保存的计算结果。
- 导出完成后在默认浏览器自动打开生成的 HTML 文件。

## 环境需求

需要：

- Windows
- Python
- Windows 的 `py` Python Launcher
- `openpyxl`

日常直接运行 `Export Viewer.bat` 即可。启动脚本会检查 `openpyxl`；如果当前 Python 环境中没有安装，会尝试自动安装。

## 文件结构

```text
OST-HTML-Exporter/
├── Export Viewer.bat
└── export_viewer.py
└── README.md
```

- `Export Viewer.bat`  
  日常使用入口。可以双击后选择工作簿，也可以直接把工作簿拖到它上面。

- `export_viewer.py`  
  实际执行 Excel → HTML 导出的程序。普通使用时不需要单独运行。

## 使用方法

### 方法一：双击 `Export Viewer.bat` 后选择工作簿

1. 双击：

   ```text
   Export Viewer.bat
   ```

2. 程序会弹出文件选择窗口。
3. 选择需要导出的 `.xlsx` 或 `.xlsm` 工作簿。
4. 等待导出完成。
5. 完成后会弹出提示，并自动在默认浏览器中打开生成的 HTML。

输出文件生成在 **原 Excel 工作簿所在目录**。

例如原文件是：

```text
Payback the Series OST.xlsm
```

则生成：

```text
Payback the Series OST_Viewer.html
```

### 方法二：直接把工作簿拖到 `Export Viewer.bat`

如果已经在文件资源管理器中找到目标工作簿，可以不先打开工具。

1. 找到需要导出的 `.xlsx` 或 `.xlsm` 文件。
2. 用鼠标把这个 Excel 文件直接拖到：

   ```text
   Export Viewer.bat
   ```

3. 松开鼠标。
4. 工具会直接处理刚刚拖入的工作簿，不再弹出文件选择窗口。
5. 导出完成后，同样会在原工作簿所在目录生成：

   ```text
   原文件名_Viewer.html
   ```

6. 导出完成后会自动打开 HTML。

## 工作表导出规则

程序首先检查工作簿中是否存在：

```text
OST Timeline
OST Collection
```

只要存在其中一个，就只导出实际存在的这些目标工作表。

如果两个都不存在，则启用通用规则：

```text
导出名称中不包含 Backup 的工作表
```

例如：

```text
Sheet1
Sheet2
Backup
OST Data Backup
```

会导出 `Sheet1` 和 `Sheet2`，名称中含 `Backup` 的工作表不会导出。

> 注意：隐藏“行 / 列”会被排除；通用规则本身不按照“工作表是否隐藏”来筛选，而是按照工作表名称是否包含 `Backup` 来筛选。

## 公式与显示结果

如果工作簿中有公式，HTML 观看版优先显示 Excel 文件中已经保存的公式计算结果。

因此，在导出包含公式的工作簿前，建议先：

1. 用 Excel / WPS 打开工作簿；
2. 让公式完成计算；
3. 保存工作簿；
4. 再运行导出。

这样可以避免 HTML 中使用到旧的计算结果。

## 输出文件怎么分享

生成的 `_Viewer.html` 是一个独立 HTML 文件。

通常只需要把这个 HTML 文件发给别人，对方就可以直接用浏览器查看，不需要运行本工具，也不需要安装 Excel。
