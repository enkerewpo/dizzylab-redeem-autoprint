# Dizzylab 兑换码卡片生成器

这个工具可以从包含兑换码的 Excel 文件中生成可打印的 Dizzylab 平台兑换卡片。

## 环境要求

- Python 3.x

## 安装步骤

1. 安装必需的 Python 包：
   ```bash
   pip install -r requirements.txt
   ```

## 使用步骤

### 1. 准备 Excel 文件

如果从 Dizzylab 下载的是 CSV 文件，需要先转换为 XLSX 格式：

1. 使用 Excel 打开 CSV 文件
2. 点击"文件" -> "另存为"
3. 选择保存类型为"Excel 工作簿 (*.xlsx)"
4. 保存文件

### 2. 创建配置文件

创建一个 `config.yml` 文件，结构如下：

```yaml
# 输入文件配置
input:
  excel_file: "path/to/your/input.xlsx"  # 输入的Excel文件路径

# 输出文件配置
output:
  pdf_file: "output.pdf"     # 输出的PDF文件路径

# 卡片内容配置
card:
  album_name: "专辑名称"
  additional_info:           # 额外信息，每行一个
    - "这里是额外信息1"
    - "这里是额外信息2"
    - "兑换网址: https://www.dizzylab.net/redeem"

# 布局配置
layout:
  grid:
    columns: 3              # 每行卡片数量
    rows: 4                 # 每页行数
  card:
    width: 6               # 卡片宽度(cm)
    height: 4              # 卡片高度(cm)
    margin: 1              # 页面边距(cm)
    font_size: 10          # 字体大小
```

### 3. 运行脚本

1. 打开终端，进入脚本所在目录
2. 运行以下命令：
   ```bash
   python generate.py config.yml
   ```

## 注意事项

Excel 文件除了 dizzylab 导出时的默认三列外，第四列可以用于标记跳过.

## TODO

1. 中文字体优化，现在用的是 PDF 自带的 Adobe 宋体（自带的只有宋体），但是 reportlab 导入 TTF 中文字体总是有问题
2. 排版优化（？）