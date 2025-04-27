# Dizzylab Redeem Card Print Generator

wheatfox (wheatfox17@icloud.com)

一个用于生成 dizzylab 兑换码卡片的工具，支持自定义背景、logo和布局。

## 功能特点

- 支持从Excel文件读取兑换码（请先把 dizzylab 导出的 csv 另存为 xlsx 文件，用于之后在第 4 列之后进行额外标记）
- 支持自定义背景图片（自动处理为勾线效果）
- 支持自定义个人logo（自动处理为勾线效果）
- 支持自定义卡片布局和样式
- 自动跳过已兑换的码（在第 4 列对应行标记 "skip" 即可）
- 支持缓存处理后的图片

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

1. 准备Excel文件，格式如下：
   - 第一列：兑换码
   - 第二列：用户ID（已兑换的码）
   - 第三列：制品号
   - 第四列：skip标记（可选）

2. 准备配置文件 `config.yml`：
   ```yaml
   input:
     excel_file: "your_excel_file.xlsx"

   output:
     pdf_file: "output.pdf"

   card:
     album_name: "专辑名称"
     additional_info:
       - "兑换网址: https://www.dizzylab.net/redeem"
     background: "background.jpg"  # 背景图片路径
     personal_logo: "personal_logo.png"  # 个人logo路径

   layout:
     grid:
       columns: 2  # 每行卡片数量
       rows: 5    # 每页行数
     card:
       width: 8   # 卡片宽度(cm)
       height: 5  # 卡片高度(cm)
       margin: 0.5  # 页面边距(cm)
       font_size: 10  # 字体大小
   ```

3. 运行生成器：
   ```bash
   python3 generate.py config.yml
   ```

## 目录结构

```
.
├── generate.py          # 主程序
├── config.yml          # 配置文件
├── requirements.txt    # 依赖列表
├── fonts/             # 字体目录
│   ├── NotoSansSC-SemiBold.ttf
│   └── SFMono-Regular.ttf
├── logo/              # logo目录
│   └── dl-n-88_2.jpg
└── cache/             # 图片缓存目录
```

## 注意事项

1. 确保Excel文件格式正确
2. 背景图片和个人logo会自动处理为勾线效果
3. 处理后的图片会缓存在cache目录
4. 已兑换的码会自动跳过
5. 支持自定义卡片布局和样式