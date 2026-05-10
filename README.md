# 竞品新品周报生成工具

这个项目用于从日常维护的钉钉 AI 表中拉取新品数据，生成接近示例版式的周报 Excel。

## 项目结构

```text
competitor-new-product-monitor/
├── README.md
├── .gitignore
├── requirements.txt
├── config/
│   ├── dingtalk.example.json
│   ├── dingtalk.json
│   ├── field_mapping.json
│   ├── report_rules.json
│   ├── excel_layout.json
│   └── font_files.json
├── assets/
│   └── fonts/
│       └── README.md
├── scripts/
│   └── generate_weekly_report.py
└── outputs/
```

- `config/dingtalk.example.json`：可提交的钉钉连接配置模板，不含真实 token。
- `config/dingtalk.json`：本地私密钉钉连接配置，放真实 Streamable HTTP URL 或使用本机 mcporter 服务名。这个文件已被 `.gitignore` 忽略。
- `config/field_mapping.json`：钉钉字段如何翻译成脚本标准字段。
- `config/report_rules.json`：这份周报的业务规则。
- `config/excel_layout.json`：Excel 怎么排版、数据写到哪里。
- `config/font_files.json`：内部中英文字体文件路径配置。
- `assets/fonts/`：字体文件本地放置位置。字体不要提交、不要上传外部。
- `scripts/generate_weekly_report.py`：生成脚本。
- `outputs/`：默认 Excel 输出目录。

## 环境准备

脚本需要 Python 3、`openpyxl`、`Pillow` 和 `mcporter`。

```bash
pip install -r requirements.txt
```

如果本机已经通过 Codex 运行，也可以使用 Codex bundled Python：

```bash
/Users/cheryl/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/generate_weekly_report.py --validate-config
```

## 配置钉钉连接

第一次使用时复制模板：

```bash
cp config/dingtalk.example.json config/dingtalk.json
```

`config/dingtalk.json` 里有两种配置方式：

1. 如果本机 `mcporter` 已经配置了 `dingtalk-ai-table`，保持 `streamableHttpUrl` 为空即可。
2. 如果交接给继任者或需要换账号，把新的 Streamable HTTP URL 填入 `streamableHttpUrl`。

注意：Streamable HTTP URL 带访问令牌，不要提交到 git，不要上传外部平台。

## 字体配置

本项目使用以下内部字体：

- 中文：`方正FW筑紫黑简 R.ttf`
- 英文/数字：`Heytea Sans Serif Regular.otf`

默认字体路径写在 `config/font_files.json`：

```json
{
  "chineseFont": {
    "sourcePath": "/Users/cheryl/Downloads/字体&办公模版/方正FW筑紫黑简 R.ttf"
  },
  "latinFont": {
    "sourcePath": "/Users/cheryl/Downloads/字体&办公模版/Heytea Sans Serif Regular.otf"
  }
}
```

也可以把字体文件放到 `assets/fonts/`，再把 `projectPath` 改成实际文件名。字体是内部独家字体，不提交、不上传外部。Excel 文件只记录字体名称，不会嵌入字体文件，打开 Excel 的电脑也需要安装对应字体。

## CLI 命令

默认生成最近一个完整“周六到周五”周期：

```bash
python scripts/generate_weekly_report.py
```

指定周期：

```bash
python scripts/generate_weekly_report.py --start-date 2026-05-02 --end-date 2026-05-08
```

指定品牌：

```bash
python scripts/generate_weekly_report.py --brands 霸王茶姬,古茗,乐乐茶
```

同时指定周期和品牌：

```bash
python scripts/generate_weekly_report.py --start-date 2026-05-02 --end-date 2026-05-08 --brands 霸王茶姬,古茗
```

指定输出文件：

```bash
python scripts/generate_weekly_report.py --output outputs/竞品新品周报2026-05-02_2026-05-08.xlsx
```

校验配置：

```bash
python scripts/generate_weekly_report.py --validate-config
```

解释当前配置：

```bash
python scripts/generate_weekly_report.py --explain-config
```

## 当前业务规则

默认跟踪品牌：

霸王茶姬、茶百道、古茗、奈雪的茶、乐乐茶、阿嬷手作、去茶山、爷爷不泡茶、瑞幸、上山喝茶、茉莉奶白

日期规则：

- 默认取最近一个完整“周六到周五”。
- 可以用 `--start-date` 和 `--end-date` 覆盖。

品类规则：

- 品类只显示一个。
- 优先级：`茶特调 > 果蔬茶 > 柠檬茶 > 主品类`。
- 先看钉钉字段 `附加品类` 是否命中前三个标签，命中多个时按优先级只取一个。
- 都未命中时显示 `主品类`。
- 后续优先级变化时，改 `config/report_rules.json` 的 `categoryRule.priority`。

备注规则：

- Excel 备注列按 `回归`、`联名`、`备注` 顺序拼接。
- `回归` 只显示“回归”，不显示“上新”。
- 钉钉 `联名` 字段只写品牌/IP，Excel 输出时自动补成 `{品牌IP}联名`。
- 非空项用中文逗号 `，` 分隔；都为空显示 `/`。

价格规则：

- 价格文本中括号里的价格加删除线，例如 `15元(17元)(中杯)` 只把 `17元` 加删除线。
- 汇总表价格列如果需要换行，只在 `/` 后换行，不在普通文字中间断开。

图片和行高：

- 产品外观图片高度统一 `4cm`。
- 表格行高按内容估算到合适状态，尽量完整显示内容，并保留最大行高上限。

## 用自然语言调整配置

后续可以直接让 Codex 按自然语言修改配置。建议这样说：

- “把默认品牌加上沪上阿姨。”
- “品类优先级改成柠檬茶 > 果蔬茶 > 茶特调 > 主品类。”
- “产品图片高度改成 5cm。”
- “卖点介绍行高上限放宽到 90。”
- “钉钉字段‘主品类’改名叫‘核心品类’，同步字段映射。”

一般对应关系：

- 钉钉字段名、字段 ID、取值方式：改 `config/field_mapping.json`
- 日期、品牌、品类、备注、价格等业务规则：改 `config/report_rules.json`
- 字体、颜色、列宽、行高、图片高度、写入位置：改 `config/excel_layout.json`
- 字体文件路径：改 `config/font_files.json`

调整后建议运行：

```bash
python scripts/generate_weekly_report.py --validate-config
python scripts/generate_weekly_report.py --explain-config
```
