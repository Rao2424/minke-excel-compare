# minke-excel-parser

基于 [excel-parser](https://github.com/knowledgestack/excel-parser) 定制开发的 Excel 解析项目。读取工作簿，保留单元格、公式、合并区域和表格结构，输出结构化 JSON 及带来源位置的文本分块。

当前沿用 Python 包名 `excel_parser`，需要 Python 3.10+。

## 安装

在包含 `pyproject.toml` 的项目目录中执行：

```powershell
python -m venv .venv
.venv/Scripts/python.exe -m pip install -e ".[dev,api]"
```

Linux / macOS 激活虚拟环境后执行 `python -m pip install -e ".[dev,api]"`。仅使用解析库时可安装 `python -m pip install -e .`。

## 使用

```python
import json
from pathlib import Path

from excel_parser import parse_workbook

result = parse_workbook(path="workbook.xlsx")
Path("result.json").write_text(
    json.dumps(result.to_json(), ensure_ascii=False, indent=2),
    encoding="utf-8",
)

for chunk in result.chunks:
    print(chunk.source_uri, chunk.render_text)
```

默认 `full` 模式生成工作簿结构及文本、HTML 分块。`mode="fast"` 跳过后续模板处理、分块渲染和公式依赖图构建，返回的 `chunks` 为空。

支持 `.xlsx`、`.xlsm` 和旧版 `.xls`。`.xls` 优先通过本机 LibreOffice 转换，缺少时回退到 xlrd；回退转换无法保留公式文本和图表。解析器不执行宏，也不重新计算公式。

## 本地 API

```powershell
.venv/Scripts/excel-parser-api.exe
```

启动后访问 <http://localhost:8080> 上传文件，或使用 `POST /parse`，以 multipart 字段 `file` 上传工作簿。接口返回解析结果和阶段验证报告。

## 开发验证

```powershell
.venv/Scripts/python.exe -m pytest tests/ -q --tb=short -W ignore::UserWarning
.venv/Scripts/python.exe -m ruff check src/ tests/
.venv/Scripts/python.exe -m mypy src/excel_parser
```

外部语料测试默认排除。如需使用自己的真实工作簿测试，将 `.xlsx` 样本放入 `tests/fixtures/corpus/<样本分组>/`，再执行：

```powershell
.venv/Scripts/python.exe -m pytest tests/test_corpus_robustness.py -m corpus -k "not download" -v
```

构建安装包：

```powershell
.venv/Scripts/python.exe -m pip install build
.venv/Scripts/python.exe -m build --wheel
```

## 项目结构

- `src/excel_parser/`：解析、结构分析、公式依赖、分块、渲染和序列化。
- `tests/`：核心回归测试与工作簿测试夹具。
- `rust/excel_core/`：可选 Rust 加速模块。
- `pyproject.toml`、`uv.lock`：依赖、打包和测试配置。

分块的 `token_count` 按字符数估算。公式值依赖文件中已有的缓存；透视表、迷你图和嵌入对象的解析存在限制。

## 来源与许可

上游项目为 [knowledgestack/excel-parser](https://github.com/knowledgestack/excel-parser)。保留原始 [MIT 许可证及版权声明](LICENSE)。
