from __future__ import annotations

import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape


DOCX_CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


DOCX_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


DOCX_DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


DOCX_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>
        <w:sz w:val="20"/>
        <w:szCs w:val="20"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>
"""


DOCX_APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Office Word</Application>
</Properties>
"""


def xml_text(text: str) -> str:
    return f'<w:t xml:space="preserve">{escape(text)}</w:t>'


def build_run_xml(text: str, *, bold: bool = False, size: int = 20) -> str:
    props = []
    if bold:
        props.extend(["<w:b/>", "<w:bCs/>"])
    props.append(f'<w:sz w:val="{size}"/>')
    props.append(f'<w:szCs w:val="{size}"/>')
    props_xml = f"<w:rPr>{''.join(props)}</w:rPr>"

    parts = []
    for i, line in enumerate((text or "").split("\n")):
        if i > 0:
            parts.append("<w:br/>")
        parts.append(xml_text(line))
    if not parts:
        parts.append(xml_text(""))
    return f"<w:r>{props_xml}{''.join(parts)}</w:r>"


def build_paragraph_xml(
    text: str,
    *,
    bold: bool = False,
    size: int = 20,
    align: str | None = None,
    space_after: int = 100,
    space_before: int = 0,
) -> str:
    props = [f'<w:spacing w:before="{space_before}" w:after="{space_after}"/>']
    if align:
        props.append(f'<w:jc w:val="{align}"/>')
    return f"<w:p><w:pPr>{''.join(props)}</w:pPr>{build_run_xml(text, bold=bold, size=size)}</w:p>"


def build_document_xml(title: str, lines: list[tuple[str, dict]]) -> str:
    body_parts = [build_paragraph_xml(title, bold=True, size=30, align="center", space_after=220)]
    for text, kwargs in lines:
        body_parts.append(build_paragraph_xml(text, **kwargs))
    body_parts.append(
        '<w:sectPr>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="900" w:bottom="1440" w:left="900" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14">'
        f"<w:body>{''.join(body_parts)}</w:body>"
        "</w:document>"
    )


def build_core_xml(title: str) -> str:
    created = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{escape(title)}</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>
"""


def write_docx(path: Path, title: str, lines: list[tuple[str, dict]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    document_xml = build_document_xml(title, lines)
    core_xml = build_core_xml(title)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", DOCX_CONTENT_TYPES)
        zf.writestr("_rels/.rels", DOCX_ROOT_RELS)
        zf.writestr("docProps/app.xml", DOCX_APP_XML)
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", DOCX_STYLES)
        zf.writestr("word/_rels/document.xml.rels", DOCX_DOCUMENT_RELS)


def write_docx_with_fallback(path: Path, title: str, lines: list[tuple[str, dict]]) -> Path:
    try:
        write_docx(path, title, lines)
        return path
    except PermissionError:
        alt = path.with_name(f"{path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{path.suffix}")
        write_docx(alt, title, lines)
        return alt


def build_manual_lines() -> list[tuple[str, dict]]:
    py_cmd = r'C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe'
    root = r'C:\Users\Administrator\Desktop\CA&US成分审核'
    lines: list[tuple[str, dict]] = []
    add = lines.append

    add((
        "适用范围：本说明书适用于 V2 版本工具，即 Excel 筛查脚本 `cosmetic_screening_windows_qwen_inline_v2.py` 和 Word 报告脚本 `cosmetic_screening_windows_qwen_inline_report_v2.py`。",
        {"size": 20, "space_after": 80},
    ))
    add((f"项目目录：{root}", {"size": 20, "space_after": 80}))

    add(("一、工具用途", {"bold": True, "size": 24, "space_before": 60, "space_after": 80}))
    add(("1. 本工具用于化妆品配方成分筛查。", {"size": 20, "space_after": 40}))
    add(("2. 第一步输出 Excel 筛查结果；第二步基于筛查结果输出 CA 和 US Word 报告。", {"size": 20, "space_after": 40}))
    add(("3. 员工日常使用时，通常只需要准备输入文件、运行脚本、检查输出。", {"size": 20, "space_after": 80}))

    add(("二、使用前准备", {"bold": True, "size": 24, "space_before": 60, "space_after": 80}))
    add(("请确认以下文件夹和文件已经放好：", {"size": 20, "space_after": 40}))
    add(("1. `input` 文件夹：放客户配方 Excel。", {"size": 20, "space_after": 30}))
    add(("2. `database` 文件夹：至少应包含 CA Hotlist Prohibited、CA Hotlist Restricted、CIR Quick Reference Table、CFR 禁用清单、US Color Additives Database。", {"size": 20, "space_after": 30}))
    add(("3. `template` 文件夹：应包含输出模板 `Cosmetic_Compliance_Output_Template_v2.xlsx`。", {"size": 20, "space_after": 30}))
    add(("4. 不要随意改数据库文件名、模板工作表名或表头。", {"size": 20, "space_after": 80}))

    add(("三、V2 Excel 筛查操作", {"bold": True, "size": 24, "space_before": 60, "space_after": 80}))
    add(("推荐做法：先打开 Windows PowerShell，再切换到项目目录。", {"size": 20, "space_after": 40}))
    add(("先运行以下命令进入项目文件夹：", {"size": 20, "space_after": 40}))
    add((f'cd "{root}"', {"size": 18, "space_after": 50}))
    add(("进入项目目录后，再运行以下命令：", {"size": 20, "space_after": 40}))
    add((f'& "{py_cmd}" ".\\cosmetic_screening_windows_qwen_inline_v2.py"', {"size": 18, "space_after": 60}))
    add(("说明：也可以把 `.py` 文件直接拖进 PowerShell 窗口，这样会自动填入脚本完整路径；但仍然需要正确的 Python 运行环境。对于不熟悉命令行的员工，更推荐使用固定命令。后续如有需要，也可以由负责人提供双击可运行的 `.bat` 启动文件。", {"size": 20, "space_after": 80}))
    add(("运行成功后，屏幕会显示：读取数据库、解析客户输入、执行筛查、输出文件路径、日志文件路径。", {"size": 20, "space_after": 40}))
    add(("运行结果会生成到：", {"size": 20, "space_after": 30}))
    add(("1. `output` 文件夹：生成 `Cosmetic_Compliance_Output_时间戳.xlsx`。", {"size": 20, "space_after": 30}))
    add(("2. `logs` 文件夹：生成 `run_log_时间戳.txt`。", {"size": 20, "space_after": 80}))

    add(("四、Excel 结果怎么看", {"bold": True, "size": 24, "space_before": 60, "space_after": 80}))
    add(("1. 每个配方会单独拆成一个 `03_配方名` 工作表。", {"size": 20, "space_after": 30}))
    add(("2. `Manual Review? = Y` 代表该成分需要人工确认。", {"size": 20, "space_after": 30}))
    add(("3. 色素一旦被识别，会进入 US color additive 模块，并强制要求人工复核。", {"size": 20, "space_after": 30}))
    add(("4. 香精目前也会被标记为人工复核。", {"size": 20, "space_after": 30}))
    add(("5. 红色、黄色、橙色高亮分别代表更高优先级风险或需要人工处理的项目。", {"size": 20, "space_after": 80}))

    add(("五、V2 Word 报告操作", {"bold": True, "size": 24, "space_before": 60, "space_after": 80}))
    add(("Word 报告脚本会从筛查结果 Excel 中读取每个配方的结果，并生成 CA 报告和 US 报告。", {"size": 20, "space_after": 40}))
    add(("1. 报告默认输出到 `output\\word_reports_时间戳` 文件夹。", {"size": 20, "space_after": 30}))
    add(("2. 每个配方会单独建立一个文件夹，每个配方通常会生成两份文件：", {"size": 20, "space_after": 30}))
    add(("   CA_Report_v2.docx：加拿大报告。", {"size": 20, "space_after": 20}))
    add(("   US_Report_v2.docx：美国报告。", {"size": 20, "space_after": 80}))

    add(("六、V2 的几个重点规则", {"bold": True, "size": 24, "space_before": 60, "space_after": 80}))
    add(("1. CAS 精确匹配优先于名称匹配。", {"size": 20, "space_after": 30}))
    add(("2. 名称未命中但 CAS 命中时，必须人工复核。", {"size": 20, "space_after": 30}))
    add(("3. 任何 restriction 命中都要重点看。", {"size": 20, "space_after": 30}))
    add(("4. 色素不对所有成分都筛查，只有识别出色素线索或 CAS 命中色素库时才进入色素模块。", {"size": 20, "space_after": 30}))
    add(("5. 色素无论最终判断是 Compliance 还是 Not Compliance，报告中都会提示人工复核。", {"size": 20, "space_after": 30}))
    add(("6. 香精会被标记为人工复核，并在报告中给出 IFRA 相关备用说明。", {"size": 20, "space_after": 80}))

    return lines


def main() -> None:
    root = Path(__file__).resolve().parent
    output_path = root / "V2使用说明书.docx"
    written = write_docx_with_fallback(output_path, "V2 使用说明书", build_manual_lines())
    print(written)


if __name__ == "__main__":
    main()
