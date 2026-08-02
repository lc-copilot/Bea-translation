"""可复用的 DOCX/OOXML 工具：校验、ZIP/XML 读写、段落枚举、稳定 ID、格式化、媒体哈希。

仅使用 Python 3.10+ 标准库。
"""
from collections.abc import Iterator
from hashlib import sha256
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
TRACKED_TAGS = {f"{{{W_NS}}}ins", f"{{{W_NS}}}del",
                f"{{{W_NS}}}moveFrom", f"{{{W_NS}}}moveTo"}
PRIMARY_PARTS = ("word/document.xml",)
OPTIONAL_PREFIXES = ("word/header", "word/footer")
OPTIONAL_PARTS = ("word/footnotes.xml", "word/endnotes.xml")

_XML_PARTS_PREFIXES = ("word/document.xml", "word/header", "word/footer",
                       "word/footnotes.xml", "word/endnotes.xml")
_HEADING_STYLES = ("Title", "Subtitle", "Heading", "heading", "标题", "题注")


class DocxError(ValueError):
    pass


def qn(name: str) -> str:
    """把 'w:p' 形式的名称展开为带命名空间的完整标签名。"""
    if name.startswith("{"):
        return name
    prefix, _, local = name.partition(":")
    ns = {"w": W_NS, "r": R_NS}.get(prefix, "")
    return f"{{{ns}}}{local}" if ns else local


def source_sha256(path: Path) -> str:
    return sha256(Path(path).read_bytes()).hexdigest()


def read_parts(path: Path) -> dict[str, bytes]:
    """读取 DOCX 为 {member: bytes}。坏 ZIP 抛 DocxError。"""
    path = Path(path)
    try:
        with zipfile.ZipFile(path) as zf:
            parts = {name: zf.read(name) for name in zf.namelist()}
    except (zipfile.BadZipFile, OSError) as exc:
        raise DocxError(f"cannot read docx package: {exc}") from exc
    return parts


def write_parts(parts: dict[str, bytes], output: Path) -> None:
    """把 parts 写为 DOCX。顺序：document → 排序 headers → 排序 footers → footnotes → endnotes → 其余。"""
    output = Path(output)

    def sort_key(name: str):
        if name == "word/document.xml":
            return (0, name)
        if name.startswith("word/header"):
            return (1, name)
        if name.startswith("word/footer"):
            return (2, name)
        if name == "word/footnotes.xml":
            return (3, name)
        if name == "word/endnotes.xml":
            return (4, name)
        return (5, name)

    ordered = sorted(parts, key=sort_key)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in ordered:
            zf.writestr(name, parts[name])


def translatable_parts(parts: dict[str, bytes]) -> tuple[str, ...]:
    """返回可翻译 part 的确定性顺序：document、排序 headers、排序 footers、footnotes、endnotes。"""
    headers = sorted(n for n in parts if n.startswith("word/header") and n.endswith(".xml"))
    footers = sorted(n for n in parts if n.startswith("word/footer") and n.endswith(".xml"))
    result = []
    for name in PRIMARY_PARTS:
        if name in parts:
            result.append(name)
    result.extend(headers)
    result.extend(footers)
    for name in OPTIONAL_PARTS:
        if name in parts:
            result.append(name)
    return tuple(result)


def has_tracked_changes(parts: dict[str, bytes]) -> bool:
    for name in parts:
        if not name.endswith(".xml") or not name.startswith(_XML_PARTS_PREFIXES):
            continue
        try:
            root = ET.fromstring(parts[name])
        except ET.ParseError:
            continue
        if any(e.tag in TRACKED_TAGS for e in root.iter()):
            return True
    return False


def validate_docx(path: Path) -> dict[str, bytes]:
    """校验 DOCX：扩展名、ZIP 完整性、必需成员、加密、修订痕迹、图表/图形文字。返回 parts。"""
    path = Path(path)
    if path.suffix.lower() != ".docx":
        raise DocxError(f"expected a .docx file, got {path.suffix or 'no extension'}")
    try:
        parts = read_parts(path)
    except DocxError as exc:
        raise DocxError(str(exc)) from exc
    if not parts:
        raise DocxError("empty docx package")
    required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
    missing = sorted(required - set(parts))
    if missing:
        raise DocxError(f"docx package missing members: {', '.join(missing)}")
    if any(name.startswith("EncryptedPackage") or name.startswith("EncryptionInfo")
           for name in parts):
        raise DocxError("encrypted docx is not supported")
    if has_tracked_changes(parts):
        raise DocxError("tracked changes are not supported")
    unsupported = [n for n in parts
                   if n.startswith("word/charts/") or n.startswith("word/diagrams/")
                   or n.startswith("word/chart")]
    if unsupported:
        raise DocxError(f"chart or diagram text is not supported: {unsupported[0]}")
    return parts


def iter_paragraphs(xml_bytes: bytes) -> Iterator[tuple[int, ET.Element]]:
    """按文档顺序产出 (段落序号, w:p 元素)；跳过空段落与脚注分隔符 ID -1/0。"""
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as exc:
        raise DocxError(f"invalid xml part: {exc}") from exc
    separators = [
        fn for fn in root.iter(qn("w:footnote"))
        if fn.get(qn("w:id")) in ("-1", "0")
    ]

    def in_separator(paragraph: ET.Element) -> bool:
        for sep in separators:
            if any(node is paragraph for node in sep.iter()):
                return True
        return False

    index = 0
    for elem in root.iter(qn("w:p")):
        ordinal = index
        index += 1
        if in_separator(elem):
            continue
        if paragraph_text(elem).strip() == "":
            continue
        yield ordinal, elem


_TEXT_CONTAINERS = {qn("w:ins"), qn("w:del"), qn("w:moveFrom"),
                    qn("w:moveTo"), qn("w:hyperlink"), qn("w:smartTag")}


def paragraph_text(paragraph: ET.Element) -> str:
    parts = []
    for child in paragraph:
        if child.tag == qn("w:r"):
            parts.append(_run_text(child))
        elif child.tag in _TEXT_CONTAINERS:
            for run in child.iter(qn("w:r")):
                parts.append(_run_text(run))
    return "".join(parts)


def _run_text(run: ET.Element) -> str:
    parts = []
    for node in run:
        if node.tag == qn("w:t"):
            parts.append(node.text or "")
        elif node.tag == qn("w:tab"):
            parts.append("\t")
        elif node.tag == qn("w:br"):
            parts.append("\n")
    return "".join(parts)


def paragraph_style(paragraph: ET.Element) -> str:
    for node in paragraph.iter():
        if node.tag == qn("w:pStyle"):
            return node.get(qn("w:val")) or ""
    return ""


def block_kind(part: str, paragraph: ET.Element) -> str:
    style = paragraph_style(paragraph)
    if part == "word/footnotes.xml":
        return "footnote"
    if part == "word/endnotes.xml":
        return "endnote"
    if part.startswith("word/header"):
        return "header"
    if part.startswith("word/footer"):
        return "footer"
    if style.startswith(_HEADING_STYLES):
        return "heading"
    return "body"


def stable_block_id(part: str, paragraph_index: int, text: str) -> str:
    digest = sha256(f"{part}\0{paragraph_index}\0{text}".encode("utf-8")).hexdigest()
    return "b-" + digest[:16]


def media_hashes(path: Path) -> dict[str, str]:
    path = Path(path)
    try:
        with zipfile.ZipFile(path) as zf:
            result = {}
            for name in zf.namelist():
                if name.startswith("word/media/"):
                    result[name] = sha256(zf.read(name)).hexdigest()
            return result
    except (zipfile.BadZipFile, OSError) as exc:
        raise DocxError(f"cannot read media hashes: {exc}") from exc
