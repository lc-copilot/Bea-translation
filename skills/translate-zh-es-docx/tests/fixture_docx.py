"""确定性 DOCX 夹具构建器（仅测试使用，Python 标准库）。"""
import base64
import zipfile
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)

REQUIRED_MEMBERS = {
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/header1.xml",
    "word/footer1.xml",
    "word/footnotes.xml",
    "word/media/image1.png",
}


def _esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _para(text, style=None, numpr=False, ins=False):
    ppr_parts = []
    if style:
        ppr_parts.append('<w:pStyle w:val="%s"/>' % style)
    if numpr:
        ppr_parts.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
    ppr = "<w:pPr>%s</w:pPr>" % "".join(ppr_parts) if ppr_parts else ""
    run = '<w:r><w:t xml:space="preserve">%s</w:t></w:r>' % _esc(text)
    if ins:
        run = '<w:ins w:id="1" w:author="test">%s</w:ins>' % run
    return "<w:p>%s%s</w:p>" % (ppr, run)


def _image_para(rid):
    return (
        "<w:p><w:pPr/><w:r><w:drawing>"
        '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1"/>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="1" name="image1.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>" % rid
    )


def _textbox_para(text):
    return (
        "<w:p><w:r><w:pict>"
        '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" style="width:100pt;height:30pt" '
        'type="#_x0000_t202"><v:textbox>'
        '<w:txbxContent><w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:txbxContent>'
        "</v:textbox></v:shape></w:pict></w:r></w:p>" % _esc(text)
    )


def _table_para(text):
    return (
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tr>"
        "<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>"
        '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
        "</w:tc></w:tr></w:tbl>" % _esc(text)
    )


DOCUMENT_BODY = """\
{p_title}{p_body}{p_list}{p_table}{p_textbox}{p_image}
<w:sectPr>
<w:headerReference w:type="default" r:id="rIdH"/>
<w:footerReference w:type="default" r:id="rIdF"/>
<w:pgSz w:w="11906" w:h="16838"/>
</w:sectPr>"""


def make_fixture(path, tracked_changes=False):
    """生成最小合法 DOCX 夹具。tracked_changes=True 时把“第一段。”包进 <w:ins>。"""
    path = Path(path)
    if tracked_changes:
        body_para = _para("第一段。", ins=True)
    else:
        body_para = _para("第一段。")

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="%s" xmlns:r="%s" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        "<w:body>"
        % (W_NS, R_NS)
        + DOCUMENT_BODY.format(
            p_title=_para("测试标题", style="Title"),
            p_body=body_para,
            p_list=_para("列表项", numpr=True),
            p_table=_table_para("表格文字"),
            p_textbox=_textbox_para("文本框文字"),
            p_image=_image_para("rIdI"),
        )
        + "</w:body></w:document>"
    )

    header_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="%s">%s</w:hdr>' % (W_NS, _para("页眉文字"))
    )
    footer_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:ftr xmlns:w="%s">%s</w:ftr>' % (W_NS, _para("页脚文字"))
    )
    footnotes_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:footnotes xmlns:w="%s">'
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        '<w:footnote w:id="2"><w:p>'
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
        '<w:r><w:t xml:space="preserve">原脚注文字</w:t></w:r>'
        "</w:p></w:footnote>"
        "</w:footnotes>" % W_NS
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="%s">'
        '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>'
        '<w:pPr><w:spacing w:before="240" w:after="60"/></w:pPr>'
        '<w:rPr><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>'
        '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        "</w:styles>" % W_NS
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="%s">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        "</Types>" % CT_NS
    )

    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        "</Relationships>" % REL_NS
    )

    document_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="%s">'
        '<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        '<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
        '<Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>'
        '<Relationship Id="rIdI" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
        "</Relationships>" % REL_NS
    )

    parts = {
        "[Content_Types].xml": content_types_xml,
        "_rels/.rels": root_rels,
        "word/document.xml": document_xml,
        "word/_rels/document.xml.rels": document_rels,
        "word/styles.xml": styles_xml,
        "word/header1.xml": header_xml,
        "word/footer1.xml": footer_xml,
        "word/footnotes.xml": footnotes_xml,
        "word/media/image1.png": PNG_BYTES,
    }
    assert set(parts) == REQUIRED_MEMBERS
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in ("[Content_Types].xml", "_rels/.rels", "word/document.xml",
                     "word/_rels/document.xml.rels", "word/styles.xml",
                     "word/header1.xml", "word/footer1.xml", "word/footnotes.xml",
                     "word/media/image1.png"):
            zf.writestr(name, parts[name])
    return path
