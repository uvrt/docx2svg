"""Documents with header and footer parts, for the probes of headers, footers and fields.

``probe_docx.package`` relates every extra part from ``word/document.xml``, which is what a
header or footer part needs, but a picture *inside* a header is related from the header
part (``word/_rels/header1.xml.rels``), and ``w:evenAndOddHeaders`` lives in
``settings.xml``.  This module writes both, and keeps ``probe_docx``'s rules: child order
is the schema's, and the face is named in every slot.

The styles are Word's own shape for a document made in Word 16 (``Normal`` with
``docDefaults`` spacing 160 after at 259 auto, and ``Header`` / ``Footer`` with Word's
centre and right tab stops and no space), so the probes measure headers as Word's users
get them, with the face held at Calibri 11 in every slot.
"""

from __future__ import annotations

import io
import struct
import zipfile
import zlib

import probe_docx
import wml

W_NS = probe_docx.W_NS
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
HEADER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
FOOTER_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
HEADER_REL = f"{R_NS}/header"
FOOTER_REL = f"{R_NS}/footer"
IMAGE_REL = f"{R_NS}/image"
STYLES_REL = f"{R_NS}/styles"
SETTINGS_REL = f"{R_NS}/settings"
IMAGE_ID = "rIdImage1"

FACE = {"ascii": "Calibri", "hAnsi": "Calibri", "eastAsia": "Calibri", "cs": "Calibri"}
#: Word 16's ``Header`` and ``Footer`` tab stops for a Letter page with 1-inch margins.
TABS = '<w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/>'


def styles(extra: list[str] | None = None) -> str:
    """Word 16's defaults: Calibri 11 (all four slots), 160 after at 259 auto; ``Header``
    and ``Footer`` based on ``Normal`` with the centre and right stops and no space."""
    story = {"tabs": TABS, "spacing": {"after": 0, "line": 240, "lineRule": "auto"}}
    return wml.styles_part(
        [wml.style("paragraph", "Normal", default=True),
         wml.style("paragraph", "Header", based_on="Normal", ppr_=story),
         wml.style("paragraph", "Footer", based_on="Normal", ppr_=story)] + list(extra or []),
        run_defaults={"rFonts": FACE, "sz": 22, "szCs": 22, "lang": {"val": "en-GB"}},
        paragraph_defaults={"spacing": {"after": 160, "line": 259, "lineRule": "auto"}},
    )


def png(grey: int = 0x80) -> bytes:
    """A 1 x 1 grey PNG."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes([0, grey]))) + chunk(b"IEND", b""))


def picture(number: int, cx: int, cy: int) -> str:
    """An inline picture run showing :data:`IMAGE_ID` (related from the part it is in)."""
    return (
        f'<w:r><w:drawing xmlns:wp="{WP_NS}" xmlns:r="{R_NS}"><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{number}" name="Picture {number}"/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{PIC_NS}"><pic:pic xmlns:pic="{PIC_NS}">'
        f'<pic:nvPicPr><pic:cNvPr id="{number}" name="Picture {number}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{IMAGE_ID}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:inline></w:drawing></w:r>"
    )


def field(instruction: str, result: str | None, *, rpr: dict | None = None, result_rpr: dict | None = None,
          instr_rpr: dict | None = None) -> str:
    """A complex field: ``begin``, the instruction, ``separate`` and ``result`` (none when
    ``None``: no ``separate``), ``end``; each in a run of its own, as Word writes them."""
    base = wml.rpr(**(rpr or {}))
    out = f'<w:r>{base}<w:fldChar w:fldCharType="begin"/></w:r>'
    out += f'<w:r>{wml.rpr(**(instr_rpr or rpr or {}))}<w:instrText xml:space="preserve"> {probe_docx.escape(instruction)} </w:instrText></w:r>'
    if result is not None:
        out += f'<w:r>{base}<w:fldChar w:fldCharType="separate"/></w:r>'
        out += f'<w:r>{wml.rpr(**(result_rpr or rpr or {}))}<w:t xml:space="preserve">{probe_docx.escape(result)}</w:t></w:r>'
    out += f'<w:r>{base}<w:fldChar w:fldCharType="end"/></w:r>'
    return out


def simple_field(instruction: str, result: str, *, rpr: dict | None = None) -> str:
    """A ``w:fldSimple`` holding one result run."""
    return (f'<w:fldSimple w:instr=" {probe_docx.escape(instruction)} ">'
            f'<w:r>{wml.rpr(**(rpr or {}))}<w:t xml:space="preserve">{probe_docx.escape(result)}</w:t></w:r>'
            "</w:fldSimple>")


def story_part(kind: str, blocks: str) -> str:
    root = "hdr" if kind == "header" else "ftr"
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:{root} xmlns:w="{W_NS}" xmlns:r="{R_NS}">{blocks}</w:{root}>')


def section(*, references: str = "", width: int = 12240, height: int = 15840, top: int = 1440,
            bottom: int = 1440, left: int = 1440, right: int = 1440, header: int = 720, footer: int = 720,
            gutter: int = 0, start: str | None = None, title_page: bool = False, v_align: str | None = None,
            page_numbers: dict | None = None, columns: int = 1) -> str:
    """A ``w:sectPr``, children in CT_SectPr order."""
    out = f'<w:sectPr xmlns:r="{R_NS}">{references}'
    if start:
        out += f'<w:type w:val="{start}"/>'
    out += f'<w:pgSz w:w="{width}" w:h="{height}"/>'
    out += (f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}"'
            f' w:header="{header}" w:footer="{footer}" w:gutter="{gutter}"/>')
    if page_numbers:
        out += "<w:pgNumType" + "".join(f' w:{k}="{v}"' for k, v in page_numbers.items()) + "/>"
    out += f'<w:cols w:space="720"{f" w:num={chr(34)}{columns}{chr(34)}" if columns > 1 else ""}/>'
    if v_align:
        out += f'<w:vAlign w:val="{v_align}"/>'
    if title_page:
        out += "<w:titlePg/>"
    out += '<w:docGrid w:linePitch="360"/></w:sectPr>'
    return out


class Parts:
    """Header and footer parts, numbered in the order they are added, and the references
    a ``w:sectPr`` makes to them."""

    def __init__(self) -> None:
        self.parts: list[tuple[str, str, str, bool]] = []

    def add(self, kind: str, blocks: str, *, pictures: bool = False) -> str:
        """Add a part; returns its relationship id."""
        number = sum(1 for part in self.parts if part[0] == kind) + 1
        self.parts.append((kind, f"word/{kind}{number}.xml", story_part(kind, blocks), pictures))
        return f"rIdStory{len(self.parts)}"

    @staticmethod
    def reference(kind: str, which: str, relationship: str) -> str:
        return f'<w:{kind}Reference w:type="{which}" r:id="{relationship}"/>'


def package(body: str, final_section: str, parts: Parts, *, styles_xml: str | None = None,
            compatibility_mode: int | None = None, even_and_odd: bool = False, body_pictures: bool = False) -> bytes:
    """A complete ``.docx``: the body, its header and footer parts, styles and settings."""
    settings_children = ""
    if even_and_odd:
        settings_children += "<w:evenAndOddHeaders/>"
    if compatibility_mode is not None:
        settings_children += (
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word"'
            f' w:val="{compatibility_mode}"/></w:compat>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:settings xmlns:w="{W_NS}">{settings_children}<w:themeFontLang w:val="en-GB"/></w:settings>'
                if settings_children else None)
    document = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{body}{final_section}</w:body></w:document>')
    overrides = ('<Override PartName="/word/document.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '<Override PartName="/word/styles.xml" ContentType="application/'
                 'vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>')
    rels = f'<Relationship Id="rId1" Type="{STYLES_REL}" Target="styles.xml"/>'
    files: dict[str, bytes | str] = {}
    if settings is not None:
        overrides += (f'<Override PartName="/word/settings.xml" ContentType="{wml.SETTINGS_CONTENT_TYPE}"/>')
        rels += f'<Relationship Id="rIdSettings" Type="{SETTINGS_REL}" Target="settings.xml"/>'
        files["word/settings.xml"] = settings
    image_needed = body_pictures
    for index, (kind, name, xml, pictures) in enumerate(parts.parts, start=1):
        content_type = HEADER_CONTENT_TYPE if kind == "header" else FOOTER_CONTENT_TYPE
        overrides += f'<Override PartName="/{name}" ContentType="{content_type}"/>'
        rels += (f'<Relationship Id="rIdStory{index}" Type="{HEADER_REL if kind == "header" else FOOTER_REL}"'
                 f' Target="{name.split("/", 1)[1]}"/>')
        files[name] = xml
        if pictures:
            image_needed = True
            files[f"word/_rels/{name.split('/', 1)[1]}.rels"] = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="{IMAGE_ID}" Type="{IMAGE_REL}" Target="media/image1.png"/></Relationships>')
    if image_needed:
        files["word/media/image1.png"] = png()
        if body_pictures:
            rels += f'<Relationship Id="{IMAGE_ID}" Type="{IMAGE_REL}" Target="media/image1.png"/>'
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Default Extension="png" ContentType="image/png"/>'
                     + overrides + "</Types>")
    out = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        f'<Relationship Id="rId1" Type="{R_NS}/officeDocument" Target="word/document.xml"/>'
                        "</Relationships>"),
        "word/_rels/document.xml.rels": ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                                         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
                                         f'relationships">{rels}</Relationships>'),
        "word/document.xml": document,
        "word/styles.xml": styles_xml if styles_xml is not None else styles(),
    }
    out.update(files)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in out.items():
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, content.encode("utf-8") if isinstance(content, str) else content)
    return buffer.getvalue()
