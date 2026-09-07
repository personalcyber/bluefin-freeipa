#!/usr/bin/env python3
"""Generate the blank document templates this image seeds into ~/Templates.

Run at container build time with the output directory as the only argument
(see build.sh, "Blank document templates in every user's Templates folder").
Writes five files:

    Text File.txt                   empty
    Markdown Document.md            empty
    Word Document.docx              minimal WordprocessingML package
    Excel Workbook.xlsx             minimal SpreadsheetML package
    PowerPoint Presentation.pptx    minimal PresentationML package

The three OOXML files are built here, from the XML parts spelled out below,
rather than committed to this repo as binaries: an .docx/.xlsx/.pptx is a ZIP
container, so a committed one would be an opaque blob that can't be reviewed
in a diff or regenerated from source. Each one carries only the parts the
format actually requires for a blank document — ECMA-376 Part 1's required
package relationships plus, for PresentationML, the slide master, layout and
theme a presentation cannot omit. LibreOffice (this image's office suite) and
Microsoft Office both open all three; that is verified against LibreOffice
before any change here lands, not assumed.

Output is byte-reproducible: every ZIP entry gets the same fixed timestamp
(the 1980-01-01 epoch ZIP itself uses), so an unchanged template produces an
unchanged file across this image's daily rebuilds, and bootc's /usr diff
stays empty for it.
"""

import sys
import zipfile
from pathlib import Path

# ZIP's own epoch. Anything earlier is unrepresentable in the format.
FIXED_TIMESTAMP = (1980, 1, 1, 0, 0, 0)

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"

# Package-level relationships part, identical in shape for all three formats:
# a single officeDocument relationship naming that format's root part.
ROOT_RELS = f"""<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="{NS_OFFICE_REL}/officeDocument" Target="%s"/>
</Relationships>
"""

# --- WordprocessingML -------------------------------------------------------

DOCX_CONTENT_TYPES = f"""<Types xmlns="{NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>
"""

# One empty paragraph so the document opens with a cursor position, and a
# sectPr giving it US Letter portrait with 1in margins (twentieths of a
# point: 12240x15840 page, 1440 margins).
DOCX_DOCUMENT = """<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p/>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

DOCX_PARTS = {
    "[Content_Types].xml": DOCX_CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS % "word/document.xml",
    "word/document.xml": DOCX_DOCUMENT,
}

# --- SpreadsheetML ----------------------------------------------------------

XLSX_CONTENT_TYPES = f"""<Types xmlns="{NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""

XLSX_WORKBOOK = f"""<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{NS_OFFICE_REL}">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""

XLSX_WORKBOOK_RELS = f"""<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="{NS_OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{NS_OFFICE_REL}/styles" Target="styles.xml"/>
</Relationships>
"""

XLSX_SHEET = """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"""

# styles.xml is optional per the schema but Excel refuses to open a workbook
# whose cells reference style index 0 without one, and writes a "repaired"
# copy on save. The five lists below are the minimum a valid stylesheet has:
# one font, the two mandatory fills (none/gray125), one empty border, and the
# cellStyleXfs/cellXfs pair that index 0 resolves through.
XLSX_STYLES = """<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""

XLSX_PARTS = {
    "[Content_Types].xml": XLSX_CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS % "xl/workbook.xml",
    "xl/workbook.xml": XLSX_WORKBOOK,
    "xl/_rels/workbook.xml.rels": XLSX_WORKBOOK_RELS,
    "xl/worksheets/sheet1.xml": XLSX_SHEET,
    "xl/styles.xml": XLSX_STYLES,
}

# --- PresentationML ---------------------------------------------------------
#
# The wordiest of the three, because a presentation cannot be reduced to just
# its slides: every slide must reference a layout, every layout a master, and
# every master a theme. So a blank deck still needs the full
# presentation -> master -> layout -> slide chain plus theme1.xml.

PPTX_CONTENT_TYPES = f"""<Types xmlns="{NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>
"""

# 16:9 at PowerPoint's default 13.333in x 7.5in (EMU: 914400 per inch).
PPTX_PRESENTATION = f"""<p:presentation xmlns:a="{NS_DRAWING}" xmlns:r="{NS_OFFICE_REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
"""

PPTX_PRESENTATION_RELS = f"""<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="{NS_OFFICE_REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="{NS_OFFICE_REL}/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="{NS_OFFICE_REL}/theme" Target="theme/theme1.xml"/>
</Relationships>
"""

PPTX_SLIDE_MASTER = f"""<p:sldMaster xmlns:a="{NS_DRAWING}" xmlns:r="{NS_OFFICE_REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" hlink="hlink" folHlink="folHlink" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>
"""

PPTX_SLIDE_MASTER_RELS = f"""<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="{NS_OFFICE_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="{NS_OFFICE_REL}/theme" Target="../theme/theme1.xml"/>
</Relationships>
"""

# type="blank": the deck opens on an empty canvas rather than a title
# placeholder waiting to be filled in, which is what a blank template wants.
PPTX_SLIDE_LAYOUT = f"""<p:sldLayout xmlns:a="{NS_DRAWING}" xmlns:r="{NS_OFFICE_REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sldLayout>
"""

PPTX_SLIDE_LAYOUT_RELS = f"""<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="{NS_OFFICE_REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
"""

PPTX_SLIDE = f"""<p:sld xmlns:a="{NS_DRAWING}" xmlns:r="{NS_OFFICE_REL}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sld>
"""

PPTX_SLIDE_RELS = f"""<Relationships xmlns="{NS_REL}">
  <Relationship Id="rId1" Type="{NS_OFFICE_REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
"""


def _theme_fill_style_list() -> str:
    """The three fill / line / effect / background styles a theme must define.

    ECMA-376 requires exactly three entries in each of fillStyleLst,
    lnStyleLst and bgFillStyleLst (subtle / moderate / intense), so they
    cannot simply be left out of a minimal theme. All three here are the
    same flat phClr fill, which is what "no styling" looks like.
    """
    solid = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
    line = (
        '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">'
        f"{solid}<a:prstDash val=\"solid\"/></a:ln>"
    )
    return (
        f"<a:fillStyleLst>{solid * 3}</a:fillStyleLst>"
        f"<a:lnStyleLst>{line * 3}</a:lnStyleLst>"
        "<a:effectStyleLst>"
        f"{'<a:effectStyle><a:effectLst/></a:effectStyle>' * 3}"
        "</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{solid * 3}</a:bgFillStyleLst>"
    )


def _theme_color_scheme() -> str:
    """The twelve theme colors, in the fixed order the schema requires.

    Office's own default (Office Theme) values, so a deck started from this
    template picks up the same accent colors a deck started from PowerPoint's
    blank presentation would.
    """
    accents = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]
    accent_xml = "".join(
        f'<a:accent{i}><a:srgbClr val="{value}"/></a:accent{i}>'
        for i, value in enumerate(accents, start=1)
    )
    return (
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
        '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
        f"{accent_xml}"
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
        "</a:clrScheme>"
    )


PPTX_THEME = (
    f'<a:theme xmlns:a="{NS_DRAWING}" name="Office Theme">'
    "<a:themeElements>"
    f"{_theme_color_scheme()}"
    '<a:fontScheme name="Office">'
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    "</a:fontScheme>"
    f'<a:fmtScheme name="Office">{_theme_fill_style_list()}</a:fmtScheme>'
    "</a:themeElements>"
    "</a:theme>\n"
)

PPTX_PARTS = {
    "[Content_Types].xml": PPTX_CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS % "ppt/presentation.xml",
    "ppt/presentation.xml": PPTX_PRESENTATION,
    "ppt/_rels/presentation.xml.rels": PPTX_PRESENTATION_RELS,
    "ppt/slideMasters/slideMaster1.xml": PPTX_SLIDE_MASTER,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": PPTX_SLIDE_MASTER_RELS,
    "ppt/slideLayouts/slideLayout1.xml": PPTX_SLIDE_LAYOUT,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": PPTX_SLIDE_LAYOUT_RELS,
    "ppt/slides/slide1.xml": PPTX_SLIDE,
    "ppt/slides/_rels/slide1.xml.rels": PPTX_SLIDE_RELS,
    "ppt/theme/theme1.xml": PPTX_THEME,
}

# Plain-text templates are blank on purpose: "New Document > Text File" should
# hand the user an empty file, not one they have to clear out first.
PLAIN_TEMPLATES = {
    "Text File.txt": "",
    "Markdown Document.md": "",
}

OOXML_TEMPLATES = {
    "Word Document.docx": DOCX_PARTS,
    "Excel Workbook.xlsx": XLSX_PARTS,
    "PowerPoint Presentation.pptx": PPTX_PARTS,
}


def write_ooxml(path: Path, parts: dict[str, str]) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, body in parts.items():
            # [Content_Types].xml is expected first by some readers; dicts
            # preserve insertion order, and it is first in every parts map.
            info = zipfile.ZipInfo(name, date_time=FIXED_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            content = body if body.startswith("<?xml") else XML_DECL + body
            package.writestr(info, content.encode("utf-8"))


def main() -> int:
    if len(sys.argv) != 2:
        print(f"usage: {sys.argv[0]} OUTPUT_DIR", file=sys.stderr)
        return 2

    output_dir = Path(sys.argv[1])
    output_dir.mkdir(parents=True, exist_ok=True)

    for name, content in PLAIN_TEMPLATES.items():
        (output_dir / name).write_text(content, encoding="utf-8")

    for name, parts in OOXML_TEMPLATES.items():
        write_ooxml(output_dir / name, parts)

    for name in list(PLAIN_TEMPLATES) + list(OOXML_TEMPLATES):
        (output_dir / name).chmod(0o644)

    return 0


if __name__ == "__main__":
    sys.exit(main())
