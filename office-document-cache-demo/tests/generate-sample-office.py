"""Generate tiny DOCX, XLSX and PDF files for the local test fixture."""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


ROOT = Path(__file__).resolve().parents[1] / "demo-files"
ROOT.mkdir(exist_ok=True)


def write_zip(filename, entries):
    with ZipFile(ROOT / filename, "w", ZIP_DEFLATED) as archive:
        for name, contents in entries.items():
            archive.writestr(name, contents.strip())


write_zip("compliance-notice.docx", {
    "[Content_Types].xml": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>""",
    "_rels/.rels": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>""",
    "word/document.xml": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>Compliance notice — local test fixture</w:t></w:r></w:p><w:sectPr/></w:body>
      </w:document>""",
})

write_zip("service-record-template.xlsx", {
    "[Content_Types].xml": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
      </Types>""",
    "_rels/.rels": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>""",
    "xl/workbook.xml": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Service Record" sheetId="1" r:id="rId1"/></sheets>
      </workbook>""",
    "xl/_rels/workbook.xml.rels": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
      </Relationships>""",
    "xl/worksheets/sheet1.xml": """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Service Record</t></is></c></row></sheetData>
      </worksheet>""",
})


def write_pdf(filename):
    """Write a deliberately tiny valid PDF without adding a test dependency."""
    newline = bytes([10])
    stream = b"BT /F1 18 Tf 72 720 Td (Cache usage guide - local test fixture) Tj ET" + newline
    objects = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        b"<< /Length " + str(len(stream)).encode() + b" >>" + newline + b"stream" + newline + stream + b"endstream",
    ]
    chunks = [b"%PDF-1.4" + newline + b"%" + bytes([0xE2, 0xE3, 0xCF, 0xD3]) + newline]
    offsets = [0]
    for number, body in enumerate(objects, start=1):
        offsets.append(sum(len(chunk) for chunk in chunks))
        chunks.append(f"{number} 0 obj".encode() + newline + body + newline + b"endobj" + newline)
    xref_offset = sum(len(chunk) for chunk in chunks)
    chunks.append(f"xref\n0 {len(objects) + 1}".encode() + newline)
    chunks.append(b"0000000000 65535 f " + newline)
    chunks.extend(f"{offset:010d} 00000 n ".encode() + newline for offset in offsets[1:])
    chunks.append(b"trailer" + newline + f"<< /Size {len(objects) + 1} /Root 1 0 R >>".encode() + newline + b"startxref" + newline + str(xref_offset).encode() + newline + b"%%EOF" + newline)
    (ROOT / filename).write_bytes(b"".join(chunks))


write_pdf("cache-usage-guide.pdf")
