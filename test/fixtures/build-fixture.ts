/**
 * Generates test/fixtures/sample.docx — a minimal but feature-complete OOXML
 * document used as the v0.1.0 integration fixture. Run once via:
 *   npx tsx test/fixtures/build-fixture.ts
 */
import JSZip from 'jszip'
import { writeFileSync } from 'fs'
import { dirname, join } from 'path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`

const RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

const WORD_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
        <w:sz w:val="24"/>
        <w:color w:val="auto"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:jc w:val="left"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="40"/>
      <w:color w:val="1F3864"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
      <w:color w:val="2E74B5"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Strong">
    <w:name w:val="Strong"/>
    <w:rPr>
      <w:b/>
    </w:rPr>
  </w:style>
</w:styles>`

const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
    </w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>`

const DOCUMENT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>

    <!-- Title -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>simple-super-doc Integration Test</w:t></w:r>
    </w:p>

    <!-- Section heading -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Mixed Runs</w:t></w:r>
    </w:p>

    <!-- Paragraph with bold, italic, underline, color -->
    <w:p>
      <w:r><w:t xml:space="preserve">This is </w:t></w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>bold</w:t>
      </w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>italic</w:t>
      </w:r>
      <w:r><w:t xml:space="preserve">, </w:t></w:r>
      <w:r>
        <w:rPr><w:u w:val="single"/></w:rPr>
        <w:t>underlined</w:t>
      </w:r>
      <w:r><w:t xml:space="preserve">, and </w:t></w:r>
      <w:r>
        <w:rPr><w:color w:val="C00000"/></w:rPr>
        <w:t>red text</w:t>
      </w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>

    <!-- Explicit bold=false override (should not be bold) -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r>
        <w:rPr><w:b w:val="0"/></w:rPr>
        <w:t>Not bold despite Heading1 style</w:t>
      </w:r>
    </w:p>

    <!-- Section heading -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Bullet List</w:t></w:r>
    </w:p>

    <!-- Bullet list (numId=1) -->
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>First item</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>Second item with </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold text</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>Third item</w:t></w:r>
    </w:p>

    <!-- Section heading -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Ordered List</w:t></w:r>
    </w:p>

    <!-- Ordered list (numId=2) -->
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>
      </w:pPr>
      <w:r><w:t>Step one</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>
      </w:pPr>
      <w:r><w:t>Step two</w:t></w:r>
    </w:p>

    <!-- Section heading -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t>Table with Merged Cells</w:t></w:r>
    </w:p>

    <!-- 3x3 table: col0 rows 0-1 vMerge, col1 row0 colSpan=2 -->
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="5000" w:type="pct"/>
      </w:tblPr>
      <w:tblGrid>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
        <w:gridCol w:w="2000"/>
      </w:tblGrid>

      <!-- Row 0 -->
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:vMerge w:val="restart"/>
          </w:tcPr>
          <w:p><w:r><w:t>Merged (rows 0-1)</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr>
            <w:gridSpan w:val="2"/>
          </w:tcPr>
          <w:p><w:r><w:t>Col-span 2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>

      <!-- Row 1 -->
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:vMerge/>
          </w:tcPr>
          <w:p><w:r><w:t></w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B2</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>C2</w:t></w:r></w:p>
        </w:tc>
      </w:tr>

      <!-- Row 2 -->
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>A3</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>B3</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>C3</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>

    <!-- Tracked change (w:ins) — runs must be processed -->
    <w:p>
      <w:ins w:id="1" w:author="Test" w:date="2026-01-01T00:00:00Z">
        <w:r><w:t>This text was inserted via tracked change.</w:t></w:r>
      </w:ins>
    </w:p>

    <!-- Paragraph using color=auto (should produce no color in IR) -->
    <w:p>
      <w:r>
        <w:rPr><w:color w:val="auto"/></w:rPr>
        <w:t>Auto color paragraph — no color in IR.</w:t>
      </w:r>
    </w:p>

  </w:body>
</w:document>`

async function main() {
  const zip = new JSZip()

  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', RELS)
  zip.file('word/_rels/document.xml.rels', WORD_RELS)
  zip.file('word/styles.xml', STYLES)
  zip.file('word/numbering.xml', NUMBERING)
  zip.file('word/document.xml', DOCUMENT)

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  const out = join(__dirname, 'sample.docx')
  writeFileSync(out, buf)
  console.log(`Written: ${out} (${buf.length} bytes)`)
}

main().catch(e => { console.error(e); process.exit(1) })
