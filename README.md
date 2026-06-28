# simple-super-doc

Browser-only TypeScript library that parses a `.docx` file into a typed
intermediate representation (IR) and renders it to HTML. Built for document
editors and viewers that need a faithful, inspectable model of the document.

- **Typed IR** — `parse()` returns a `DocxDocument` of `Block`s (paragraphs,
  tables) with computed styles, lists, images, and page size. Inspect or
  transform it before rendering.
- **HTML rendering** — `render()` paints the IR into a container element.
- **No DOM at parse time** — parsing is pure; only `render()` touches the DOM.
- **Security** — text is written with `textContent`, never `innerHTML`.

## Install

```bash
npm install simple-super-doc
```

## Usage

```ts
import { parse, render } from 'simple-super-doc'

const buf = await file.arrayBuffer()       // a .docx File/Blob
const doc = await parse(buf)               // typed DocxDocument IR
render(doc, document.getElementById('view')!)
```

`parse` throws `DocxParseError` (`code: 'INVALID_ZIP' | 'MISSING_ENTRY'`) for
non-docx or malformed input.

## What it covers

Paragraph and run styles (bold/italic/underline, font size/family, color,
alignment, shading), the style cascade (docDefaults → named style → direct
formatting), numbered and bulleted lists, tables with `gridSpan`/`vMerge`,
inline and anchored images (as base64 data URLs), and document block ordering
recovered from the raw XML (so paragraphs and tables keep their real sequence).

## Page-aware rendering

Some templates paginate themselves with full-page background images (a cover
slide, framed body pages, a closing slide) but carry **no explicit page
breaks** — no `<w:br w:type="page"/>` and no `<w:lastRenderedPageBreak/>`.
Word/LibreOffice decide the pages purely in their layout engine.

When `render()` detects a full-page background (a `behindDoc=1` anchor), it
reconstructs the pages with a two-pass approach and a set of heuristics. These
live in [`src/renderer/layout.ts`](src/renderer/layout.ts) as pure, unit-tested
functions ([`test/layout-heuristics.test.ts`](test/layout-heuristics.test.ts)):

- **Two-pass measurement** — each block is measured in a hidden container
  (Pass 1), then content is distributed into page-sized boxes (Pass 2). This
  needs a real layout engine, so it runs in the browser, not jsdom.
- **Empty paragraphs are spacers** — they collapse to 0px in a browser but
  occupy a line in Word, so empty paragraphs get a min line height
  (`EMPTY_LINE_EM`). Their vertical rhythm positions things like a cover's
  customer-name field.
- **Section headings force a new page** — a heading is large text
  (`fontSize >= 24`) or a wide-short "text-as-image" banner; a small icon that
  precedes a text heading travels with it. Each framed section starts a page.
- **Watermarks overlay, not flow** — large floating images (decorative
  frames/watermarks) are absolutely positioned behind the text and contribute
  zero pagination height, instead of pushing content onto extra pages.
- **Background regions by image order** — floating backgrounds are positioned
  by absolute page coordinates, so the *order* of distinct background images
  drives which page each covers (cover → page 0, body frame → pages 1+), not
  their XML position.
- **Full-bleed slides** — a page whose only content is a near-full-width image
  is drawn edge-to-edge; blank pages are dropped.

These thresholds are tuned for proposal/letter-style templates. A document
without a full-page background skips all of this and renders as a normal flow.

## Development

```bash
npm run typecheck   # tsc --noEmit
npm test            # vitest
npm run build       # esm + d.ts into dist/
npm run build:demo  # IIFE bundle into demo/ for the drag-and-drop demo
```

## License

MIT
