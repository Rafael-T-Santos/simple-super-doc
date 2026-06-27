import { describe, it, expect } from 'vitest'
import { parse, DocxParseError } from '../src/index.js'

describe('parse error paths', () => {
  it('throws INVALID_ZIP on corrupt buffer', async () => {
    const corrupt = new TextEncoder().encode('not a zip file').buffer
    await expect(parse(corrupt)).rejects.toMatchObject({
      name: 'DocxParseError',
      code: 'INVALID_ZIP',
    })
  })

  it('throws MISSING_ENTRY when word/document.xml is absent', async () => {
    // Build a valid zip that is missing word/document.xml
    const JSZip = (await import('jszip')).default
    const zip = new JSZip()
    zip.file('word/styles.xml', '<w:styles/>')
    const buffer = await zip.generateAsync({ type: 'arraybuffer' })
    await expect(parse(buffer)).rejects.toMatchObject({
      name: 'DocxParseError',
      code: 'MISSING_ENTRY',
    })
  })

  it('DocxParseError is an Error instance', () => {
    const err = new DocxParseError('test', 'INVALID_ZIP')
    expect(err).toBeInstanceOf(Error)
    expect(err.code).toBe('INVALID_ZIP')
  })
})
