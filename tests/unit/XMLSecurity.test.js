import { describe, it, expect } from 'vitest'
const {
  XMLParser,
  validateXml,
  validateXML,
  safeParseXml,
  scanForEntities,
  analyzeXmlFile,
  reportXmlComplexity,
} = require('../../src/index.js')

describe('XML Parsing Security & Large Files', () => {
  const xmlBillionLaughs = `<?xml version="1.0"?>
<!DOCTYPE lolz [
 <!ENTITY lol "lol">
 <!ELEMENT lolz (#PCDATA)>
 <!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
 <!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">
 <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
 <!ENTITY lol4 "&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;">
]>
<lolz>&lol4;</lolz>`

  const xmlXXE = `<?xml version="1.0"?>
<!DOCTYPE foo [  
  <!ELEMENT foo ANY >
  <!ENTITY xxe SYSTEM "file:///etc/passwd" >]>
<foo>&xxe;</foo>`

  const xmlMalformed = `<root><child attr="val">Text without closing tag</root>`

  it('validateXml should detect and reject DTD/DOCTYPE declarations', () => {
    const res = validateXml(xmlBillionLaughs)
    expect(res.valid).toBe(false)
    expect(res.error).toContain('Custom entity declaration detected')
    expect(res.line).toBe(3)
    expect(res.recommendation).toContain('Do not declare custom entities')
  })

  it('validateXml should detect and reject custom ENTITY declarations', () => {
    // Directly test block with ENTITY but no DOCTYPE
    const xmlEntityDirect = `<root><![CDATA[some data]]> <!ENTITY inline "value"></root>`
    const res = validateXml(xmlEntityDirect)
    expect(res.valid).toBe(false)
    expect(res.error).toContain('Custom entity declaration detected')
    expect(res.recommendation).toContain('Do not declare custom entities')
  })

  it('validateXml should detect and reject XXE external references (SYSTEM)', () => {
    const res = validateXml(xmlXXE)
    expect(res.valid).toBe(false)
    expect(res.error).toContain('External reference SYSTEM/PUBLIC detected')
    expect(res.recommendation).toContain('Remove external system/public identifiers')
  })

  it('validateXml should detect and reject XXE external references (PUBLIC)', () => {
    const xmlXXEPublic = `<?xml version="1.0"?><!DOCTYPE test [<!ENTITY xxe PUBLIC "pubid" "http://host/file">]><test>&xxe;</test>`
    const res = validateXml(xmlXXEPublic)
    expect(res.valid).toBe(false)
    expect(res.error).toContain('External reference SYSTEM/PUBLIC detected')
  })

  it('validateXml should flag malformed XML', () => {
    const res = validateXml(xmlMalformed)
    expect(res.valid).toBe(false)
    expect(res.error).toBeDefined()
    expect(res.line).toBe(1)
  })

  it('validateXML compatibility wrapper behaves identically for basic output', () => {
    const res1 = validateXML(xmlBillionLaughs)
    expect(res1.valid).toBe(false)
    expect(res1.error).toContain('Custom entity declaration detected')

    const res2 = validateXML('<root>valid</root>')
    expect(res2.valid).toBe(true)
    expect(res2.error).toBeNull()
  })

  it('safeParseXml should throw on security checks and provide diagnostics', () => {
    try {
      safeParseXml(xmlBillionLaughs, 'slideBillion.xml')
      expect.fail('Should have thrown on Billion Laughs XML')
    } catch (err) {
      expect(err.diagnostic).toBeDefined()
      expect(err.diagnostic.file).toBe('slideBillion.xml')
      expect(err.diagnostic.line).toBe(3)
      expect(err.diagnostic.error).toContain('Custom entity declaration detected')
      expect(err.diagnostic.recommendation).toContain('custom entities')
    }
  })

  it('safeParseXml should throw on malformed XML and report syntax error diagnostics', () => {
    try {
      safeParseXml(xmlMalformed, 'slideBroken.xml')
      expect.fail('Should have thrown on malformed XML')
    } catch (err) {
      expect(err.diagnostic).toBeDefined()
      expect(err.diagnostic.file).toBe('slideBroken.xml')
      expect(err.diagnostic.line).toBe(1)
      expect(err.diagnostic.recommendation).toContain('XML syntax errors')
    }
  })

  it('should parse files with 1000+ standard entities successfully without expansion limit exceeded', () => {
    // Generate a large XML with 1500 standard &amp; entities
    let textNode = 'A'
    for (let i = 0; i < 1500; i++) {
      textNode += ' &amp; B'
    }
    const xmlLarge = `<root>${textNode}</root>`

    const parser = new XMLParser()
    const parsed = parser.parse(xmlLarge)
    expect(parsed).toBeDefined()
    expect(parsed.root).toContain('A & B')
  })

  it('should correctly round-trip standard and numeric entities', () => {
    const xml = `<p:sp attr="value &amp; &lt; &gt; &quot; &apos;"><text>Hello &amp; world &#x41; &#66;</text></p:sp>`
    const parser = new XMLParser()
    const parsed = parser.parse(xml)

    // Unescaped values in JS representation
    expect(parsed['p:sp'][0]['@_attr']).toBe('value & < > " \'')
    expect(parsed['p:sp'][0]['text']).toBe('Hello & world A B')

    const built = parser.build(parsed)
    // Re-encoded in serialized XML representation
    expect(built).toContain('attr="value &amp; &lt; &gt; &quot; &apos;"')
    expect(built).toContain('<text>Hello &amp; world A B</text>') // numeric character references become standard characters, which builder outputs safely or unescaped since they are normal chars
  })

  it('scanForEntities should scan standard, custom, numeric, and hex entities correctly', () => {
    const xml = `<text>&amp; &lt; &lol; &#13; &#x0D; &amp;</text>`
    const stats = scanForEntities(xml)
    expect(stats.total).toBe(6)
    expect(stats.standard).toBe(3) // 2x &amp;, 1x &lt;
    expect(stats.custom).toBe(1) // &lol;
    expect(stats.numeric).toBe(1) // &#13;
    expect(stats.hex).toBe(1) // &#x0D;
    expect(stats.entities).toContain('&lol;')
  })

  it('analyzeXmlFile should count sizing, elements, attributes, and entities', () => {
    const xml = `<p:sp id="1" name="title"><a:t>text &amp;</a:t></p:sp>`
    const analysis = analyzeXmlFile(xml)
    expect(analysis.sizeBytes).toBeGreaterThan(0)
    expect(analysis.elementCount).toBe(2) // p:sp, a:t
    expect(analysis.attributeCount).toBe(2) // id, name
    expect(analysis.entityStats.standard).toBe(1)
  })

  it('reportXmlComplexity should calculate depth, nodeCount, and ratioTextToMarkup', () => {
    const xml = `<p:sp><p:txBody><a:p><a:r><a:t>hello</a:t></a:r></a:p></p:txBody></p:sp>`
    const comp = reportXmlComplexity(xml)
    expect(comp.maxDepth).toBe(5) // p:sp -> p:txBody -> a:p -> a:r -> a:t
    expect(comp.nodeCount).toBe(10) // 5 opening, 5 closing tags
    expect(comp.ratioTextToMarkup).toBeLessThan(0.5)
  })
})
