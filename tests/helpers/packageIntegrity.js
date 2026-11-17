/**
 * Validates structural integrity of a generated PPTX buffer.
 * Returns an array of error strings; empty array means OK.
 */

import JSZip from 'jszip'
import { XMLParser } from '../../src/parsers/XMLParser.js'

const parser = new XMLParser()

/**
 * @param {Buffer} buffer
 * @returns {Promise<string[]>}
 */
export async function validatePackageIntegrity(buffer) {
  const errors = []
  const zip = await JSZip.loadAsync(buffer)

  const presXml = await zip.file('ppt/presentation.xml')?.async('text')
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('text')
  const ctXml = await zip.file('[Content_Types].xml')?.async('text')

  if (!presXml || !presRelsXml || !ctXml) {
    errors.push('Missing core package parts ([Content_Types].xml, presentation.xml, or rels)')
    return errors
  }

  const presObj = parser.parse(presXml)
  const presRelsObj = parser.parse(presRelsXml)

  const rels = Array.isArray(presRelsObj.Relationships.Relationship)
    ? presRelsObj.Relationships.Relationship
    : [presRelsObj.Relationships.Relationship]

  const slideRels = new Map(
    rels.filter(r => r['@_Type'].endsWith('/slide')).map(r => [r['@_Id'], r['@_Target']])
  )

  const sldIds = presObj['p:presentation']?.['p:sldIdLst']?.['p:sldId']
  if (!sldIds) {
    errors.push('presentation.xml is missing p:sldIdLst')
    return errors
  }

  const sldIdArr = Array.isArray(sldIds) ? sldIds : [sldIds]

  if (sldIdArr.length !== slideRels.size) {
    errors.push(
      `sldIdLst count (${sldIdArr.length}) does not match slide relationships (${slideRels.size})`
    )
  }

  const referencedSlides = new Set()
  for (const sldId of sldIdArr) {
    const rId = sldId['@_r:id']
    const id = sldId['@_id']
    if (!slideRels.has(rId)) {
      errors.push(`sldId ${id} references missing relationship ${rId}`)
      continue
    }
    const target = slideRels.get(rId)
    const slidePath = `ppt/slides/${target.split('/').pop()}`
    referencedSlides.add(slidePath)

    if (!zip.file(slidePath)) {
      errors.push(`sldId ${id} (${rId}) points to missing part ${slidePath}`)
    }
  }

  const slideFiles = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))

  for (const slidePath of slideFiles) {
    if (!referencedSlides.has(slidePath)) {
      errors.push(`Orphan slide part not referenced by presentation: ${slidePath}`)
    }

    const partName = `/${slidePath}`
    if (!ctXml.includes(`PartName="${partName}"`)) {
      errors.push(`Missing [Content_Types].xml override for ${slidePath}`)
    }

    const slideXml = await zip.file(slidePath).async('text')
    const relsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    const slideRelsFile = zip.file(relsPath)
    if (!slideRelsFile) {
      errors.push(`Missing relationships file for ${slidePath}`)
      continue
    }

    const slideRelsObj = parser.parse(await slideRelsFile.async('text'))
    const sRels = Array.isArray(slideRelsObj.Relationships.Relationship)
      ? slideRelsObj.Relationships.Relationship
      : [slideRelsObj.Relationships.Relationship]
    const validIds = new Set(sRels.map(r => r['@_Id']))

    for (const match of slideXml.matchAll(/\br:(?:embed|id|link)="(rId\d+)"/g)) {
      if (!validIds.has(match[1])) {
        errors.push(`${slidePath} references ${match[1]} which is not defined in ${relsPath}`)
      }
    }

    for (const rel of sRels) {
      if (rel['@_TargetMode'] === 'External') continue
      const resolved = resolveRelativeTarget(slidePath, rel['@_Target'])
      if (!zip.file(resolved)) {
        errors.push(`${relsPath} ${rel['@_Id']} points to missing target ${resolved}`)
      }
    }
  }

  // Sequential slide filenames with no gaps
  const expectedCount = sldIdArr.length
  const expectedFiles = Array.from(
    { length: expectedCount },
    (_, i) => `ppt/slides/slide${i + 1}.xml`
  )
  const actualSlideFiles = slideFiles.sort()
  if (JSON.stringify(actualSlideFiles) !== JSON.stringify(expectedFiles)) {
    errors.push(
      `Slide filenames are not sequential: expected ${expectedFiles.join(', ')}, got ${actualSlideFiles.join(', ')}`
    )
  }

  return errors
}

function resolveRelativeTarget(partPath, target) {
  if (target.startsWith('http://') || target.startsWith('https://')) return target
  const baseParts = partPath.split('/')
  baseParts.pop()
  for (const segment of target.split('/')) {
    if (segment === '..') baseParts.pop()
    else if (segment !== '.') baseParts.push(segment)
  }
  return baseParts.join('/')
}
