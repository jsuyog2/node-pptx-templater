/**
 * @fileoverview CommentManager - Manages OpenXML comments and comment authors.
 *
 * OpenXML Comments Architecture (ISO/IEC 29500 / ECMA-376):
 * ─────────────────────────────────────────────────────────────
 * Presentation-level:
 *   - ppt/commentAuthors.xml
 *     ContentType: application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml
 *     Relationship: http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors (from ppt/presentation.xml)
 *     Contains <p:cmAuthorLst> with <p:cmAuthor id="0" name="..." initials="..." lastIdx="1" clrIdx="0"/>
 *
 * Slide-level:
 *   - ppt/comments/commentN.xml
 *     ContentType: application/vnd.openxmlformats-officedocument.presentationml.comments+xml
 *     Relationship: http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments (from ppt/slides/slideN.xml)
 *     Target in slideN.xml.rels: ../comments/commentN.xml
 *     Contains <p:cmLst> with <p:cm authorId="0" dt="2026-10-17T00:00:00.000" idx="1"><p:pos x="100000" y="100000"/><p:text>Comment</p:text></p:cm>
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { REL_TYPES } = require('./RelationshipManager.js')
const { toEmu } = require('../spatial/units.js')

const logger = createLogger('CommentManager')

const COMMENT_AUTHORS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml'
const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'

class CommentManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /** @private @type {RelationshipManager} */
  #relationshipManager

  /** @private @type {ContentTypesManager} */
  #contentTypesManager

  /** @private @type {ZipManager} */
  #zipManager

  /** @private @type {Object|null} */
  #authorsObj = null

  /**
   * @param {XMLParser} xmlParser
   * @param {RelationshipManager} relationshipManager
   * @param {ContentTypesManager} contentTypesManager
   */
  constructor(xmlParser, relationshipManager, contentTypesManager) {
    this.#xmlParser = xmlParser
    this.#relationshipManager = relationshipManager
    this.#contentTypesManager = contentTypesManager
  }

  /**
   * Initializes with the zipManager reference.
   * @param {ZipManager} zipManager
   */
  initialize(zipManager) {
    this.#zipManager = zipManager
  }

  /**
   * Ensures ppt/commentAuthors.xml is loaded or created, and returns parsed object.
   * @private
   */
  async #getOrCreateAuthorsObj() {
    if (this.#authorsObj) return this.#authorsObj

    const authorsPath = 'ppt/commentAuthors.xml'
    if (this.#zipManager.hasFile(authorsPath)) {
      const xml = await this.#zipManager.readFile(authorsPath)
      this.#authorsObj = this.#xmlParser.parse(xml, 'commentAuthors.xml')
    } else {
      this.#authorsObj = {
        'p:cmAuthorLst': {
          '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
          'p:cmAuthor': [],
        },
      }
    }

    // Ensure p:cmAuthor is an array
    const lst = this.#authorsObj['p:cmAuthorLst']
    if (!lst['p:cmAuthor']) {
      lst['p:cmAuthor'] = []
    } else if (!Array.isArray(lst['p:cmAuthor'])) {
      lst['p:cmAuthor'] = [lst['p:cmAuthor']]
    }

    return this.#authorsObj
  }

  /**
   * Saves ppt/commentAuthors.xml and ensures relationships and content types are registered.
   * @private
   */
  async #flushAuthorsObj() {
    if (!this.#authorsObj) return
    const authorsPath = 'ppt/commentAuthors.xml'
    const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    const xml = this.#xmlParser.build(this.#authorsObj, decl)
    this.#zipManager.writeFile(authorsPath, xml)

    // Register content type
    this.#contentTypesManager.addOverride(authorsPath, COMMENT_AUTHORS_CONTENT_TYPE)

    // Ensure presentation.xml.rels has a relationship to commentAuthors.xml
    const presRels = this.#relationshipManager.getRelationships('ppt/presentation.xml')
    const hasRel = presRels.some(
      r =>
        r.type === REL_TYPES.COMMENT_AUTHORS ||
        r.target === 'commentAuthors.xml' ||
        r.target.endsWith('/commentAuthors.xml')
    )
    if (!hasRel) {
      this.#relationshipManager.addRelationship(
        'ppt/presentation.xml',
        REL_TYPES.COMMENT_AUTHORS,
        'commentAuthors.xml'
      )
      this.#relationshipManager.flushRelationships('ppt/presentation.xml')
    }
  }

  /**
   * Finds or creates an author in commentAuthors.xml.
   * @param {string} [name='Author']
   * @param {string} [initials='A']
   * @returns {Promise<{ id: string, name: string, initials: string, lastIdx: number }>}
   */
  async getOrCreateAuthor(name = 'Author', initials = 'A') {
    const authorsObj = await this.#getOrCreateAuthorsObj()
    const authors = authorsObj['p:cmAuthorLst']['p:cmAuthor']

    let author = authors.find(a => a['@_name'] === name)
    if (!author) {
      const existingIds = authors.map(a => parseInt(a['@_id'] || '0', 10))
      const nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 0
      author = {
        '@_id': String(nextId),
        '@_name': name,
        '@_initials': initials,
        '@_lastIdx': '0',
        '@_clrIdx': String(nextId % 10),
      }
      authors.push(author)
    }

    return {
      id: String(author['@_id']),
      name: author['@_name'],
      initials: author['@_initials'],
      lastIdx: parseInt(author['@_lastIdx'] || '0', 10),
      raw: author,
    }
  }

  /**
   * Resolves or creates comments XML part for a specific slide.
   * @private
   * @param {SlideInfo} slideInfo
   * @returns {Promise<{ commentsZipPath: string, commentsObj: Object }>}
   */
  async #getOrCreateSlideCommentsObj(slideInfo) {
    const slideRels = this.#relationshipManager.getRelationships(slideInfo.zipPath)
    const commentRel = slideRels.find(
      r => r.type === REL_TYPES.COMMENTS || (r.target && r.target.includes('comments/comment'))
    )

    let commentsZipPath
    let commentsObj

    if (commentRel) {
      commentsZipPath = this.#relationshipManager.resolveTarget(
        slideInfo.zipPath,
        commentRel.target
      )
      let xml = this.#zipManager.readCachedFile(commentsZipPath)
      if (!xml) {
        xml = await this.#zipManager.readFile(commentsZipPath)
      }
      if (xml) {
        commentsObj = this.#xmlParser.parse(xml, commentsZipPath.split('/').pop())
      }
    }

    if (!commentsObj) {
      // Allocate next commentN.xml
      let nextNum = 1
      while (this.#zipManager.hasFile(`ppt/comments/comment${nextNum}.xml`)) {
        nextNum++
      }
      commentsZipPath = `ppt/comments/comment${nextNum}.xml`
      const commentFileName = `comment${nextNum}.xml`

      commentsObj = {
        'p:cmLst': {
          '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
          'p:cm': [],
        },
      }

      // Add relationship from slide to comment
      this.#relationshipManager.addRelationship(
        slideInfo.zipPath,
        REL_TYPES.COMMENTS,
        `../comments/${commentFileName}`
      )
      this.#relationshipManager.flushRelationships(slideInfo.zipPath)
      this.#contentTypesManager.addOverride(commentsZipPath, COMMENTS_CONTENT_TYPE)
    }

    const lst = commentsObj['p:cmLst']
    if (!lst['p:cm']) {
      lst['p:cm'] = []
    } else if (!Array.isArray(lst['p:cm'])) {
      lst['p:cm'] = [lst['p:cm']]
    }

    return { commentsZipPath, commentsObj }
  }

  /**
   * Adds a collaborative comment to slide(s).
   *
   * @param {SlideManager} slideManager
   * @param {number[]} slideIndices - 1-based slide indices.
   * @param {string} text - Comment body.
   * @param {Object} [options={}] - Author and positioning details.
   */
  async addComment(slideManager, slideIndices, text, options = {}) {
    if (!text || typeof text !== 'string') {
      throw new PPTXError('addComment requires a non-empty text string')
    }

    const authorName = options.author || 'Author'
    const initials = options.initials || authorName.slice(0, 2).toUpperCase()
    const authorData = await this.getOrCreateAuthor(authorName, initials)

    const now = options.dateTime || new Date().toISOString()
    const xEmu = options.x !== undefined ? toEmu(options.x) : 952500 // ~1 inch
    const yEmu = options.y !== undefined ? toEmu(options.y) : 952500

    for (const slideIndex of slideIndices) {
      const slideInfo = slideManager.getSlideInfo(slideIndex)
      const { commentsZipPath, commentsObj } = await this.#getOrCreateSlideCommentsObj(slideInfo)

      authorData.lastIdx++
      authorData.raw['@_lastIdx'] = String(authorData.lastIdx)

      const commentNode = {
        '@_authorId': authorData.id,
        '@_dt': now,
        '@_idx': String(authorData.lastIdx),
        'p:pos': {
          '@_x': String(Math.round(xEmu)),
          '@_y': String(Math.round(yEmu)),
        },
        'p:text': text,
      }

      commentsObj['p:cmLst']['p:cm'].push(commentNode)

      // Flush comment file
      const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      const xml = this.#xmlParser.build(commentsObj, decl)
      this.#zipManager.writeFile(commentsZipPath, xml)
    }

    await this.#flushAuthorsObj()
  }

  /**
   * Removes a comment by index (1-based), text match, or author filter from slide(s).
   *
   * @param {SlideManager} slideManager
   * @param {number[]} slideIndices
   * @param {number|string} commentIdentifier - 1-based comment position index, comment text substring, or comment idx.
   * @param {Object} [options={}] - Options (e.g. { author: '...' }).
   */
  async removeComment(slideManager, slideIndices, commentIdentifier, options = {}) {
    await this.#getOrCreateAuthorsObj()

    for (const slideIndex of slideIndices) {
      const slideInfo = slideManager.getSlideInfo(slideIndex)
      const slideRels = this.#relationshipManager.getRelationships(slideInfo.zipPath)
      const commentRel = slideRels.find(
        r => r.type === REL_TYPES.COMMENTS || (r.target && r.target.includes('comments/comment'))
      )
      if (!commentRel) continue

      const commentsZipPath = this.#relationshipManager.resolveTarget(
        slideInfo.zipPath,
        commentRel.target
      )
      let xml = this.#zipManager.readCachedFile(commentsZipPath)
      if (!xml) {
        xml = await this.#zipManager.readFile(commentsZipPath)
      }
      if (!xml) continue

      const commentsObj = this.#xmlParser.parse(xml, commentsZipPath.split('/').pop())
      const lst = commentsObj['p:cmLst']
      if (!lst || !lst['p:cm']) continue

      const comments = Array.isArray(lst['p:cm']) ? lst['p:cm'] : [lst['p:cm']]

      let authorIdFilter = null
      if (options.author && this.#authorsObj) {
        const authors = this.#authorsObj['p:cmAuthorLst']['p:cmAuthor']
        const matchingAuthor = authors.find(
          a => a['@_name'] === options.author || a['@_id'] === String(options.author)
        )
        if (matchingAuthor) {
          authorIdFilter = String(matchingAuthor['@_id'])
        }
      }

      const originalLength = comments.length
      if (typeof commentIdentifier === 'number') {
        // 1-based index removal
        const targetIdx = commentIdentifier - 1
        if (targetIdx >= 0 && targetIdx < comments.length) {
          if (!authorIdFilter || comments[targetIdx]['@_authorId'] === authorIdFilter) {
            comments.splice(targetIdx, 1)
          }
        }
      } else if (typeof commentIdentifier === 'string') {
        const query = commentIdentifier.trim().toLowerCase()
        const removeIdx = comments.findIndex(c => {
          if (authorIdFilter && c['@_authorId'] !== authorIdFilter) return false
          const text = String(c['p:text'] || '').toLowerCase()
          const idxAttr = String(c['@_idx'] || '')
          return text.includes(query) || idxAttr === query
        })
        if (removeIdx !== -1) {
          comments.splice(removeIdx, 1)
        }
      }

      if (comments.length !== originalLength) {
        lst['p:cm'] = comments
        const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        const updatedXml = this.#xmlParser.build(commentsObj, decl)
        this.#zipManager.writeFile(commentsZipPath, updatedXml)
        logger.debug(`Removed comment from slide ${slideIndex} at ${commentsZipPath}`)
      }
    }
  }

  /**
   * Retrieves comments on a slide.
   * @param {SlideManager} slideManager
   * @param {number} slideIndex
   * @returns {Promise<Array<{ index: number, author: string, initials: string, text: string, x: number, y: number, dateTime: string }>>}
   */
  async getComments(slideManager, slideIndex) {
    await this.#getOrCreateAuthorsObj()
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideRels = this.#relationshipManager.getRelationships(slideInfo.zipPath)
    const commentRel = slideRels.find(
      r => r.type === REL_TYPES.COMMENTS || (r.target && r.target.includes('comments/comment'))
    )
    if (!commentRel) return []

    const commentsZipPath = this.#relationshipManager.resolveTarget(
      slideInfo.zipPath,
      commentRel.target
    )
    let xml = this.#zipManager.readCachedFile(commentsZipPath)
    if (!xml) {
      xml = await this.#zipManager.readFile(commentsZipPath)
    }
    if (!xml) return []

    const commentsObj = this.#xmlParser.parse(xml, commentsZipPath.split('/').pop())
    const lst = commentsObj['p:cmLst']
    if (!lst || !lst['p:cm']) return []

    const comments = Array.isArray(lst['p:cm']) ? lst['p:cm'] : [lst['p:cm']]
    const authors = this.#authorsObj['p:cmAuthorLst']['p:cmAuthor']

    return comments.map((c, i) => {
      const author = authors.find(a => a['@_id'] === c['@_authorId'])
      return {
        index: i + 1,
        author: author ? author['@_name'] : 'Unknown',
        initials: author ? author['@_initials'] : 'U',
        text: String(c['p:text'] || ''),
        x: parseInt(c['p:pos']?.['@_x'] || '0', 10),
        y: parseInt(c['p:pos']?.['@_y'] || '0', 10),
        dateTime: c['@_dt'] || '',
      }
    })
  }
}

module.exports = { CommentManager, COMMENT_AUTHORS_CONTENT_TYPE, COMMENTS_CONTENT_TYPE }
