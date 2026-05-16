/**
 * @fileoverview Custom error classes for node-pptx-templater.
 *
 * All errors extend from PPTXError to allow consumers to catch all
 * library errors with a single catch clause:
 *
 *   try {
 *     await ppt.saveToFile('./out.pptx');
 *   } catch (err) {
 *     if (err instanceof PPTXError) {
 *       // Handle all node-pptx-templater errors
 *     }
 *   }
 */

/**
 * @class PPTXError
 * @description Base error class for all node-pptx-templater errors.
 * @extends Error
 */
export class PPTXError extends Error {
  /**
   * @param {string} message - Human-readable error description.
   * @param {Error} [cause] - Original underlying error (for error chaining).
   */
  constructor(message, cause) {
    super(message);
    this.name = 'PPTXError';
    if (cause) this.cause = cause;

    // Maintains proper stack trace for where error was thrown (V8 only)
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, this.constructor);
    }
  }
}

/**
 * @class SlideNotFoundError
 * @description Thrown when a slide reference cannot be resolved.
 * @extends PPTXError
 */
export class SlideNotFoundError extends PPTXError {
  /**
   * @param {string} message
   */
  constructor(message) {
    super(message);
    this.name = 'SlideNotFoundError';
  }
}

/**
 * @class ChartNotFoundError
 * @description Thrown when a chart cannot be located in a slide.
 * @extends PPTXError
 */
export class ChartNotFoundError extends PPTXError {
  /**
   * @param {string} message
   */
  constructor(message) {
    super(message);
    this.name = 'ChartNotFoundError';
  }
}

/**
 * @class TableNotFoundError
 * @description Thrown when a table cannot be located in a slide.
 * @extends PPTXError
 */
export class TableNotFoundError extends PPTXError {
  /**
   * @param {string} message
   */
  constructor(message) {
    super(message);
    this.name = 'TableNotFoundError';
  }
}

/**
 * @class XMLParseError
 * @description Thrown when XML parsing or building fails.
 * @extends PPTXError
 */
export class XMLParseError extends PPTXError {
  /**
   * @param {string} message
   * @param {Error} [cause]
   */
  constructor(message, cause) {
    super(message, cause);
    this.name = 'XMLParseError';
  }
}

/**
 * @class InvalidTemplateError
 * @description Thrown when a PPTX file has an invalid or unsupported structure.
 * @extends PPTXError
 */
export class InvalidTemplateError extends PPTXError {
  /**
   * @param {string} message
   * @param {Error} [cause]
   */
  constructor(message, cause) {
    super(message, cause);
    this.name = 'InvalidTemplateError';
  }
}

/**
 * @class MediaEmbedError
 * @description Thrown when a media file cannot be embedded.
 * @extends PPTXError
 */
export class MediaEmbedError extends PPTXError {
  /**
   * @param {string} message
   * @param {Error} [cause]
   */
  constructor(message, cause) {
    super(message, cause);
    this.name = 'MediaEmbedError';
  }
}
