/**
 * @fileoverview Logger utility - lightweight structured logging.
 *
 * Provides contextual logging with module names.
 * Respects the PPTX_LOG_LEVEL environment variable, or can be
 * configured at runtime via setGlobalLogLevel().
 *
 * Log levels (lowest to highest severity):
 *   verbose → debug → info → warn → error → silent
 *
 * Usage:
 *   const logger = createLogger('MyModule');
 *   logger.debug('Details here');
 *   logger.info('Something happened');
 *   logger.warn('Watch out');
 *   logger.error('Something failed');
 *
 * Environment variable (set before process starts):
 *   PPTX_LOG_LEVEL=debug   → show debug, info, warn, error
 *   PPTX_LOG_LEVEL=info    → show info, warn, error
 *   PPTX_LOG_LEVEL=warn    → show warn and error only (default)
 *   PPTX_LOG_LEVEL=error   → show only errors
 *   PPTX_LOG_LEVEL=silent  → suppress all output
 *
 * Runtime control (overrides env var):
 *   const { setGlobalLogLevel } = require('./logger');
 *   setGlobalLogLevel('debug');
 */

/** @type {Object.<string, number>} */
const LOG_LEVELS = { verbose: -1, debug: 0, info: 1, warn: 2, error: 3, silent: 4 }

/** @type {number} Initial level from environment variable */
const envLevel = LOG_LEVELS[(process.env.PPTX_LOG_LEVEL || 'warn').toLowerCase()] ?? LOG_LEVELS.warn

/** @type {number|null} Runtime override — null means use envLevel */
let runtimeLevel = null

/**
 * Gets the current effective log level.
 * @returns {number}
 */
function getEffectiveLevel() {
  return runtimeLevel !== null ? runtimeLevel : envLevel
}

/**
 * Sets the global log level at runtime, overriding the environment variable.
 * This affects all logger instances immediately.
 *
 * @param {string} level - One of: 'verbose', 'debug', 'info', 'warn', 'error', 'silent'
 * @throws {Error} If an invalid level is provided.
 *
 * @example
 * const { setGlobalLogLevel } = require('node-pptx-templater');
 * setGlobalLogLevel('debug'); // Enable verbose output
 * setGlobalLogLevel('silent'); // Suppress everything
 */
function setGlobalLogLevel(level) {
  const normalized = String(level).toLowerCase()
  if (LOG_LEVELS[normalized] === undefined) {
    throw new Error(
      `Invalid log level: "${level}". Valid levels: verbose, debug, info, warn, error, silent`
    )
  }
  runtimeLevel = LOG_LEVELS[normalized]
}

/**
 * Resets the log level back to the environment variable default.
 */
function resetLogLevel() {
  runtimeLevel = null
}

/**
 * ANSI color codes for terminal output.
 * @private
 */
const COLORS = {
  reset: '\x1b[0m',
  dim: '\x1b[2m',
  cyan: '\x1b[36m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  red: '\x1b[31m',
  magenta: '\x1b[35m',
}

/**
 * Formats the current timestamp as HH:MM:SS.mmm
 * @returns {string}
 */
function timestamp() {
  return new Date().toISOString().substring(11, 23)
}

/**
 * @typedef {Object} Logger
 * @property {Function} verbose - Log verbose message (most detailed).
 * @property {Function} debug - Log debug message.
 * @property {Function} info - Log info message.
 * @property {Function} warn - Log warning.
 * @property {Function} error - Log error.
 */

/**
 * Creates a named logger instance.
 *
 * @param {string} moduleName - Name of the module (shown in log output).
 * @returns {Logger}
 *
 * @example
 * const logger = createLogger('SlideManager');
 * logger.info('Loaded 5 slides');
 * logger.debug('Processing slide XML...');
 */
function createLogger(moduleName) {
  const isTTY = process.stdout.isTTY

  const log = (level, levelNum, color, message, ...args) => {
    if (levelNum < getEffectiveLevel()) return

    const prefix = isTTY
      ? `${COLORS.dim}${timestamp()}${COLORS.reset} ${color}[${level.toUpperCase().padEnd(7)}]${COLORS.reset} ${COLORS.cyan}[${moduleName}]${COLORS.reset}`
      : `${timestamp()} [${level.toUpperCase().padEnd(7)}] [${moduleName}]`

    const output =
      args.length > 0
        ? `${message} ${args.map(a => JSON.stringify(a, null, 0)).join(' ')}`
        : message
    const stream = level === 'error' ? process.stderr : process.stdout
    stream.write(`${prefix} ${output}\n`)
  }

  return {
    /**
     * Logs a verbose-level message. Only shown when level=verbose.
     * @param {string} message
     * @param {...*} args
     */
    verbose: (message, ...args) =>
      log('verbose', LOG_LEVELS.verbose, COLORS.magenta, message, ...args),

    /**
     * Logs a debug-level message. Only shown when level=debug or lower.
     * @param {string} message
     * @param {...*} args
     */
    debug: (message, ...args) => log('debug', LOG_LEVELS.debug, COLORS.dim, message, ...args),

    /**
     * Logs an info-level message.
     * @param {string} message
     * @param {...*} args
     */
    info: (message, ...args) => log('info', LOG_LEVELS.info, COLORS.green, message, ...args),

    /**
     * Logs a warning.
     * @param {string} message
     * @param {...*} args
     */
    warn: (message, ...args) => log('warn', LOG_LEVELS.warn, COLORS.yellow, message, ...args),

    /**
     * Logs an error.
     * @param {string} message
     * @param {...*} args
     */
    error: (message, ...args) => log('error', LOG_LEVELS.error, COLORS.red, message, ...args),
  }
}

module.exports = {
  createLogger,
  setGlobalLogLevel,
  resetLogLevel,
  LOG_LEVELS,
}
