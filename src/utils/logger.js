/**
 * @fileoverview Logger utility - lightweight structured logging.
 *
 * Provides contextual logging with module names.
 * Respects the PPTX_LOG_LEVEL environment variable.
 *
 * Log levels (lowest to highest severity):
 *   debug → info → warn → error
 *
 * Usage:
 *   const logger = createLogger('MyModule');
 *   logger.debug('Details here');
 *   logger.info('Something happened');
 *   logger.warn('Watch out');
 *   logger.error('Something failed');
 *
 * Environment:
 *   PPTX_LOG_LEVEL=debug  → show all logs
 *   PPTX_LOG_LEVEL=info   → show info, warn, error (default)
 *   PPTX_LOG_LEVEL=warn   → show warn and error only
 *   PPTX_LOG_LEVEL=error  → show only errors
 *   PPTX_LOG_LEVEL=silent → suppress all logs
 */

const LOG_LEVELS = { debug: 0, info: 1, warn: 2, error: 3, silent: 4 };

const currentLevel = LOG_LEVELS[
  (process.env.PPTX_LOG_LEVEL || 'warn').toLowerCase()
] ?? LOG_LEVELS.warn;

/**
 * ANSI color codes for terminal output.
 */
const COLORS = {
  reset: '\x1b[0m',
  dim: '\x1b[2m',
  cyan: '\x1b[36m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  red: '\x1b[31m',
};

/**
 * Formats the current timestamp as HH:MM:SS.mmm
 * @returns {string}
 */
function timestamp() {
  return new Date().toISOString().substring(11, 23);
}

/**
 * @typedef {Object} Logger
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
 */
function createLogger(moduleName) {
  const isTTY = process.stdout.isTTY;

  const log = (level, levelNum, color, message, ...args) => {
    if (levelNum < currentLevel) return;

    const prefix = isTTY
      ? `${COLORS.dim}${timestamp()}${COLORS.reset} ${color}[${level.toUpperCase().padEnd(5)}]${COLORS.reset} ${COLORS.cyan}[${moduleName}]${COLORS.reset}`
      : `${timestamp()} [${level.toUpperCase().padEnd(5)}] [${moduleName}]`;

    const output = args.length > 0 ? `${message} ${args.map(a => JSON.stringify(a, null, 0)).join(' ')}` : message;
    const stream = level === 'error' ? process.stderr : process.stdout;
    stream.write(`${prefix} ${output}\n`);
  };

  return {
    /**
     * Logs a debug-level message. Only shown when PPTX_LOG_LEVEL=debug.
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
  };
}

module.exports = {
  createLogger
};
