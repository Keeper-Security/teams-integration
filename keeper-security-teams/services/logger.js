/**
 * Logger Service
 * 
 * Lightweight structured logging utility with log levels.
 * No external dependencies required.
 */

const LOG_LEVELS = {
  error: 0,
  warn: 1,
  info: 2,
  debug: 3,
};

const LOG_LEVEL = process.env.LOG_LEVEL?.toLowerCase() || 'info';
const CURRENT_LEVEL = LOG_LEVELS[LOG_LEVEL] ?? LOG_LEVELS.info;

/**
 * Format timestamp for log output
 */
function getTimestamp() {
  return new Date().toISOString();
}

/**
 * Format log message with context
 */
function formatMessage(level, context, message, meta) {
  const ts = getTimestamp();
  const prefix = context ? `[${context}]` : '';
  let output = `${ts} [${level.toUpperCase()}] ${prefix} ${message}`;
  
  if (meta !== undefined && meta !== null) {
    if (typeof meta === 'object') {
      try {
        const metaStr = JSON.stringify(meta);
        if (metaStr !== '{}') {
          output += ' ' + metaStr;
        }
      } catch (e) {
        output += ' [Object]';
      }
    } else {
      output += ' ' + meta;
    }
  }
  
  return output;
}

/**
 * Create a logger instance with optional context
 * @param {string} context - Module/component name for log context
 * @returns {Object} Logger instance with error, warn, info, debug methods
 */
function createLogger(context = '') {
  return {
    error(message, meta) {
      if (CURRENT_LEVEL >= LOG_LEVELS.error) {
        console.error(formatMessage('error', context, message, meta));
      }
    },
    
    warn(message, meta) {
      if (CURRENT_LEVEL >= LOG_LEVELS.warn) {
        console.warn(formatMessage('warn', context, message, meta));
      }
    },
    
    info(message, meta) {
      if (CURRENT_LEVEL >= LOG_LEVELS.info) {
        console.log(formatMessage('info', context, message, meta));
      }
    },
    
    debug(message, meta) {
      if (CURRENT_LEVEL >= LOG_LEVELS.debug) {
        console.log(formatMessage('debug', context, message, meta));
      }
    },
  };
}

/**
 * Default logger instance (no context)
 */
const defaultLogger = createLogger();

module.exports = {
  createLogger,
  error: defaultLogger.error,
  warn: defaultLogger.warn,
  info: defaultLogger.info,
  debug: defaultLogger.debug,
  LOG_LEVELS,
};
