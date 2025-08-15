/**
 * Production-ready logging utility
 * Automatically manages console output based on environment
 */

import { CONFIG } from './config.js';

class Logger {
  constructor() {
    this.isDevelopment = CONFIG.isDevelopment;
    this.enabledLevels = this.isDevelopment 
      ? ['log', 'info', 'warn', 'error', 'debug'] 
      : ['warn', 'error']; // Only warnings and errors in production
  }

  log(...args) {
    if (this.enabledLevels.includes('log')) {
      console.log(...args);
    }
  }

  info(...args) {
    if (this.enabledLevels.includes('info')) {
      console.info(...args);
    }
  }

  warn(...args) {
    if (this.enabledLevels.includes('warn')) {
      console.warn(...args);
    }
  }

  error(...args) {
    if (this.enabledLevels.includes('error')) {
      console.error(...args);
    }
  }

  debug(...args) {
    if (this.enabledLevels.includes('debug')) {
      console.debug(...args);
    }
  }

  // Special methods for important operations that should always log
  apiCall(method, url, data = null) {
    if (this.enabledLevels.includes('log')) {
      console.log(`🌐 API Call: ${method} ${url}`, data ? { data } : '');
    }
  }

  apiResponse(status, data) {
    if (this.enabledLevels.includes('log')) {
      console.log(`📦 API Response: ${status}`, data);
    }
  }

  apiError(error, context = '') {
    // Always log API errors, even in production
    console.error(`❌ API Error ${context}:`, error);
  }

  auth(message, data = null) {
    if (this.enabledLevels.includes('log')) {
      console.log(`🔐 Auth: ${message}`, data || '');
    }
  }

  credits(message, data = null) {
    if (this.enabledLevels.includes('log')) {
      console.log(`💳 Credits: ${message}`, data || '');
    }
  }
}

// Export a singleton instance
export const logger = new Logger();
export default logger;