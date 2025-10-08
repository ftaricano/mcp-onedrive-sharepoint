/**
 * Security validation utilities for file paths and inputs
 * Prevents path traversal, injection attacks, and validates inputs
 */

import { normalize, resolve, isAbsolute } from 'path';

export interface ValidationResult {
  isValid: boolean;
  error?: string;
  sanitized?: string;
}

export class SecurityValidator {
  
  // Forbidden patterns for security
  private static readonly FORBIDDEN_PATTERNS = [
    /\.\./g,           // Directory traversal
    /~/g,              // Home directory
    /\0/g,             // Null bytes
    /[\x00-\x1f]/g,    // Control characters
    /<script/gi,       // Script injection
    /javascript:/gi,   // Javascript injection
    /data:/gi,         // Data URLs
    /vbscript:/gi,     // VBScript injection
    /on\w+\s*=/gi,     // Event handlers
  ];

  // Forbidden file extensions
  private static readonly FORBIDDEN_EXTENSIONS = [
    '.exe', '.bat', '.cmd', '.com', '.pif', '.scr', '.vbs', '.js', '.jar',
    '.ps1', '.psm1', '.psd1', '.ps1xml', '.psc1', '.app', '.deb', '.pkg',
    '.dmg', '.msi', '.run', '.bin', '.sh', '.bash', '.zsh', '.fish'
  ];

  // Maximum path length
  private static readonly MAX_PATH_LENGTH = 400;
  private static readonly MAX_FILE_NAME_LENGTH = 255;

  /**
   * Validate file path for security issues
   */
  static validatePath(path: string): ValidationResult {
    if (!path || typeof path !== 'string') {
      return {
        isValid: false,
        error: 'Path must be a non-empty string'
      };
    }

    // Check length limits
    if (path.length > this.MAX_PATH_LENGTH) {
      return {
        isValid: false,
        error: `Path too long (max ${this.MAX_PATH_LENGTH} characters)`
      };
    }

    // Check for forbidden patterns
    for (const pattern of this.FORBIDDEN_PATTERNS) {
      if (pattern.test(path)) {
        return {
          isValid: false,
          error: 'Path contains forbidden patterns'
        };
      }
    }

    // Check for absolute paths (OneDrive uses relative paths)
    if (isAbsolute(path)) {
      return {
        isValid: false,
        error: 'Absolute paths are not allowed'
      };
    }

    // Normalize and check again
    const normalized = normalize(path);
    if (normalized !== path && normalized.includes('..')) {
      return {
        isValid: false,
        error: 'Path traversal detected after normalization'
      };
    }

    return {
      isValid: true,
      sanitized: normalized
    };
  }

  /**
   * Validate file name for security and compliance
   */
  static validateFileName(fileName: string): ValidationResult {
    if (!fileName || typeof fileName !== 'string') {
      return {
        isValid: false,
        error: 'File name must be a non-empty string'
      };
    }

    // Check length
    if (fileName.length > this.MAX_FILE_NAME_LENGTH) {
      return {
        isValid: false,
        error: `File name too long (max ${this.MAX_FILE_NAME_LENGTH} characters)`
      };
    }

    // Check for Windows reserved names
    const reservedNames = [
      'CON', 'PRN', 'AUX', 'NUL',
      'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
      'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
    ];

    const nameWithoutExt = fileName.split('.')[0].toUpperCase();
    if (reservedNames.includes(nameWithoutExt)) {
      return {
        isValid: false,
        error: 'File name is a reserved system name'
      };
    }

    // Check for forbidden characters
    const forbiddenChars = /[<>:"|?*\x00-\x1f]/g;
    if (forbiddenChars.test(fileName)) {
      return {
        isValid: false,
        error: 'File name contains forbidden characters'
      };
    }

    // Check for dangerous extensions
    const extension = fileName.toLowerCase().substring(fileName.lastIndexOf('.'));
    if (this.FORBIDDEN_EXTENSIONS.includes(extension)) {
      return {
        isValid: false,
        error: 'File extension is not allowed for security reasons'
      };
    }

    return {
      isValid: true,
      sanitized: fileName.trim()
    };
  }

  /**
   * Validate search query for injection attacks
   */
  static validateSearchQuery(query: string): ValidationResult {
    if (!query || typeof query !== 'string') {
      return {
        isValid: false,
        error: 'Search query must be a non-empty string'
      };
    }

    // Check length
    if (query.length > 500) {
      return {
        isValid: false,
        error: 'Search query too long (max 500 characters)'
      };
    }

    // Check for script injection attempts
    const scriptPatterns = [
      /<script/gi,
      /javascript:/gi,
      /vbscript:/gi,
      /on\w+\s*=/gi,
      /expression\s*\(/gi,
      /eval\s*\(/gi,
      /setTimeout/gi,
      /setInterval/gi
    ];

    for (const pattern of scriptPatterns) {
      if (pattern.test(query)) {
        return {
          isValid: false,
          error: 'Search query contains potentially malicious content'
        };
      }
    }

    // Remove excessive whitespace and sanitize
    const sanitized = query
      .replace(/\s+/g, ' ')
      .trim()
      .substring(0, 500);

    return {
      isValid: true,
      sanitized
    };
  }

  /**
   * Validate file size for uploads
   */
  static validateFileSize(size: number, maxSize: number = 100 * 1024 * 1024): ValidationResult {
    if (typeof size !== 'number' || size < 0) {
      return {
        isValid: false,
        error: 'File size must be a positive number'
      };
    }

    if (size === 0) {
      return {
        isValid: false,
        error: 'File cannot be empty'
      };
    }

    if (size > maxSize) {
      const maxSizeMB = Math.round(maxSize / (1024 * 1024));
      return {
        isValid: false,
        error: `File too large (max ${maxSizeMB}MB)`
      };
    }

    return { isValid: true };
  }

  /**
   * Validate URL for redirection attacks
   */
  static validateUrl(url: string): ValidationResult {
    if (!url || typeof url !== 'string') {
      return {
        isValid: false,
        error: 'URL must be a non-empty string'
      };
    }

    try {
      const parsedUrl = new URL(url);
      
      // Only allow HTTPS for external URLs
      if (!['https:', 'http:'].includes(parsedUrl.protocol)) {
        return {
          isValid: false,
          error: 'Only HTTP/HTTPS URLs are allowed'
        };
      }

      // Block localhost and private IP ranges
      const hostname = parsedUrl.hostname.toLowerCase();
      const privatePatterns = [
        /^localhost$/,
        /^127\./,
        /^10\./,
        /^172\.(1[6-9]|2[0-9]|3[01])\./,
        /^192\.168\./,
        /^169\.254\./,
        /^::1$/,
        /^fc00::/,
        /^fe80::/
      ];

      for (const pattern of privatePatterns) {
        if (pattern.test(hostname)) {
          return {
            isValid: false,
            error: 'Private IP addresses and localhost are not allowed'
          };
        }
      }

      return { isValid: true };
    } catch (error) {
      return {
        isValid: false,
        error: 'Invalid URL format'
      };
    }
  }

  /**
   * Sanitize user input for logging
   */
  static sanitizeForLogging(input: any): string {
    if (input === null || input === undefined) {
      return 'null';
    }

    let str = typeof input === 'string' ? input : JSON.stringify(input);
    
    // Remove potential secrets
    str = str.replace(/("(?:password|token|key|secret|auth)[^"]*":\s*")[^"]*"/gi, '$1[REDACTED]"');
    str = str.replace(/Bearer\s+[A-Za-z0-9\-._~+/]+/gi, 'Bearer [REDACTED]');
    
    // Remove HTML/script tags for security
    str = str.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '[SCRIPT_REMOVED]');
    str = str.replace(/<[^>]+>/g, '[HTML_TAG_REMOVED]');
    
    // Remove javascript: and data: URLs
    str = str.replace(/javascript:[^;]*/gi, '[JS_URL_REMOVED]');
    str = str.replace(/data:[^;]*/gi, '[DATA_URL_REMOVED]');
    
    // Remove event handlers
    str = str.replace(/on\w+\s*=\s*[^>\s]*/gi, '[EVENT_HANDLER_REMOVED]');
    
    // Limit length
    if (str.length > 1000) {
      str = str.substring(0, 1000) + '... [TRUNCATED]';
    }

    return str;
  }

  /**
   * Validate OData query parameters
   */
  static validateODataQuery(query: Record<string, any>): ValidationResult {
    const allowedParams = [
      '$select', '$filter', '$orderby', '$top', '$skip', 
      '$count', '$expand', '$search', '$format'
    ];

    for (const key of Object.keys(query)) {
      if (!allowedParams.includes(key)) {
        return {
          isValid: false,
          error: `Parameter '${key}' is not allowed in OData queries`
        };
      }

      const value = query[key];
      if (typeof value === 'string') {
        // Check for injection attempts in OData values
        const dangerousPatterns = [
          /;\s*(drop|delete|update|insert|create|alter|exec|execute)/gi,
          /union\s+select/gi,
          /'\s*or\s*'1'\s*=\s*'1/gi,
          /--/g,
          /\/\*/g
        ];

        for (const pattern of dangerousPatterns) {
          if (pattern.test(value)) {
            return {
              isValid: false,
              error: 'OData query contains potentially malicious content'
            };
          }
        }
      }
    }

    return { isValid: true };
  }
}

/**
 * Path helper utilities with security validation
 */
export class SecurePath {
  
  /**
   * Safely join path components
   */
  static join(...components: string[]): ValidationResult {
    const cleanComponents = components
      .filter(c => c && typeof c === 'string')
      .map(c => c.trim())
      .filter(c => c.length > 0);

    if (cleanComponents.length === 0) {
      return {
        isValid: false,
        error: 'No valid path components provided'
      };
    }

    // Validate each component
    for (const component of cleanComponents) {
      const validation = SecurityValidator.validatePath(component);
      if (!validation.isValid) {
        return validation;
      }
    }

    const joined = cleanComponents.join('/');
    return SecurityValidator.validatePath(joined);
  }

  /**
   * Extract safe file name from path
   */
  static extractFileName(path: string): ValidationResult {
    const validation = SecurityValidator.validatePath(path);
    if (!validation.isValid) {
      return validation;
    }

    const fileName = path.split('/').pop() || '';
    return SecurityValidator.validateFileName(fileName);
  }

  /**
   * Get parent directory safely
   */
  static getParentDir(path: string): ValidationResult {
    const validation = SecurityValidator.validatePath(path);
    if (!validation.isValid) {
      return validation;
    }

    const parts = path.split('/');
    if (parts.length <= 1) {
      return {
        isValid: true,
        sanitized: ''
      };
    }

    const parentPath = parts.slice(0, -1).join('/');
    return SecurityValidator.validatePath(parentPath);
  }
}

/**
 * Audit logger with secure input sanitization
 */
export class AuditLogger {
  private static logEntries: Array<{
    timestamp: string;
    operation: string;
    user: string;
    resource: string;
    result: 'success' | 'failure';
    details?: string;
  }> = [];

  static log(
    operation: string,
    user: string,
    resource: string,
    result: 'success' | 'failure',
    details?: any
  ): void {
    const entry = {
      timestamp: new Date().toISOString(),
      operation: SecurityValidator.sanitizeForLogging(operation),
      user: SecurityValidator.sanitizeForLogging(user),
      resource: SecurityValidator.sanitizeForLogging(resource),
      result,
      details: details ? SecurityValidator.sanitizeForLogging(details) : undefined
    };

    this.logEntries.push(entry);
    
    // Keep only last 1000 entries
    if (this.logEntries.length > 1000) {
      this.logEntries = this.logEntries.slice(-1000);
    }

    // Log to console for development
    console.log(`[AUDIT] ${entry.timestamp} | ${entry.operation} | ${entry.user} | ${entry.resource} | ${entry.result}`);
  }

  static getRecentLogs(limit: number = 100): typeof AuditLogger.logEntries {
    return this.logEntries.slice(-limit);
  }

  static clearLogs(): void {
    this.logEntries = [];
  }
}