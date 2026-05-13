/**
 * Input validators and sanitizers
 * Enhanced validation for all tool inputs
 */

import * as path from "path";
import * as fs from "fs";

/**
 * Validates and sanitizes file paths
 */
export class PathValidator {
  private static readonly INVALID_CHARS = /[<>:"|?*]/g;
  private static readonly MAX_PATH_LENGTH = 255;
  private static readonly RESERVED_NAMES = [
    "CON",
    "PRN",
    "AUX",
    "NUL",
    "COM1",
    "COM2",
    "COM3",
    "COM4",
    "COM5",
    "COM6",
    "COM7",
    "COM8",
    "COM9",
    "LPT1",
    "LPT2",
    "LPT3",
    "LPT4",
    "LPT5",
    "LPT6",
    "LPT7",
    "LPT8",
    "LPT9",
  ];

  /**
   * Validates a remote path for OneDrive/SharePoint
   */
  static validateRemotePath(remotePath: string): {
    valid: boolean;
    sanitized: string;
    errors: string[];
  } {
    const errors: string[] = [];
    let sanitized = remotePath;

    // Remove leading/trailing slashes
    sanitized = sanitized.replace(/^\/+|\/+$/g, "");

    // Check for invalid characters
    if (this.INVALID_CHARS.test(sanitized)) {
      errors.push("Path contains invalid characters");
      sanitized = sanitized.replace(this.INVALID_CHARS, "_");
    }

    // Check path length
    if (sanitized.length > this.MAX_PATH_LENGTH) {
      errors.push(
        `Path exceeds maximum length of ${this.MAX_PATH_LENGTH} characters`,
      );
      sanitized = sanitized.substring(0, this.MAX_PATH_LENGTH);
    }

    // Check for reserved names
    const segments = sanitized.split("/");
    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i];
      const upperSegment = segment.toUpperCase();

      if (this.RESERVED_NAMES.includes(upperSegment)) {
        errors.push(`Path contains reserved name: ${segment}`);
        segments[i] = `_${segment}`;
      }

      // Check for empty segments
      if (!segment && i < segments.length - 1) {
        errors.push("Path contains empty segments");
      }
    }

    sanitized = segments.filter((s) => s).join("/");

    return {
      valid: errors.length === 0,
      sanitized,
      errors,
    };
  }

  /**
   * Validates a local file path
   */
  static validateLocalPath(localPath: string): {
    valid: boolean;
    absolute: string;
    errors: string[];
  } {
    const errors: string[] = [];

    // Convert to absolute path
    const absolute = path.resolve(localPath);

    // Check if path exists
    if (!fs.existsSync(absolute)) {
      errors.push("Path does not exist");
    }

    // Check for path traversal attempts
    const normalized = path.normalize(absolute);
    if (normalized !== absolute) {
      errors.push("Path contains traversal attempts");
    }

    // Check if file is readable (for files)
    if (fs.existsSync(absolute) && fs.statSync(absolute).isFile()) {
      try {
        fs.accessSync(absolute, fs.constants.R_OK);
      } catch {
        errors.push("File is not readable");
      }
    }

    return {
      valid: errors.length === 0,
      absolute,
      errors,
    };
  }
}

/**
 * Validates SharePoint-specific inputs
 */
export class SharePointValidator {
  /**
   * Validates a SharePoint site ID
   */
  static validateSiteId(siteId: string): boolean {
    // SharePoint site IDs are typically in format: {hostname},{guid},{guid}
    const siteIdPattern = /^[a-zA-Z0-9\-\.]+,[a-f0-9\-]+,[a-f0-9\-]+$/;
    return siteIdPattern.test(siteId);
  }

  /**
   * Validates a SharePoint list ID
   */
  static validateListId(listId: string): boolean {
    // List IDs are GUIDs
    const guidPattern =
      /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
    return guidPattern.test(listId);
  }

  /**
   * Validates OData filter expressions
   */
  static validateODataFilter(filter: string): {
    valid: boolean;
    errors: string[];
  } {
    const errors: string[] = [];

    // Check for SQL injection patterns
    const dangerousPatterns = [
      /;\s*(drop|delete|truncate|alter|create|insert|update)\s+/i,
      /union\s+select/i,
      /exec\s*\(/i,
      /xp_cmdshell/i,
    ];

    for (const pattern of dangerousPatterns) {
      if (pattern.test(filter)) {
        errors.push("Filter contains potentially dangerous patterns");
        break;
      }
    }

    // Check for balanced parentheses
    let parenCount = 0;
    for (const char of filter) {
      if (char === "(") parenCount++;
      if (char === ")") parenCount--;
      if (parenCount < 0) {
        errors.push("Unbalanced parentheses in filter");
        break;
      }
    }
    if (parenCount !== 0) {
      errors.push("Unbalanced parentheses in filter");
    }

    // Check for valid OData operators
    const validOperators = [
      "eq",
      "ne",
      "gt",
      "ge",
      "lt",
      "le",
      "and",
      "or",
      "not",
      "contains",
      "startswith",
      "endswith",
    ];
    const operatorPattern =
      /\b(eq|ne|gt|ge|lt|le|and|or|not|contains|startswith|endswith)\b/gi;
    const usedOperators = filter.match(operatorPattern) || [];

    for (const op of usedOperators) {
      if (!validOperators.includes(op.toLowerCase())) {
        errors.push(`Invalid OData operator: ${op}`);
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }
}

/**
 * Validates Excel-specific inputs
 */
export class ExcelValidator {
  /**
   * Validates an Excel range address
   */
  static validateRange(range: string): {
    valid: boolean;
    normalized: string;
    errors: string[];
  } {
    const errors: string[] = [];
    const normalized = range.toUpperCase().trim();

    // Basic range pattern: A1 or A1:B10
    const rangePattern = /^[A-Z]+\d+(:[A-Z]+\d+)?$/;

    if (!rangePattern.test(normalized)) {
      errors.push("Invalid Excel range format");
    }

    // Check for reasonable bounds
    const parts = normalized.split(":");
    for (const part of parts) {
      const colMatch = part.match(/^([A-Z]+)/);
      const rowMatch = part.match(/(\d+)$/);

      if (colMatch && rowMatch) {
        const col = colMatch[1];
        const row = parseInt(rowMatch[1]);

        // Excel max column is XFD (16384)
        if (col.length > 3 || this.columnToNumber(col) > 16384) {
          errors.push(`Column ${col} exceeds Excel limits`);
        }

        // Excel max row is 1048576
        if (row > 1048576 || row < 1) {
          errors.push(`Row ${row} exceeds Excel limits`);
        }
      }
    }

    return {
      valid: errors.length === 0,
      normalized,
      errors,
    };
  }

  /**
   * Converts column letters to number
   */
  private static columnToNumber(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
  }

  /**
   * Validates worksheet name
   */
  static validateWorksheetName(name: string): {
    valid: boolean;
    sanitized: string;
    errors: string[];
  } {
    const errors: string[] = [];
    let sanitized = name;

    // Excel worksheet name limitations
    const invalidChars = /[\\\/\*\?\[\]:]/g;
    const maxLength = 31;

    if (invalidChars.test(sanitized)) {
      errors.push("Worksheet name contains invalid characters");
      sanitized = sanitized.replace(invalidChars, "_");
    }

    if (sanitized.length > maxLength) {
      errors.push(`Worksheet name exceeds ${maxLength} characters`);
      sanitized = sanitized.substring(0, maxLength);
    }

    if (!sanitized) {
      errors.push("Worksheet name cannot be empty");
      sanitized = "Sheet1";
    }

    return {
      valid: errors.length === 0,
      sanitized,
      errors,
    };
  }

  /**
   * Validates a 2D array of values for Excel
   */
  static validateValues(values: any[][]): { valid: boolean; errors: string[] } {
    const errors: string[] = [];

    if (!Array.isArray(values)) {
      errors.push("Values must be an array");
      return { valid: false, errors };
    }

    if (values.length === 0) {
      errors.push("Values array cannot be empty");
      return { valid: false, errors };
    }

    // Check if all rows have the same length
    const firstRowLength = values[0]?.length || 0;
    for (let i = 0; i < values.length; i++) {
      if (!Array.isArray(values[i])) {
        errors.push(`Row ${i} is not an array`);
      } else if (values[i].length !== firstRowLength) {
        errors.push(`Row ${i} has inconsistent column count`);
      }
    }

    // Check for Excel value limits
    for (let i = 0; i < values.length; i++) {
      for (let j = 0; j < values[i]?.length; j++) {
        const value = values[i][j];

        // Check string length (Excel limit is 32767 characters)
        if (typeof value === "string" && value.length > 32767) {
          errors.push(`Cell [${i},${j}] exceeds Excel string limit`);
        }

        // Check number range
        if (typeof value === "number") {
          if (!isFinite(value)) {
            errors.push(`Cell [${i},${j}] contains invalid number`);
          }
        }
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }
}

/**
 * Validates email addresses and recipients
 */
export class EmailValidator {
  /**
   * Validates an email address
   */
  static validateEmail(email: string): boolean {
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailPattern.test(email);
  }

  /**
   * Validates multiple email addresses
   */
  static validateEmails(emails: string[]): {
    valid: boolean;
    invalid: string[];
  } {
    const invalid: string[] = [];

    for (const email of emails) {
      if (!this.validateEmail(email)) {
        invalid.push(email);
      }
    }

    return {
      valid: invalid.length === 0,
      invalid,
    };
  }
}

/**
 * Validates file operations
 */
export class FileOperationValidator {
  /**
   * Validates file size for upload
   */
  static validateFileSize(
    filePath: string,
    maxSizeMB: number = 250,
  ): { valid: boolean; size: number; errors: string[] } {
    const errors: string[] = [];
    let size = 0;

    try {
      const stats = fs.statSync(filePath);
      size = stats.size;

      const sizeMB = size / (1024 * 1024);
      if (sizeMB > maxSizeMB) {
        errors.push(
          `File size ${sizeMB.toFixed(2)}MB exceeds maximum of ${maxSizeMB}MB`,
        );
      }
    } catch (error) {
      errors.push("Cannot read file size");
    }

    return {
      valid: errors.length === 0,
      size,
      errors,
    };
  }

  /**
   * Validates MIME type
   */
  static validateMimeType(
    filePath: string,
    allowedTypes?: string[],
  ): { valid: boolean; mimeType: string; errors: string[] } {
    const errors: string[] = [];
    const ext = path.extname(filePath).toLowerCase();

    // Basic MIME type mapping
    const mimeTypes: Record<string, string> = {
      ".txt": "text/plain",
      ".pdf": "application/pdf",
      ".doc": "application/msword",
      ".docx":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      ".xls": "application/vnd.ms-excel",
      ".xlsx":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      ".ppt": "application/vnd.ms-powerpoint",
      ".pptx":
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".png": "image/png",
      ".gif": "image/gif",
      ".zip": "application/zip",
      ".json": "application/json",
      ".xml": "application/xml",
    };

    const mimeType = mimeTypes[ext] || "application/octet-stream";

    if (allowedTypes && allowedTypes.length > 0) {
      if (!allowedTypes.includes(mimeType)) {
        errors.push(`File type ${mimeType} is not allowed`);
      }
    }

    return {
      valid: errors.length === 0,
      mimeType,
      errors,
    };
  }
}

/**
 * General input sanitizer
 */
export class InputSanitizer {
  /**
   * Sanitizes a string for safe display
   */
  static sanitizeString(input: string, maxLength: number = 1000): string {
    // Remove control characters except tabs and newlines
    let sanitized = input.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, "");

    // Truncate if too long
    if (sanitized.length > maxLength) {
      sanitized = sanitized.substring(0, maxLength) + "...";
    }

    return sanitized;
  }

  /**
   * Sanitizes an object for safe JSON serialization
   */
  static sanitizeObject(
    obj: any,
    maxDepth: number = 10,
    currentDepth: number = 0,
  ): any {
    if (currentDepth >= maxDepth) {
      return "[Max depth reached]";
    }

    if (obj === null || obj === undefined) {
      return obj;
    }

    if (typeof obj !== "object") {
      if (typeof obj === "string") {
        return this.sanitizeString(obj);
      }
      return obj;
    }

    if (Array.isArray(obj)) {
      return obj.map((item) =>
        this.sanitizeObject(item, maxDepth, currentDepth + 1),
      );
    }

    const sanitized: any = {};
    for (const [key, value] of Object.entries(obj)) {
      // Skip circular references
      if (value === obj) {
        sanitized[key] = "[Circular reference]";
      } else {
        sanitized[key] = this.sanitizeObject(value, maxDepth, currentDepth + 1);
      }
    }

    return sanitized;
  }
}
