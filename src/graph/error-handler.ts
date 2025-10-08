/**
 * Enhanced error handling for Microsoft Graph API
 * Business-aware error categorization and user-friendly messaging
 */

import { GraphError, ApiError } from './models';

export class GraphApiError extends Error {
  public readonly code: string;
  public readonly statusCode?: number;
  public readonly context?: string;
  public readonly originalError?: any;
  public readonly category: ErrorCategory;
  public readonly severity: ErrorSeverity;
  public readonly isRetryable: boolean;
  public readonly suggestedAction?: string;

  constructor(
    error: GraphError | Error | any,
    context?: string,
    statusCode?: number
  ) {
    const message = GraphApiError.extractMessage(error);
    super(message);
    
    this.name = 'GraphApiError';
    this.code = GraphApiError.extractCode(error);
    this.statusCode = statusCode;
    this.context = context;
    this.originalError = error;
    
    const errorInfo = GraphApiError.categorizeError(this.code, statusCode);
    this.category = errorInfo.category;
    this.severity = errorInfo.severity;
    this.isRetryable = errorInfo.isRetryable;
    this.suggestedAction = errorInfo.suggestedAction;
  }

  static extractMessage(error: any): string {
    if (typeof error === 'string') return error;
    if (error?.error?.message) return error.error.message;
    if (error?.message) return error.message;
    if (error?.error_description) return error.error_description;
    return 'Unknown Microsoft Graph API error';
  }

  static extractCode(error: any): string {
    if (error?.error?.code) return error.error.code;
    if (error?.code) return error.code;
    if (error?.error) return error.error;
    return 'UnknownError';
  }

  static categorizeError(code: string, statusCode?: number): ErrorInfo {
    // Authentication and authorization errors
    if (AUTH_ERRORS.includes(code) || statusCode === 401) {
      return {
        category: 'Authentication',
        severity: 'High',
        isRetryable: false,
        suggestedAction: 'Please re-authenticate using the setup-auth script'
      };
    }

    // Permission and access errors
    if (PERMISSION_ERRORS.includes(code) || statusCode === 403) {
      return {
        category: 'Permission',
        severity: 'High',
        isRetryable: false,
        suggestedAction: 'Check that your app has the required permissions and admin consent'
      };
    }

    // Resource not found errors
    if (NOT_FOUND_ERRORS.includes(code) || statusCode === 404) {
      return {
        category: 'NotFound',
        severity: 'Medium',
        isRetryable: false,
        suggestedAction: 'Verify the resource ID/path is correct and the resource exists'
      };
    }

    // Rate limiting and throttling
    if (THROTTLING_ERRORS.includes(code) || statusCode === 429) {
      return {
        category: 'Throttling',
        severity: 'Medium',
        isRetryable: true,
        suggestedAction: 'Reduce request frequency and implement exponential backoff'
      };
    }

    // Quota and capacity errors
    if (QUOTA_ERRORS.includes(code) || statusCode === 507) {
      return {
        category: 'Quota',
        severity: 'High',
        isRetryable: false,
        suggestedAction: 'Free up storage space or upgrade your Office 365 plan'
      };
    }

    // Validation and input errors
    if (VALIDATION_ERRORS.includes(code) || statusCode === 400) {
      return {
        category: 'Validation',
        severity: 'Medium',
        isRetryable: false,
        suggestedAction: 'Check input parameters and request format'
      };
    }

    // Conflict errors (optimistic concurrency, etc.)
    if (CONFLICT_ERRORS.includes(code) || statusCode === 409) {
      return {
        category: 'Conflict',
        severity: 'Medium',
        isRetryable: true,
        suggestedAction: 'Resource was modified by another process. Refresh and retry'
      };
    }

    // Server and service errors
    if (SERVER_ERRORS.includes(code) || (statusCode && statusCode >= 500)) {
      return {
        category: 'Server',
        severity: 'High',
        isRetryable: true,
        suggestedAction: 'Microsoft Graph service may be experiencing issues. Try again later'
      };
    }

    // Network and connectivity errors
    if (NETWORK_ERRORS.includes(code)) {
      return {
        category: 'Network',
        severity: 'Medium',
        isRetryable: true,
        suggestedAction: 'Check internet connectivity and try again'
      };
    }

    // Default case
    return {
      category: 'Unknown',
      severity: 'Medium',
      isRetryable: false,
      suggestedAction: 'Review the error details and contact support if the issue persists'
    };
  }

  toApiError(): ApiError {
    return {
      code: this.code,
      message: this.message,
      details: this.suggestedAction,
      statusCode: this.statusCode,
      context: this.context
    };
  }

  toString(): string {
    let errorStr = `${this.category} Error [${this.code}]: ${this.message}`;
    
    if (this.context) {
      errorStr += `\nContext: ${this.context}`;
    }
    
    if (this.statusCode) {
      errorStr += `\nHTTP Status: ${this.statusCode}`;
    }
    
    if (this.suggestedAction) {
      errorStr += `\nSuggested Action: ${this.suggestedAction}`;
    }
    
    return errorStr;
  }
}

// Error categorization types
type ErrorCategory = 
  | 'Authentication'
  | 'Permission' 
  | 'NotFound'
  | 'Throttling'
  | 'Quota'
  | 'Validation'
  | 'Conflict'
  | 'Server'
  | 'Network'
  | 'Unknown';

type ErrorSeverity = 'Low' | 'Medium' | 'High' | 'Critical';

interface ErrorInfo {
  category: ErrorCategory;
  severity: ErrorSeverity;
  isRetryable: boolean;
  suggestedAction: string;
}

// Error code classifications
const AUTH_ERRORS = [
  'InvalidAuthenticationToken',
  'AuthenticationFailure',
  'ExpiredAuthenticationToken',
  'MalformedAuthenticationToken',
  'InvalidTokenType',
  'TokenNotFound',
  'Unauthenticated'
];

const PERMISSION_ERRORS = [
  'Forbidden',
  'InsufficientScope',
  'AccessDenied',
  'Unauthorized',
  'InsufficientPrivileges',
  'InvalidPermission',
  'ConsentRequired',
  'UserNotInTenant'
];

const NOT_FOUND_ERRORS = [
  'NotFound',
  'ResourceNotFound',
  'ItemNotFound',
  'FileNotFound',
  'DriveItemNotFound',
  'SiteNotFound',
  'ListNotFound',
  'WorkbookNotFound',
  'WorksheetNotFound'
];

const THROTTLING_ERRORS = [
  'TooManyRequests',
  'RateLimitExceeded',
  'RequestThrottled',
  'ThrottledRequest',
  'ServiceUnavailable',
  'TemporarilyUnavailable'
];

const QUOTA_ERRORS = [
  'InsufficientStorage',
  'QuotaExceeded',
  'StorageQuotaExceeded',
  'DiskQuotaExceeded',
  'FileSizeLimitExceeded',
  'ItemCountLimitExceeded'
];

const VALIDATION_ERRORS = [
  'BadRequest',
  'InvalidRequest',
  'InvalidParameter',
  'MalformedRequest',
  'InvalidRange',
  'InvalidFilter',
  'InvalidOrderBy',
  'ValidationError',
  'InvalidFileName',
  'InvalidPath',
  'UnsupportedMediaType'
];

const CONFLICT_ERRORS = [
  'Conflict',
  'ResourceModified',
  'ConcurrencyFailure',
  'NameAlreadyExists',
  'ItemAlreadyExists',
  'DuplicateName',
  'VersionConflict'
];

const SERVER_ERRORS = [
  'InternalServerError',
  'ServiceUnavailable',
  'BadGateway',
  'GatewayTimeout',
  'ServiceError',
  'TemporaryFailure',
  'UnexpectedError'
];

const NETWORK_ERRORS = [
  'NetworkError',
  'ConnectionError',
  'TimeoutError',
  'DNSError',
  'ConnectionRefused',
  'ConnectionTimeout'
];

// Retry logic helper
export class RetryHelper {
  static shouldRetry(error: GraphApiError, attempt: number, maxAttempts: number = 3): boolean {
    if (attempt >= maxAttempts) return false;
    if (!error.isRetryable) return false;
    
    // Don't retry authentication or permission errors
    if (error.category === 'Authentication' || error.category === 'Permission') {
      return false;
    }
    
    return true;
  }

  static getRetryDelay(attempt: number, baseDelay: number = 1000): number {
    // Exponential backoff with jitter
    const exponentialDelay = baseDelay * Math.pow(2, attempt - 1);
    const jitter = Math.random() * 0.1 * exponentialDelay;
    return exponentialDelay + jitter;
  }

  static async withRetry<T>(
    operation: () => Promise<T>,
    context: string,
    maxAttempts: number = 3,
    baseDelay: number = 1000
  ): Promise<T> {
    let lastError: GraphApiError;
    
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        return await operation();
      } catch (error) {
        const graphError = error instanceof GraphApiError 
          ? error 
          : new GraphApiError(error, context);
        
        lastError = graphError;
        
        if (!this.shouldRetry(graphError, attempt, maxAttempts)) {
          throw graphError;
        }
        
        if (attempt < maxAttempts) {
          const delay = this.getRetryDelay(attempt, baseDelay);
          console.log(`Retrying ${context} in ${delay}ms (attempt ${attempt}/${maxAttempts})`);
          await new Promise(resolve => setTimeout(resolve, delay));
        }
      }
    }
    
    throw lastError!;
  }
}

// Utility functions for common error scenarios
export function isAuthenticationError(error: any): boolean {
  const graphError = error instanceof GraphApiError 
    ? error 
    : new GraphApiError(error);
  return graphError.category === 'Authentication';
}

export function isPermissionError(error: any): boolean {
  const graphError = error instanceof GraphApiError 
    ? error 
    : new GraphApiError(error);
  return graphError.category === 'Permission';
}

export function isNotFoundError(error: any): boolean {
  const graphError = error instanceof GraphApiError 
    ? error 
    : new GraphApiError(error);
  return graphError.category === 'NotFound';
}

export function isThrottlingError(error: any): boolean {
  const graphError = error instanceof GraphApiError 
    ? error 
    : new GraphApiError(error);
  return graphError.category === 'Throttling';
}

export function createUserFriendlyError(error: any, context?: string): string {
  const graphError = error instanceof GraphApiError 
    ? error 
    : new GraphApiError(error, context);
  
  let message = `${graphError.category} Error: ${graphError.message}`;
  
  if (graphError.suggestedAction) {
    message += `\n\nSuggested Action: ${graphError.suggestedAction}`;
  }
  
  return message;
}