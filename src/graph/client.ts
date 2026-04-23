/**
 * Enhanced Microsoft Graph API client
 * Optimized for OneDrive, SharePoint, and Excel operations
 */
import axios, { AxiosInstance, AxiosResponse, AxiosError } from "axios";
import { basename } from "node:path";
import { getAuthInstance } from "../auth/microsoft-graph-auth.js";
import { GraphApiError, RetryHelper } from "./error-handler.js";
import { GRAPH_BASE_URL, buildUrl } from "../config/endpoints.js";
import { GraphResponse, WorkbookSession, McpResponse } from "./models.js";
import { assertGraphPayloadHasNoError } from "./contracts.js";
import {
  metadataCache,
  searchCache,
  driveCache,
} from "../utils/cache-manager.js";
import { SecurityValidator, AuditLogger } from "../utils/security-validator.js";
import * as FormData from "form-data";
import { createReadStream } from "fs";
import { lookup } from "mime-types";

export interface RequestOptions {
  timeout?: number;
  retries?: number;
  headers?: Record<string, string>;
  validateStatus?: (status: number) => boolean;
}

export interface UploadOptions extends RequestOptions {
  conflictBehavior?: "fail" | "replace" | "rename";
  onProgress?: (loaded: number, total: number) => void;
}

export interface PaginationOptions extends RequestOptions {
  /**
   * Hard cap on items returned across all pages. Defaults to 10_000 so a
   * misconfigured caller cannot OOM the process on very large drives.
   * Pass `Infinity` explicitly to opt out (not recommended).
   */
  maxItems?: number;
  /** Hard cap on page fetches. Defaults to 1_000. */
  maxPages?: number;
}

export const DEFAULT_MAX_PAGINATION_ITEMS = 10_000;
export const DEFAULT_MAX_PAGINATION_PAGES = 1_000;

export class GraphClient {
  private axios: AxiosInstance;
  private sessionCache: Map<string, WorkbookSession> = new Map();
  private rateLimitDelay = 0;
  private lastRequestTime = 0;

  constructor() {
    this.axios = axios.create({
      baseURL: GRAPH_BASE_URL,
      timeout: 30000,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
    });

    this.setupInterceptors();
  }

  private setupInterceptors(): void {
    // Request interceptor for authentication and rate limiting
    this.axios.interceptors.request.use(
      async (config) => {
        // Add authentication header
        const { getAuthInstance } = await import(
          "../auth/microsoft-graph-auth.js"
        );
        const auth = getAuthInstance();
        const token = await auth.getAccessToken();
        config.headers.Authorization = `Bearer ${token}`;

        // Rate limiting
        await this.handleRateLimit();

        return config;
      },
      (error) => Promise.reject(new GraphApiError(error, "Request setup")),
    );

    // Response interceptor for error handling
    this.axios.interceptors.response.use(
      (response) => {
        this.updateRateLimitInfo(response);
        return response;
      },
      (error: AxiosError) => {
        const context = `${error.config?.method?.toUpperCase()} ${error.config?.url}`;
        const statusCode = error.response?.status;
        const responseData = error.response?.data;

        return Promise.reject(
          new GraphApiError(responseData || error, context, statusCode),
        );
      },
    );
  }

  private async handleRateLimit(): Promise<void> {
    if (this.rateLimitDelay > 0) {
      const elapsed = Date.now() - this.lastRequestTime;
      if (elapsed < this.rateLimitDelay) {
        await new Promise((resolve) =>
          setTimeout(resolve, this.rateLimitDelay - elapsed),
        );
      }
    }
    this.lastRequestTime = Date.now();
  }

  private updateRateLimitInfo(response: AxiosResponse): void {
    const retryAfter = response.headers["retry-after"];
    if (retryAfter) {
      this.rateLimitDelay = parseInt(retryAfter, 10) * 1000;
    } else {
      this.rateLimitDelay = 0;
    }
  }

  // Core HTTP methods with retry logic and caching

  async get<T>(
    endpoint: string,
    params?: Record<string, any>,
    options: RequestOptions = {},
  ): Promise<McpResponse<T>> {
    // Validate OData parameters
    if (params) {
      const validation = SecurityValidator.validateODataQuery(params);
      if (!validation.isValid) {
        throw new GraphApiError(validation.error!, "Parameter validation");
      }
    }

    // Check cache for metadata requests
    const cacheKey = this.generateCacheKey(endpoint, params);
    if (this.isCacheableRequest(endpoint)) {
      const cached = metadataCache.get(cacheKey);
      if (cached) {
        return this.wrapResponse(cached, "onedrive");
      }
    }

    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, params, false);
        const response = await this.axios.get<T>(url, {
          timeout: options.timeout,
          headers: options.headers,
          validateStatus: options.validateStatus,
        });
        const responseData = assertGraphPayloadHasNoError(
          response.data,
          `GET ${endpoint}`,
        );

        // Cache successful metadata responses
        if (this.isCacheableRequest(endpoint) && responseData) {
          metadataCache.set(cacheKey, responseData);
        }

        return this.wrapResponse(responseData, "onedrive");
      },
      `GET ${endpoint}`,
      options.retries,
    );
  }

  async post<T>(
    endpoint: string,
    data?: any,
    options: RequestOptions = {},
  ): Promise<McpResponse<T>> {
    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, {}, false);
        const response = await this.axios.post<T>(url, data, {
          timeout: options.timeout,
          headers: options.headers,
          validateStatus: options.validateStatus,
        });

        return this.wrapResponse(response.data, "onedrive");
      },
      `POST ${endpoint}`,
      options.retries,
    );
  }

  async put<T>(
    endpoint: string,
    data?: any,
    options: RequestOptions = {},
  ): Promise<McpResponse<T>> {
    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, {}, false);
        const response = await this.axios.put<T>(url, data, {
          timeout: options.timeout,
          headers: options.headers,
          validateStatus: options.validateStatus,
        });

        return this.wrapResponse(response.data, "onedrive");
      },
      `PUT ${endpoint}`,
      options.retries,
    );
  }

  async patch<T>(
    endpoint: string,
    data?: any,
    options: RequestOptions = {},
  ): Promise<McpResponse<T>> {
    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, {}, false);
        const response = await this.axios.patch<T>(url, data, {
          timeout: options.timeout,
          headers: options.headers,
          validateStatus: options.validateStatus,
        });

        return this.wrapResponse(response.data, "onedrive");
      },
      `PATCH ${endpoint}`,
      options.retries,
    );
  }

  async delete<T>(
    endpoint: string,
    options: RequestOptions = {},
  ): Promise<McpResponse<T>> {
    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, {}, false);
        const response = await this.axios.delete<T>(url, {
          timeout: options.timeout,
          headers: options.headers,
          validateStatus: options.validateStatus,
        });

        return this.wrapResponse(response.data, "onedrive");
      },
      `DELETE ${endpoint}`,
      options.retries,
    );
  }

  // Specialized methods for file operations

  async downloadFile(
    endpoint: string,
    options: RequestOptions = {},
  ): Promise<McpResponse<Buffer>> {
    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, {}, false);
        const response = await this.axios.get(url, {
          responseType: "arraybuffer",
          timeout: options.timeout || 60000,
          headers: options.headers,
        });

        return this.wrapResponse(Buffer.from(response.data), "onedrive");
      },
      `DOWNLOAD ${endpoint}`,
      options.retries,
    );
  }

  async uploadFile(
    endpoint: string,
    filePath: string,
    fileName?: string,
    options: UploadOptions = {},
  ): Promise<McpResponse<any>> {
    return this.executeWithRetry(
      async () => {
        const url = buildUrl(endpoint, {}, false);
        const mimeType = lookup(filePath) || "application/octet-stream";

        // For small files (< 4MB), use simple upload
        const stats = await import("fs").then((fs) =>
          fs.promises.stat(filePath),
        );
        if (stats.size < 4 * 1024 * 1024) {
          const fileBuffer = await import("fs").then((fs) =>
            fs.promises.readFile(filePath),
          );

          const response = await this.axios.put(url, fileBuffer, {
            headers: {
              "Content-Type": mimeType,
              ...options.headers,
            },
            timeout: options.timeout || 60000,
            onUploadProgress: options.onProgress
              ? (progress) => {
                  options.onProgress!(
                    progress.loaded,
                    progress.total || stats.size,
                  );
                }
              : undefined,
          });

          return this.wrapResponse(response.data, "onedrive");
        } else {
          // For large files, use resumable upload
          return this.uploadLargeFile(endpoint, filePath, fileName, options);
        }
      },
      `UPLOAD ${endpoint}`,
      options.retries,
    );
  }

  private async uploadLargeFile(
    endpoint: string,
    filePath: string,
    fileName?: string,
    options: UploadOptions = {},
  ): Promise<McpResponse<any>> {
    const stats = await import("fs").then((fs) => fs.promises.stat(filePath));
    const fileSize = stats.size;

    // Create upload session
    const sessionUrl = endpoint + "/createUploadSession";
    const sessionData = {
      item: {
        "@microsoft.graph.conflictBehavior":
          options.conflictBehavior || "rename",
        name: fileName || basename(filePath),
      },
    };

    const sessionResponse = await this.post<any>(sessionUrl, sessionData);
    if (!sessionResponse.success || !sessionResponse.data?.uploadUrl) {
      throw new GraphApiError("Failed to create upload session");
    }

    const uploadUrl = sessionResponse.data.uploadUrl;
    const chunkSize = 320 * 1024; // 320KB chunks

    const fs = await import("fs");
    const fileStream = fs.createReadStream(filePath);

    // Safety net: if the outer promise rejects for any reason, make sure the
    // read stream is destroyed so we do not leak a file descriptor.
    const cleanup = () => {
      if (!fileStream.destroyed) {
        fileStream.destroy();
      }
    };

    try {
      return await this.uploadChunksFromStream(
        fileStream,
        uploadUrl,
        fileSize,
        chunkSize,
        options,
      );
    } finally {
      cleanup();
    }
  }

  private uploadChunksFromStream(
    fileStream: NodeJS.ReadableStream,
    uploadUrl: string,
    fileSize: number,
    chunkSize: number,
    options: UploadOptions,
  ): Promise<McpResponse<any>> {
    let uploadedBytes = 0;

    return new Promise((resolve, reject) => {
      const chunks: Buffer[] = [];
      let currentChunk = Buffer.alloc(0);

      fileStream.on("data", (chunk: string | Buffer) => {
        const bufferChunk = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
        currentChunk = Buffer.concat([currentChunk, bufferChunk]);

        while (currentChunk.length >= chunkSize) {
          const toUpload = currentChunk.slice(0, chunkSize);
          chunks.push(toUpload);
          currentChunk = currentChunk.slice(chunkSize);
        }
      });

      fileStream.on("end", async () => {
        if (currentChunk.length > 0) {
          chunks.push(currentChunk);
        }

        try {
          let result: any;
          for (let i = 0; i < chunks.length; i++) {
            const chunk = chunks[i];
            const start = uploadedBytes;
            const end = uploadedBytes + chunk.length - 1;

            // Route the chunk PUT through executeWithRetry so 429 / transient
            // failures honour the same retry + retry-after policy used by the
            // rest of the client (bare axios.put previously ignored throttling).
            const chunkResponse = await this.executeWithRetry(
              () =>
                axios.put(uploadUrl, chunk, {
                  headers: {
                    "Content-Range": `bytes ${start}-${end}/${fileSize}`,
                    "Content-Length": chunk.length.toString(),
                  },
                  timeout: options.timeout || 60000,
                }),
              `UPLOAD CHUNK ${start}-${end}`,
              options.retries,
            );

            uploadedBytes += chunk.length;

            if (options.onProgress) {
              options.onProgress(uploadedBytes, fileSize);
            }

            if (chunkResponse.status === 201 || chunkResponse.status === 200) {
              result = chunkResponse.data;
              break;
            }
          }

          resolve(this.wrapResponse(result, "onedrive"));
        } catch (error) {
          reject(new GraphApiError(error, "Large file upload"));
        }
      });

      fileStream.on("error", (error) => {
        reject(new GraphApiError(error, "File read error"));
      });
    });
  }

  // Excel session management

  async createExcelSession(
    itemId: string,
    persistChanges = true,
  ): Promise<string> {
    const cacheKey = `${itemId}-${persistChanges}`;

    if (this.sessionCache.has(cacheKey)) {
      return this.sessionCache.get(cacheKey)!.id;
    }

    const response = await this.post<WorkbookSession>(
      `/me/drive/items/${itemId}/workbook/createSession`,
      { persistChanges },
    );

    if (response.success && response.data) {
      this.sessionCache.set(cacheKey, response.data);
      return response.data.id;
    }

    throw new GraphApiError("Failed to create Excel session");
  }

  async closeExcelSession(itemId: string, sessionId?: string): Promise<void> {
    if (sessionId) {
      await this.post(`/me/drive/items/${itemId}/workbook/closeSession`, {
        sessionId,
      });
    }

    // Clear from cache
    for (const [key, session] of this.sessionCache.entries()) {
      if (key.startsWith(itemId) && (!sessionId || session.id === sessionId)) {
        this.sessionCache.delete(key);
      }
    }
  }

  // Batch requests for efficiency

  async batch(
    requests: Array<{
      id: string;
      method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
      url: string;
      body?: any;
      headers?: Record<string, string>;
    }>,
  ): Promise<McpResponse<any[]>> {
    const batchData = {
      requests: requests.map((req) => ({
        id: req.id,
        method: req.method,
        url: req.url.startsWith("/") ? req.url : `/${req.url}`,
        body: req.body,
        headers: req.headers,
      })),
    };

    const response = await this.post<any>("/$batch", batchData);

    if (response.success && response.data?.responses) {
      return this.wrapResponse(response.data.responses, "onedrive");
    }

    throw new GraphApiError("Batch request failed");
  }

  // Pagination helper

  async getAllPages<T>(
    endpoint: string,
    params?: Record<string, any>,
    options: PaginationOptions = {},
  ): Promise<McpResponse<T[]>> {
    const maxItems = options.maxItems ?? DEFAULT_MAX_PAGINATION_ITEMS;
    const maxPages = options.maxPages ?? DEFAULT_MAX_PAGINATION_PAGES;

    const allItems: T[] = [];
    let nextLink: string | undefined = buildUrl(endpoint, params, false);
    let pageCount = 0;

    while (nextLink) {
      if (pageCount >= maxPages) {
        throw new GraphApiError(
          `Pagination page cap reached (${maxPages}) for ${endpoint}. ` +
            `Increase maxPages explicitly if you really need more.`,
          `GET PAGINATED ${endpoint}`,
        );
      }

      const response = assertGraphPayloadHasNoError(
        await this.executeWithRetry(
          async () => {
            const axiosResponse = await this.axios.get<GraphResponse<T>>(
              nextLink!,
              {
                timeout: options.timeout,
                headers: options.headers,
              },
            );
            return axiosResponse.data;
          },
          `GET PAGINATED ${endpoint}`,
          options.retries,
        ),
        `GET PAGINATED ${endpoint}`,
      );

      pageCount++;

      if (response.value) {
        for (const item of response.value) {
          if (allItems.length >= maxItems) {
            throw new GraphApiError(
              `Pagination item cap reached (${maxItems}) for ${endpoint}. ` +
                `Increase maxItems explicitly or use a narrower filter.`,
              `GET PAGINATED ${endpoint}`,
            );
          }
          allItems.push(item);
        }
      }

      nextLink = response["@odata.nextLink"];
    }

    return this.wrapResponse(allItems, "onedrive");
  }

  // Utility methods

  private async executeWithRetry<T>(
    operation: () => Promise<T>,
    context: string,
    maxRetries = 3,
  ): Promise<T> {
    return RetryHelper.withRetry(operation, context, maxRetries);
  }

  private wrapResponse<T>(
    data: T,
    source: "onedrive" | "sharepoint" | "excel" = "onedrive",
  ): McpResponse<T> {
    const validatedData = assertGraphPayloadHasNoError(data);

    return {
      success: true,
      data: validatedData,
      metadata: {
        timestamp: new Date().toISOString(),
        source,
      },
    };
  }

  private generateCacheKey(
    endpoint: string,
    params?: Record<string, any>,
  ): string {
    const paramStr = params ? JSON.stringify(params) : "";
    return `${endpoint}:${paramStr}`;
  }

  private isCacheableRequest(endpoint: string): boolean {
    // Cache metadata requests but not content downloads
    const cacheablePatterns = [
      /\/me$/,
      /\/drives$/,
      /\/sites$/,
      /\/children$/,
      /\/metadata$/,
      /\/lists$/,
      /\/columns$/,
      /\/items\/[^\/]+$/,
    ];

    const nonCacheablePatterns = [
      /\/content$/,
      /\/thumbnails$/,
      /\/preview$/,
      /\/download$/,
      /\/createUploadSession$/,
    ];

    // Don't cache if it's a non-cacheable pattern
    if (nonCacheablePatterns.some((pattern) => pattern.test(endpoint))) {
      return false;
    }

    // Cache if it matches cacheable patterns
    return cacheablePatterns.some((pattern) => pattern.test(endpoint));
  }

  // Enhanced file operations with security validation

  async validateAndDownloadFile(
    endpoint: string,
    options: RequestOptions = {},
  ): Promise<McpResponse<Buffer>> {
    // Validate the endpoint doesn't contain path traversal
    const validation = SecurityValidator.validatePath(endpoint);
    if (!validation.isValid) {
      throw new GraphApiError(validation.error!, "Path validation");
    }

    const user = await this.getCurrentUserSafe();
    AuditLogger.log("file_download", user, endpoint, "success");

    return this.downloadFile(endpoint, options);
  }

  async validateAndUploadFile(
    endpoint: string,
    filePath: string,
    fileName?: string,
    options: UploadOptions = {},
  ): Promise<McpResponse<any>> {
    // Validate file path
    const pathValidation = SecurityValidator.validatePath(endpoint);
    if (!pathValidation.isValid) {
      throw new GraphApiError(pathValidation.error!, "Path validation");
    }

    // Validate file name
    const actualFileName = fileName || basename(filePath);
    const nameValidation = SecurityValidator.validateFileName(actualFileName);
    if (!nameValidation.isValid) {
      throw new GraphApiError(nameValidation.error!, "File name validation");
    }

    // Validate file size
    const stats = await import("fs").then((fs) => fs.promises.stat(filePath));
    const sizeValidation = SecurityValidator.validateFileSize(stats.size);
    if (!sizeValidation.isValid) {
      throw new GraphApiError(sizeValidation.error!, "File size validation");
    }

    const user = await this.getCurrentUserSafe();
    AuditLogger.log(
      "file_upload",
      user,
      `${endpoint}/${actualFileName}`,
      "success",
      {
        size: stats.size,
        fileName: actualFileName,
      },
    );

    return this.uploadFile(endpoint, filePath, actualFileName, options);
  }

  // Enhanced search with caching

  async searchWithCache<T>(
    endpoint: string,
    query: string,
    params?: Record<string, any>,
    options: RequestOptions = {},
  ): Promise<McpResponse<T[]>> {
    // Validate search query
    const validation = SecurityValidator.validateSearchQuery(query);
    if (!validation.isValid) {
      throw new GraphApiError(validation.error!, "Search query validation");
    }

    // Check cache
    const cacheKey = searchCache.generateKey(validation.sanitized!, params);
    const cached = searchCache.get(cacheKey);
    if (cached) {
      const user = await this.getCurrentUserSafe();
      AuditLogger.log("search", user, validation.sanitized!, "success", {
        cached: true,
      });
      return this.wrapResponse(cached, "onedrive");
    }

    // Execute search
    const searchParams = { ...params, q: validation.sanitized };
    const result = await this.get<{ value: T[] }>(
      endpoint,
      searchParams,
      options,
    );

    if (result.success && result.data?.value) {
      // Cache successful search results
      searchCache.set(cacheKey, result.data.value);

      const user = await this.getCurrentUserSafe();
      AuditLogger.log("search", user, validation.sanitized!, "success", {
        resultCount: result.data.value.length,
        cached: false,
      });

      return this.wrapResponse(result.data.value, "onedrive");
    }

    throw new GraphApiError("Search returned no results", "Search operation");
  }

  private async getCurrentUserSafe(): Promise<string> {
    try {
      const { getAuthInstance } = await import(
        "../auth/microsoft-graph-auth.js"
      );
      const auth = getAuthInstance();
      const user = await auth.getCurrentUser();
      return user?.username || "unknown";
    } catch {
      return "unknown";
    }
  }

  // Cache management methods

  clearCaches(): void {
    metadataCache.clear();
    searchCache.clear();
    driveCache.clear();
  }

  getCacheStats(): any {
    return {
      metadata: metadataCache.getStats(),
      search: searchCache.getStats(),
      drive: driveCache.getStats(),
    };
  }

  // Health check

  async healthCheck(): Promise<McpResponse<{ status: string; user: any }>> {
    try {
      const response = await this.get<any>("/me");

      if (response.success) {
        return this.wrapResponse({
          status: "healthy",
          user: response.data,
        });
      }

      throw new GraphApiError("Health check failed");
    } catch (error) {
      return {
        success: false,
        error: error instanceof GraphApiError ? error.message : "Unknown error",
        metadata: {
          timestamp: new Date().toISOString(),
          source: "onedrive",
        },
      };
    }
  }

  // Resource cleanup

  async cleanup(): Promise<void> {
    // Close all Excel sessions
    for (const [key, session] of this.sessionCache.entries()) {
      try {
        const itemId = key.split("-")[0];
        await this.closeExcelSession(itemId, session.id);
      } catch (error) {
        console.warn(`Failed to close Excel session ${session.id}:`, error);
      }
    }

    this.sessionCache.clear();

    // Clear all caches
    this.clearCaches();

    // Cleanup cache managers
    const { cleanupAllCaches } = await import("../utils/cache-manager.js");
    cleanupAllCaches();
  }
}

// Singleton instance
let clientInstance: GraphClient | null = null;

export function getGraphClient(): GraphClient {
  if (!clientInstance) {
    clientInstance = new GraphClient();
  }
  return clientInstance;
}

export function resetGraphClient(): void {
  if (clientInstance) {
    clientInstance.cleanup();
    clientInstance = null;
  }
}

export function __setGraphClientInstanceForTests(
  client: GraphClient | null,
): void {
  clientInstance = client;
}
