/**
 * Enhanced caching system for Microsoft Graph operations
 * Implements LRU cache with TTL support and automatic cleanup
 */

interface CacheEntry<T> {
  data: T;
  timestamp: number;
  ttl: number;
  hits: number;
}

interface CacheOptions {
  maxSize?: number;
  defaultTTL?: number;
  cleanupInterval?: number;
}

export class CacheManager<T = any> {
  private cache: Map<string, CacheEntry<T>> = new Map();
  private readonly maxSize: number;
  private readonly defaultTTL: number;
  private cleanupTimer?: NodeJS.Timeout;
  private stats = {
    hits: 0,
    misses: 0,
    evictions: 0,
  };

  constructor(options: CacheOptions = {}) {
    this.maxSize = options.maxSize || 100;
    this.defaultTTL = options.defaultTTL || 300000; // 5 minutes default

    if (options.cleanupInterval !== 0) {
      this.startCleanup(options.cleanupInterval || 60000); // 1 minute default
    }
  }

  /**
   * Get item from cache
   */
  get(key: string): T | null {
    const entry = this.cache.get(key);

    if (!entry) {
      this.stats.misses++;
      return null;
    }

    if (this.isExpired(entry)) {
      this.cache.delete(key);
      this.stats.misses++;
      return null;
    }

    entry.hits++;
    this.stats.hits++;

    // Move to end (LRU)
    this.cache.delete(key);
    this.cache.set(key, entry);

    return entry.data;
  }

  /**
   * Set item in cache
   */
  set(key: string, data: T, ttl?: number): void {
    // Evict if at capacity
    if (this.cache.size >= this.maxSize && !this.cache.has(key)) {
      this.evictLRU();
    }

    const entry: CacheEntry<T> = {
      data,
      timestamp: Date.now(),
      ttl: ttl || this.defaultTTL,
      hits: 0,
    };

    this.cache.set(key, entry);
  }

  /**
   * Check if key exists and is valid
   */
  has(key: string): boolean {
    const entry = this.cache.get(key);
    if (!entry) return false;

    if (this.isExpired(entry)) {
      this.cache.delete(key);
      return false;
    }

    return true;
  }

  /**
   * Delete item from cache
   */
  delete(key: string): boolean {
    return this.cache.delete(key);
  }

  /**
   * Clear entire cache
   */
  clear(): void {
    this.cache.clear();
    this.stats = { hits: 0, misses: 0, evictions: 0 };
  }

  /**
   * Get cache statistics
   */
  getStats() {
    const hitRate =
      this.stats.hits + this.stats.misses > 0
        ? (this.stats.hits / (this.stats.hits + this.stats.misses)) * 100
        : 0;

    return {
      ...this.stats,
      size: this.cache.size,
      maxSize: this.maxSize,
      hitRate: `${hitRate.toFixed(2)}%`,
    };
  }

  /**
   * Cleanup expired entries
   */
  cleanup(): void {
    const now = Date.now();
    const toDelete: string[] = [];

    for (const [key, entry] of this.cache.entries()) {
      if (this.isExpired(entry)) {
        toDelete.push(key);
      }
    }

    toDelete.forEach((key) => this.cache.delete(key));
  }

  /**
   * Stop cleanup timer
   */
  destroy(): void {
    if (this.cleanupTimer) {
      clearInterval(this.cleanupTimer);
      this.cleanupTimer = undefined;
    }
    this.clear();
  }

  private isExpired(entry: CacheEntry<T>): boolean {
    return Date.now() - entry.timestamp > entry.ttl;
  }

  private evictLRU(): void {
    // First entry is the least recently used
    const firstKey = this.cache.keys().next().value;
    if (firstKey) {
      this.cache.delete(firstKey);
      this.stats.evictions++;
    }
  }

  private startCleanup(interval: number): void {
    this.cleanupTimer = setInterval(() => {
      this.cleanup();
    }, interval);

    this.cleanupTimer.unref?.();
  }
}

// Specialized cache instances
export class MetadataCache extends CacheManager<any> {
  constructor() {
    super({
      maxSize: 500,
      defaultTTL: 600000, // 10 minutes
      cleanupInterval: 120000, // 2 minutes
    });
  }

  generateKey(itemId: string, type: string = "metadata"): string {
    return `${type}:${itemId}`;
  }
}

export class SearchCache extends CacheManager<any[]> {
  constructor() {
    super({
      maxSize: 50,
      defaultTTL: 180000, // 3 minutes
      cleanupInterval: 60000, // 1 minute
    });
  }

  generateKey(query: string, filters?: Record<string, any>): string {
    const filterStr = filters ? JSON.stringify(filters) : "";
    return `search:${query}:${filterStr}`;
  }
}

export class DriveCache extends CacheManager<any> {
  constructor() {
    super({
      maxSize: 100,
      defaultTTL: 900000, // 15 minutes
      cleanupInterval: 300000, // 5 minutes
    });
  }

  generateKey(driveId: string, path?: string): string {
    return path ? `drive:${driveId}:${path}` : `drive:${driveId}`;
  }
}

// Global cache instances
export const metadataCache = new MetadataCache();
export const searchCache = new SearchCache();
export const driveCache = new DriveCache();

// Cleanup function for graceful shutdown
export function cleanupAllCaches(): void {
  metadataCache.destroy();
  searchCache.destroy();
  driveCache.destroy();
}
