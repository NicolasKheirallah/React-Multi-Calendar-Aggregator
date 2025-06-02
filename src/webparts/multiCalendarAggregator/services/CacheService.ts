import { AppConstants } from '../constants/AppConstants';
import { ICalendarEvent, ICalendarSource } from '../models/ICalendarModels';

export interface ICacheItem<T> {
  data: T;
  timestamp: number;
  expiry: number;
}

export class CacheService {
  private static instance: CacheService;
  private cache: Map<string, ICacheItem<unknown>> = new Map();
  private defaultTtl: number = AppConstants.CACHE_DURATION_MINUTES * 60 * 1000; // milliseconds

  /**
   * Singleton pattern - get instance
   */
  public static getInstance(): CacheService {
    if (!CacheService.instance) {
      CacheService.instance = new CacheService();
    }
    return CacheService.instance;
  }

  /**
   * Generate cache key
   */
  private generateKey(prefix: string, identifier: string): string {
    return `${AppConstants.CACHE_KEY_PREFIX}${prefix}-${identifier}`;
  }

  /**
   * Set item in cache with TTL
   */
  public set<T>(key: string, data: T, ttlMinutes?: number): void {
    const ttl = ttlMinutes ? ttlMinutes * 60 * 1000 : this.defaultTtl;
    const now = Date.now();
    
    const cacheItem: ICacheItem<T> = {
      data,
      timestamp: now,
      expiry: now + ttl
    };

    this.cache.set(key, cacheItem);
    
    // Clean up expired items periodically
    this.cleanupExpired();
  }

  /**
   * Get item from cache
   */
  public get<T>(key: string): T | undefined {
    const item = this.cache.get(key) as ICacheItem<T>;
    
    if (!item) {
      return undefined;
    }

    // Check if expired
    if (Date.now() > item.expiry) {
      this.cache.delete(key);
      return undefined;
    }

    return item.data;
  }

  /**
   * Check if item exists in cache and is not expired
   */
  public has(key: string): boolean {
    const item = this.cache.get(key);
    
    if (!item) {
      return false;
    }

    // Check if expired
    if (Date.now() > item.expiry) {
      this.cache.delete(key);
      return false;
    }

    return true;
  }

  /**
   * Remove item from cache
   */
  public delete(key: string): boolean {
    return this.cache.delete(key);
  }

  /**
   * Clear all cache
   */
  public clear(): void {
    this.cache.clear();
  }

  /**
   * Get cache statistics
   */
  public getStats(): { size: number; keys: string[] } {
    return {
      size: this.cache.size,
      keys: Array.from(this.cache.keys())
    };
  }

  /**
   * Clean up expired items
   */
  private cleanupExpired(): void {
    const now = Date.now();
    const expiredKeys: string[] = [];

    this.cache.forEach((item, key) => {
      if (now > item.expiry) {
        expiredKeys.push(key);
      }
    });

    expiredKeys.forEach(key => this.cache.delete(key));
  }

  // Calendar Sources Cache Methods

  /**
   * Cache calendar sources
   */
  public setCachedSources(sources: ICalendarSource[], ttlMinutes?: number): void {
    const key = this.generateKey(AppConstants.CACHE_SOURCES_KEY, 'all');
    this.set(key, sources, ttlMinutes);
  }

  /**
   * Get cached calendar sources
   */
  public getCachedSources(): ICalendarSource[] | undefined {
    const key = this.generateKey(AppConstants.CACHE_SOURCES_KEY, 'all');
    return this.get<ICalendarSource[]>(key);
  }

  /**
   * Cache SharePoint sources
   */
  public setCachedSharePointSources(sources: ICalendarSource[], ttlMinutes?: number): void {
    const key = this.generateKey(AppConstants.CACHE_SOURCES_KEY, 'sharepoint');
    this.set(key, sources, ttlMinutes);
  }

  /**
   * Get cached SharePoint sources
   */
  public getCachedSharePointSources(): ICalendarSource[] | undefined {
    const key = this.generateKey(AppConstants.CACHE_SOURCES_KEY, 'sharepoint');
    return this.get<ICalendarSource[]>(key);
  }

  /**
   * Cache Exchange sources
   */
  public setCachedExchangeSources(sources: ICalendarSource[], ttlMinutes?: number): void {
    const key = this.generateKey(AppConstants.CACHE_SOURCES_KEY, 'exchange');
    this.set(key, sources, ttlMinutes);
  }

  /**
   * Get cached Exchange sources
   */
  public getCachedExchangeSources(): ICalendarSource[] | undefined {
    const key = this.generateKey(AppConstants.CACHE_SOURCES_KEY, 'exchange');
    return this.get<ICalendarSource[]>(key);
  }

  // Events Cache Methods

  /**
   * Cache events for a specific calendar
   */
  public setCachedEvents(calendarId: string, events: ICalendarEvent[], ttlMinutes?: number): void {
    const key = this.generateKey(AppConstants.CACHE_EVENTS_KEY, calendarId);
    this.set(key, events, ttlMinutes);
  }

  /**
   * Get cached events for a specific calendar
   */
  public getCachedEvents(calendarId: string): ICalendarEvent[] | undefined {
    const key = this.generateKey(AppConstants.CACHE_EVENTS_KEY, calendarId);
    return this.get<ICalendarEvent[]>(key);
  }

  /**
   * Cache aggregated events from multiple calendars
   */
  public setCachedAggregatedEvents(calendarIds: string[], events: ICalendarEvent[], ttlMinutes?: number): void {
    const sortedIds = calendarIds.sort().join(',');
    const key = this.generateKey(AppConstants.CACHE_EVENTS_KEY, `aggregated-${sortedIds}`);
    this.set(key, events, ttlMinutes);
  }

  /**
   * Get cached aggregated events
   */
  public getCachedAggregatedEvents(calendarIds: string[]): ICalendarEvent[] | undefined {
    const sortedIds = calendarIds.sort().join(',');
    const key = this.generateKey(AppConstants.CACHE_EVENTS_KEY, `aggregated-${sortedIds}`);
    return this.get<ICalendarEvent[]>(key);
  }

  /**
   * Cache search results
   */
  public setCachedSearchResults(query: string, results: ICalendarEvent[], ttlMinutes?: number): void {
    const normalizedQuery = query.toLowerCase().trim();
    const key = this.generateKey('search', normalizedQuery);
    this.set(key, results, ttlMinutes || 5); // Search results expire faster
  }

  /**
   * Get cached search results
   */
  public getCachedSearchResults(query: string): ICalendarEvent[] | undefined {
    const normalizedQuery = query.toLowerCase().trim();
    const key = this.generateKey('search', normalizedQuery);
    return this.get<ICalendarEvent[]>(key);
  }

  // Date Range Cache Methods

  /**
   * Cache events for date range
   */
  public setCachedDateRangeEvents(
    startDate: Date, 
    endDate: Date, 
    calendarIds: string[], 
    events: ICalendarEvent[], 
    ttlMinutes?: number
  ): void {
    const dateKey = `${startDate.toISOString().split('T')[0]}-${endDate.toISOString().split('T')[0]}`;
    const sortedIds = calendarIds.sort().join(',');
    const key = this.generateKey('daterange', `${dateKey}-${sortedIds}`);
    this.set(key, events, ttlMinutes);
  }

  /**
   * Get cached events for date range
   */
  public getCachedDateRangeEvents(startDate: Date, endDate: Date, calendarIds: string[]): ICalendarEvent[] | undefined {
    const dateKey = `${startDate.toISOString().split('T')[0]}-${endDate.toISOString().split('T')[0]}`;
    const sortedIds = calendarIds.sort().join(',');
    const key = this.generateKey('daterange', `${dateKey}-${sortedIds}`);
    return this.get<ICalendarEvent[]>(key);
  }

  // User Preferences Cache

  /**
   * Cache user preferences
   */
  public setCachedUserPreferences(userId: string, preferences: Record<string, unknown>, ttlMinutes?: number): void {
    const key = this.generateKey('userprefs', userId);
    this.set(key, preferences, ttlMinutes || 60); // User preferences cache for 1 hour
  }

  /**
   * Get cached user preferences
   */
  public getCachedUserPreferences(userId: string): Record<string, unknown> | undefined {
    const key = this.generateKey('userprefs', userId);
    return this.get(key);
  }

  // Calendar Metadata Cache

  /**
   * Cache calendar metadata (permissions, stats, etc.)
   */
  public setCachedCalendarMetadata(calendarId: string, metadata: Record<string, unknown>, ttlMinutes?: number): void {
    const key = this.generateKey('metadata', calendarId);
    this.set(key, metadata, ttlMinutes || 30); // Metadata cache for 30 minutes
  }

  /**
   * Get cached calendar metadata
   */
  public getCachedCalendarMetadata(calendarId: string): Record<string, unknown> | undefined {
    const key = this.generateKey('metadata', calendarId);
    return this.get(key);
  }

  // Batch Operations

  /**
   * Clear all calendar-related cache
   */
  public clearCalendarCache(): void {
    const keysToDelete: string[] = [];
    
    this.cache.forEach((_, key) => {
      if (key.includes(AppConstants.CACHE_SOURCES_KEY) || 
          key.includes(AppConstants.CACHE_EVENTS_KEY) ||
          key.includes('metadata') ||
          key.includes('daterange')) {
        keysToDelete.push(key);
      }
    });

    keysToDelete.forEach(key => this.cache.delete(key));
  }

  /**
   * Clear cache for specific calendar
   */
  public clearCalendarSpecificCache(calendarId: string): void {
    const keysToDelete: string[] = [];
    
    this.cache.forEach((_, key) => {
      if (key.includes(calendarId)) {
        keysToDelete.push(key);
      }
    });

    keysToDelete.forEach(key => this.cache.delete(key));
  }

  /**
   * Invalidate cache based on patterns
   */
  public invalidateByPattern(pattern: string): void {
    const keysToDelete: string[] = [];
    
    this.cache.forEach((_, key) => {
      if (key.includes(pattern)) {
        keysToDelete.push(key);
      }
    });

    keysToDelete.forEach(key => this.cache.delete(key));
  }

  // Cache Health Methods

  /**
   * Get cache health information
   */
  public getCacheHealth(): {
    totalItems: number;
    expiredItems: number;
    memoryUsage: number;
    hitRate: number;
  } {
    const now = Date.now();
    let expiredCount = 0;
    let totalSize = 0;

    this.cache.forEach((item) => {
      if (now > item.expiry) {
        expiredCount++;
      }
      totalSize += JSON.stringify(item.data).length;
    });

    return {
      totalItems: this.cache.size,
      expiredItems: expiredCount,
      memoryUsage: totalSize,
      hitRate: 0 // Would need to track hits/misses for accurate calculation
    };
  }

  /**
   * Optimize cache by removing expired items and compacting
   */
  public optimizeCache(): void {
    this.cleanupExpired();
    
    // If cache is still too large, remove oldest items
    if (this.cache.size > 100) { // Max 100 items
      const entries = Array.from(this.cache.entries());
      entries.sort((a, b) => a[1].timestamp - b[1].timestamp);
      
      // Remove oldest 25% of items
      const toRemove = Math.floor(entries.length * 0.25);
      for (let i = 0; i < toRemove; i++) {
        this.cache.delete(entries[i][0]);
      }
    }
  }

  /**
   * Set cache configuration
   */
  public configure(options: { defaultTtlMinutes?: number }): void {
    if (options.defaultTtlMinutes) {
      this.defaultTtl = options.defaultTtlMinutes * 60 * 1000;
    }
  }

  /**
   * Export cache for debugging
   */
  public exportCache(): Record<string, {
    data: unknown;
    timestamp: string;
    expiry: string;
    isExpired: boolean;
  }> {
    const exported: Record<string, {
      data: unknown;
      timestamp: string;
      expiry: string;
      isExpired: boolean;
    }> = {};
    
    this.cache.forEach((item, key) => {
      exported[key] = {
        data: item.data,
        timestamp: new Date(item.timestamp).toISOString(),
        expiry: new Date(item.expiry).toISOString(),
        isExpired: Date.now() > item.expiry
      };
    });

    return exported;
  }
}