# -*- coding: utf-8 -*-
"""
Workbook Cache System for Excel Tools
Provides efficient caching of openpyxl workbooks to avoid repeated file loading
"""

import os
import time
import threading
from collections import OrderedDict
from openpyxl import load_workbook
from openpyxl.workbook import Workbook


class WorkbookCache:
    """
    Thread-safe LRU cache for openpyxl workbooks with file modification time checking
    """
    
    def __init__(self, max_size=10, max_age_seconds=300):
        """
        Initialize the workbook cache
        
        Args:
            max_size: Maximum number of workbooks to keep in cache
            max_age_seconds: Maximum age of cached workbooks in seconds (default 5 minutes)
        """
        self.max_size = max_size
        self.max_age_seconds = max_age_seconds
        self.cache = OrderedDict()
        self.lock = threading.RLock()
        self._stats = {
            'hits': 0,
            'misses': 0,
            'evictions': 0,
            'errors': 0
        }
    
    def get_workbook(self, file_path, read_only=True, data_only=True, force_read_only=True):
        """
        Get a workbook from cache or load it if not cached
        
        Args:
            file_path: Path to the Excel file
            read_only: Whether to open in read-only mode
            data_only: Whether to read only calculated values
            force_read_only: Force read-only mode to prevent file locking (default: True)
            
        Returns:
            openpyxl.Workbook: The loaded workbook
            
        Raises:
            FileNotFoundError: If the file doesn't exist
            Exception: If the file cannot be loaded
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        # Normalize path for consistent caching
        normalized_path = os.path.normpath(os.path.abspath(file_path))
        
        # Force read-only mode to prevent file locking issues
        if force_read_only:
            read_only = True
        
        cache_key = f"{normalized_path}|{read_only}|{data_only}"
        
        with self.lock:
            # Check if workbook is in cache and still valid
            if cache_key in self.cache:
                cached_item = self.cache[cache_key]
                
                # Check if cache entry is still valid
                if self._is_cache_valid(cached_item, normalized_path):
                    # Move to end (most recently used)
                    self.cache.move_to_end(cache_key)
                    self._stats['hits'] += 1
                    print(f"Cache HIT: {os.path.basename(file_path)}")
                    return cached_item['workbook']
                else:
                    # Remove invalid cache entry
                    del self.cache[cache_key]
                    print(f"Cache EXPIRED: {os.path.basename(file_path)}")
            
            # Cache miss - load the workbook
            self._stats['misses'] += 1
            print(f"Cache MISS: Loading {os.path.basename(file_path)}")
            
            try:
                # Load workbook with specified parameters and safety measures
                workbook = load_workbook(
                    filename=normalized_path,
                    read_only=read_only,
                    data_only=data_only,
                    keep_vba=False,  # Don't load VBA for performance
                    keep_links=False  # Don't load external links for safety
                )
                
                # Create cache entry
                cache_entry = {
                    'workbook': workbook,
                    'file_path': normalized_path,
                    'file_mtime': os.path.getmtime(normalized_path),
                    'cache_time': time.time(),
                    'read_only': read_only,
                    'data_only': data_only
                }
                
                # Add to cache
                self.cache[cache_key] = cache_entry
                
                # Enforce cache size limit
                self._enforce_cache_limit()
                
                return workbook
                
            except Exception as e:
                self._stats['errors'] += 1
                print(f"Cache ERROR loading {os.path.basename(file_path)}: {e}")
                raise
    
    def _is_cache_valid(self, cached_item, file_path):
        """
        Check if a cached workbook is still valid
        
        Args:
            cached_item: The cached workbook entry
            file_path: Path to the original file
            
        Returns:
            bool: True if cache is valid, False otherwise
        """
        try:
            # Check age limit
            if time.time() - cached_item['cache_time'] > self.max_age_seconds:
                return False
            
            # Check if file has been modified
            current_mtime = os.path.getmtime(file_path)
            if current_mtime != cached_item['file_mtime']:
                return False
            
            # Check if file still exists
            if not os.path.exists(file_path):
                return False
            
            return True
            
        except (OSError, KeyError):
            return False
    
    def _enforce_cache_limit(self):
        """
        Remove oldest entries if cache exceeds size limit
        """
        while len(self.cache) > self.max_size:
            # Remove least recently used item
            oldest_key, oldest_item = self.cache.popitem(last=False)
            self._stats['evictions'] += 1
            print(f"Cache EVICTED: {os.path.basename(oldest_item['file_path'])}")
            
            # Close the workbook if it's not read-only to free memory
            try:
                if hasattr(oldest_item['workbook'], 'close'):
                    oldest_item['workbook'].close()
            except:
                pass  # Ignore close errors
    
    def clear(self):
        """
        Clear all cached workbooks
        """
        with self.lock:
            # Close all workbooks safely
            for cache_entry in self.cache.values():
                try:
                    workbook = cache_entry['workbook']
                    if hasattr(workbook, 'close'):
                        workbook.close()
                    # Force garbage collection for better memory cleanup
                    del workbook
                except Exception as e:
                    print(f"Warning: Could not close workbook properly: {e}")
                    pass
            
            self.cache.clear()
            print("Cache CLEARED")
    
    def remove(self, file_path):
        """
        Remove a specific file from cache
        
        Args:
            file_path: Path to the file to remove from cache
        """
        normalized_path = os.path.normpath(os.path.abspath(file_path))
        
        with self.lock:
            keys_to_remove = [key for key in self.cache.keys() if key.startswith(normalized_path + "|")]
            
            for key in keys_to_remove:
                cache_entry = self.cache[key]
                try:
                    if hasattr(cache_entry['workbook'], 'close'):
                        cache_entry['workbook'].close()
                except:
                    pass
                del self.cache[key]
                print(f"Cache REMOVED: {os.path.basename(file_path)}")
    
    def get_stats(self):
        """
        Get cache statistics
        
        Returns:
            dict: Cache statistics including hits, misses, etc.
        """
        with self.lock:
            total_requests = self._stats['hits'] + self._stats['misses']
            hit_rate = (self._stats['hits'] / total_requests * 100) if total_requests > 0 else 0
            
            return {
                'cache_size': len(self.cache),
                'max_size': self.max_size,
                'hits': self._stats['hits'],
                'misses': self._stats['misses'],
                'evictions': self._stats['evictions'],
                'errors': self._stats['errors'],
                'hit_rate_percent': round(hit_rate, 2),
                'cached_files': [os.path.basename(entry['file_path']) for entry in self.cache.values()]
            }
    
    def print_stats(self):
        """
        Print cache statistics to console
        """
        stats = self.get_stats()
        print("\n=== Workbook Cache Statistics ===")
        print(f"Cache Size: {stats['cache_size']}/{stats['max_size']}")
        print(f"Hit Rate: {stats['hit_rate_percent']}%")
        print(f"Hits: {stats['hits']}, Misses: {stats['misses']}")
        print(f"Evictions: {stats['evictions']}, Errors: {stats['errors']}")
        if stats['cached_files']:
            print(f"Cached Files: {', '.join(stats['cached_files'])}")
        print("================================\n")


# Global cache instance
_global_cache = None
_cache_lock = threading.Lock()


def get_global_cache():
    """
    Get the global workbook cache instance (singleton pattern)
    
    Returns:
        WorkbookCache: The global cache instance
    """
    global _global_cache
    
    if _global_cache is None:
        with _cache_lock:
            if _global_cache is None:
                _global_cache = WorkbookCache(max_size=15, max_age_seconds=600)  # 10 minutes
    
    return _global_cache


def clear_global_cache():
    """
    Clear the global cache
    """
    global _global_cache
    
    if _global_cache is not None:
        _global_cache.clear()


def get_cached_workbook(file_path, read_only=True, data_only=True):
    """
    Convenience function to get a workbook using the global cache
    
    Args:
        file_path: Path to the Excel file
        read_only: Whether to open in read-only mode
        data_only: Whether to read only calculated values
        
    Returns:
        openpyxl.Workbook: The loaded workbook
    """
    cache = get_global_cache()
    return cache.get_workbook(file_path, read_only, data_only)


def print_cache_stats():
    """
    Print global cache statistics
    """
    cache = get_global_cache()
    cache.print_stats()


# Test function
if __name__ == "__main__":
    # Simple test
    cache = WorkbookCache(max_size=3)
    print("Workbook cache system initialized successfully!")
    cache.print_stats()