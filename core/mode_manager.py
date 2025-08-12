# -*- coding: utf-8 -*-
"""
Mode Manager Module

This module manages the switching between Normal Mode and Inspect Mode.
Handles window configuration, layout changes, and mode state management.
"""

import tkinter as tk
from tkinter import ttk
from enum import Enum

class AppMode(Enum):
    """Application mode enumeration"""
    NORMAL = "normal"
    INSPECT = "inspect"

class ModeManager:
    """Manages application mode switching and window configuration"""
    
    def __init__(self, root_window):
        """
        Initialize the Mode Manager
        
        Args:
            root_window: The main tkinter window
        """
        self.root = root_window
        self.current_mode = AppMode.NORMAL
        self.always_on_top = False
        
        # Mode configurations
        self.mode_configs = {
            AppMode.NORMAL: {
                "size": "900x1100",
                "title": "Excel Tools - Integrated",
                "layout": "notebook",
                "resizable": (True, True)
            },
            AppMode.INSPECT: {
                "size": "1200x600", 
                "title": "Excel Tools - Inspect Mode",
                "layout": "dual_pane",
                "resizable": (True, True)
            }
        }
        
        # Store original window state
        self._original_geometry = None
        self._original_title = None
        
        # Callbacks for mode switching
        self._mode_switch_callbacks = []
    
    def get_current_mode(self):
        """Get the current application mode"""
        return self.current_mode
    
    def is_inspect_mode(self):
        """Check if currently in inspect mode"""
        return self.current_mode == AppMode.INSPECT
    
    def is_normal_mode(self):
        """Check if currently in normal mode"""
        return self.current_mode == AppMode.NORMAL
    
    def get_mode_config(self, mode=None):
        """
        Get configuration for a specific mode
        
        Args:
            mode: AppMode enum, defaults to current mode
            
        Returns:
            dict: Mode configuration
        """
        if mode is None:
            mode = self.current_mode
        return self.mode_configs.get(mode, {})
    
    def register_mode_switch_callback(self, callback):
        """
        Register a callback to be called when mode switches
        
        Args:
            callback: Function to call with (old_mode, new_mode) parameters
        """
        if callback not in self._mode_switch_callbacks:
            self._mode_switch_callbacks.append(callback)
    
    def unregister_mode_switch_callback(self, callback):
        """Unregister a mode switch callback"""
        if callback in self._mode_switch_callbacks:
            self._mode_switch_callbacks.remove(callback)
    
    def _notify_mode_switch(self, old_mode, new_mode):
        """Notify all registered callbacks about mode switch"""
        for callback in self._mode_switch_callbacks:
            try:
                callback(old_mode, new_mode)
            except Exception as e:
                print(f"Error in mode switch callback: {e}")
    
    def switch_to_normal_mode(self):
        """Switch to Normal Mode"""
        if self.current_mode == AppMode.NORMAL:
            return  # Already in normal mode
        
        old_mode = self.current_mode
        self.current_mode = AppMode.NORMAL
        
        # Apply window configuration
        self._apply_window_config()
        
        # Notify callbacks
        self._notify_mode_switch(old_mode, self.current_mode)
    
    def switch_to_inspect_mode(self):
        """Switch to Inspect Mode"""
        if self.current_mode == AppMode.INSPECT:
            return  # Already in inspect mode
        
        # Store original state if switching from normal mode
        if self.current_mode == AppMode.NORMAL:
            self._store_original_state()
        
        old_mode = self.current_mode
        self.current_mode = AppMode.INSPECT
        
        # Apply window configuration
        self._apply_window_config()
        
        # Notify callbacks
        self._notify_mode_switch(old_mode, self.current_mode)
    
    def toggle_mode(self):
        """Toggle between Normal and Inspect modes"""
        if self.current_mode == AppMode.NORMAL:
            self.switch_to_inspect_mode()
        else:
            self.switch_to_normal_mode()
    
    def _store_original_state(self):
        """Store the original window state"""
        self._original_geometry = self.root.geometry()
        self._original_title = self.root.title()
    
    def _apply_window_config(self):
        """Apply window configuration for current mode"""
        config = self.get_mode_config()
        
        if not config:
            return
        
        # Apply window size
        if "size" in config:
            self.root.geometry(config["size"])
        
        # Apply window title
        if "title" in config:
            self.root.title(config["title"])
        
        # Apply resizable settings
        if "resizable" in config:
            self.root.resizable(*config["resizable"])
        
        # Apply always on top if enabled
        self._apply_always_on_top()
    
    def set_always_on_top(self, enabled):
        """
        Set always on top state
        
        Args:
            enabled (bool): Whether to keep window always on top
        """
        self.always_on_top = enabled
        self._apply_always_on_top()
    
    def toggle_always_on_top(self):
        """Toggle always on top state"""
        self.set_always_on_top(not self.always_on_top)
    
    def _apply_always_on_top(self):
        """Apply the always on top setting"""
        try:
            self.root.attributes("-topmost", self.always_on_top)
        except Exception as e:
            print(f"Error setting topmost attribute: {e}")
    
    def restore_original_state(self):
        """Restore the original window state (if stored)"""
        if self._original_geometry:
            self.root.geometry(self._original_geometry)
        
        if self._original_title:
            self.root.title(self._original_title)
        
        # Reset always on top
        self.set_always_on_top(False)
    
    def get_window_center_position(self, width, height):
        """
        Calculate center position for window
        
        Args:
            width (int): Window width
            height (int): Window height
            
        Returns:
            tuple: (x, y) position for centering
        """
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        
        return x, y
    
    def center_window(self, mode=None):
        """
        Center the window for the specified mode
        
        Args:
            mode: AppMode enum, defaults to current mode
        """
        config = self.get_mode_config(mode)
        size_str = config.get("size", "900x1100")
        
        try:
            # Parse size string (e.g., "1200x600")
            width, height = map(int, size_str.split('x'))
            x, y = self.get_window_center_position(width, height)
            
            # Apply centered geometry
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            
        except Exception as e:
            print(f"Error centering window: {e}")
    
    def get_status_info(self):
        """
        Get current mode status information
        
        Returns:
            dict: Status information
        """
        return {
            "current_mode": self.current_mode.value,
            "always_on_top": self.always_on_top,
            "window_geometry": self.root.geometry(),
            "window_title": self.root.title()
        }

# Convenience functions for external use
def create_mode_manager(root_window):
    """Create and return a new ModeManager instance"""
    return ModeManager(root_window)

def is_inspect_mode_available():
    """Check if inspect mode functionality is available"""
    # Could be used for feature flags or conditional loading
    return True