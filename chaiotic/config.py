"""Configuration module for the application."""

import os
import json
from typing import Any, Dict, Optional
from pathlib import Path
import traceback

# Load environment variables from .env file
def load_env_file():
    """Load environment variables from .env file."""
    env_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.env')
    if os.path.exists(env_path):
        print(f"Found .env file at {env_path}")
        try:
            with open(env_path, 'r') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#') or line.startswith('//'):
                        continue
                    if '=' in line:
                        key, value = line.split('=', 1)
                        os.environ[key] = value
                        print(f"Loaded environment variable: {key}")
        except Exception as e:
            print(f"Error loading .env file: {e}")
            traceback.print_exc()
    else:
        print(f"No .env file found at {env_path}")

# Load environment variables at module import time
load_env_file()

# Constants - read from environment or use defaults
GPT4O_MINI_MODEL = os.environ.get("GPT4O_MINI_MODEL", "gpt-4o-mini")
GPT4_TURBO_MODEL = os.environ.get("GPT4_MODEL", "gpt-4o")
DISABLE_CACHE = False

class Config:
    """Configuration class for the application."""
    
    def __init__(self):
        """Initialize the configuration."""
        self._config = {}
        self._cache_enabled = not DISABLE_CACHE
        self._load_config()
    
    def _load_config(self):
        """Load configuration from file if exists."""
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json')
        if os.path.exists(config_file):
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    self._config = json.load(f)
            except Exception as e:
                print(f"Error loading config: {e}")
                self._config = {}
    
    def save_config(self):
        """Save configuration to file."""
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'config.json')
        try:
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=4)
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def set_cache_enabled(self, enabled: bool):
        """Set whether caching is enabled."""
        self._cache_enabled = enabled
    
    def is_cache_enabled(self) -> bool:
        """Check if caching is enabled."""
        return self._cache_enabled
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value with a default fallback."""
        return self._config.get(key, default)
    
    def set(self, key: str, value: Any):
        """Set a configuration value."""
        self._config[key] = value
        self.save_config()
    
    def __getattr__(self, name: str) -> Any:
        """Allow accessing config values as attributes."""
        return self._config.get(name)
    
    def __setattr__(self, name: str, value: Any):
        """Allow setting config values as attributes."""
        if name.startswith('_'):
            # Private attributes are set directly
            super().__setattr__(name, value)
        else:
            # Public attributes are stored in the config
            self._config[name] = value
            self.save_config()

_config_instance = None

def load_config() -> Config:
    """Load the configuration."""
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    return _config_instance