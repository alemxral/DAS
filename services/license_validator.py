"""
License Validation Service
Validates application license against remote server.
"""
import hashlib
import requests
from typing import Tuple

class LicenseValidator:
    """Validates application license."""
    
    def __init__(self):
        """Initialize license validator."""
        self.base_url = "https://alemxral.github.io/cv"
        self.timeout = 10
        # Fixed hash for all installations
        self.hash = "das_license_2025"
    
    def validate(self) -> Tuple[bool, str]:
        """
        Validate license against remote server.
        Simple ping: if URL returns 200, license is valid.
        
        Returns:
            Tuple of (is_valid, message)
        """
        try:
            url = f"{self.base_url}/{self.hash}.txt"
            
            print(f"[License] Checking: {url}")
            
            response = requests.get(url, timeout=self.timeout)
            
            if response.status_code == 200:
                print(f"[License] ✅ Access granted (HTTP 200)")
                return True, "License valid"
            else:
                print(f"[License] ❌ Access denied (HTTP {response.status_code})")
                return False, "Service not available"
                
        except requests.exceptions.Timeout:
            print(f"[License] ❌ Timeout")
            return False, "Service not available - timeout"
        except requests.exceptions.ConnectionError:
            print(f"[License] ❌ Connection error")
            return False, "Service not available - no connection"
        except Exception as e:
            print(f"[License] ❌ Error: {e}")
            return False, f"Service not available"
