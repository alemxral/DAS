"""Test license validation."""
import sys
sys.path.insert(0, r'c:\Users\pc\autoarendt')

from services.license_validator import LicenseValidator

validator = LicenseValidator()

print(f"\nTesting license check...")
print(f"Hash: {validator.hash}")
print(f"URL: {validator.base_url}/{validator.hash}.txt")
print("-" * 60)

is_valid, message = validator.validate()

print("-" * 60)
if is_valid:
    print(f"[OK] Result: APPROVED - All users can access")
else:
    print(f"[ERROR] Result: DENIED - All users are blocked")
print(f"Message: {message}")
print("-" * 60)
print(f"\nTo enable access for ALL users:")
print(f"  Create file: {validator.hash}.txt")
print(f"  Upload to: {validator.base_url}/")
print(f"\nTo block access for ALL users:")
print(f"  Delete file: {validator.hash}.txt from GitHub Pages")

