"""
Generate Machine Hash for License Registration
This script generates the machine hash that needs to be registered.
"""
import sys
sys.path.insert(0, r'c:\Users\pc\autoarendt')

from services.license_validator import LicenseValidator

def main():
    validator = LicenseValidator()
    info = validator.get_machine_info()
    
    print("=" * 70)
    print("MACHINE INFORMATION FOR LICENSE REGISTRATION")
    print("=" * 70)
    print()
    print(f"Machine Hash:  {info['hash']}")
    print(f"Hostname:      {info['hostname']}")
    print(f"System:        {info['system']}")
    print(f"Platform:      {info['platform']}")
    print(f"Processor:     {info['processor']}")
    print()
    print("=" * 70)
    print("TO REGISTER THIS MACHINE:")
    print("=" * 70)
    print()
    print(f"1. Create a file named: {info['hash']}.txt")
    print(f"2. Add any content (e.g., 'licensed', hostname, date, etc.)")
    print(f"3. Upload to: https://alemxral.github.io/cv/")
    print(f"4. The file should be accessible at:")
    print(f"   https://alemxral.github.io/cv/{info['hash']}.txt")
    print()
    print("=" * 70)
    
    # Save to file for easy reference
    output_file = "machine_license_info.txt"
    with open(output_file, 'w') as f:
        f.write(f"Machine Hash: {info['hash']}\n")
        f.write(f"Hostname: {info['hostname']}\n")
        f.write(f"System: {info['system']}\n")
        f.write(f"Platform: {info['platform']}\n")
        f.write(f"Processor: {info['processor']}\n")
        f.write(f"\nLicense file to create: {info['hash']}.txt\n")
        f.write(f"Upload to: https://alemxral.github.io/cv/\n")
    
    print(f"✅ Information saved to: {output_file}")

if __name__ == '__main__':
    main()
