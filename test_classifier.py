import sys
from pathlib import Path
sys.path.insert(0, str(Path.cwd()))
from parsers.registry import registry

def test():
    file = Path("../Software Logs/Ansys log/license from 27 feb-2025 to 27-02-2026.log")
    if file.exists():
        info = registry.detect_and_report(file)
        print("Explicit file detection:")
        print(f"  Classified as: {info['classified_as']}")
        print(f"  Vendor: {info['vendor']}")
    else:
        print("File not found to test.")

if __name__ == '__main__':
    test()
