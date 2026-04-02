import sys
from pathlib import Path
sys.path.insert(0, str(Path.cwd()))
from parsers.registry import registry

def test_autoscan():
    target = Path("../Software Logs")
    for f in target.rglob("*"):
        if f.is_file() and not f.name.startswith("."):
            print(f"File: {f.relative_to(target.parent)}")
            info = registry.detect_and_report(f)
            print(f"  Detected -> {info['software_name']} (Key: {info['classified_as']}, Confidence: {info['confidence']})")

if __name__ == '__main__':
    test_autoscan()
