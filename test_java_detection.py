#!/usr/bin/env python3
"""
Test script to check Java installation and tabula-py functionality
"""
import subprocess
import sys

def check_java_installation():
    """Check if Java is installed and accessible"""
    print("Checking Java installation...")
    try:
        result = subprocess.run(['java', '-version'], 
                              capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print("✅ Java is installed and accessible")
            print("Java version info:")
            print(result.stderr)  # Java version info goes to stderr
            return True
        else:
            print("❌ Java command failed")
            return False
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError) as e:
        print(f"❌ Java not found: {e}")
        return False

def check_tabula_py():
    """Check if tabula-py is available"""
    print("\nChecking tabula-py availability...")
    try:
        import tabula
        print("✅ tabula-py is available")
        return True
    except ImportError as e:
        print(f"❌ tabula-py not available: {e}")
        return False

def check_pdfplumber():
    """Check if pdfplumber is available"""
    print("\nChecking pdfplumber availability...")
    try:
        import pdfplumber
        print("✅ pdfplumber is available")
        return True
    except ImportError as e:
        print(f"❌ pdfplumber not available: {e}")
        return False

def main():
    print("PDF Tool Dependency Check")
    print("=" * 40)
    
    java_ok = check_java_installation()
    tabula_ok = check_tabula_py()
    pdfplumber_ok = check_pdfplumber()
    
    print("\n" + "=" * 40)
    print("SUMMARY:")
    
    if java_ok and tabula_ok:
        print("✅ Best case: Java + tabula-py available")
        print("   PDF to Excel will use tabula-py for optimal results")
    elif pdfplumber_ok:
        print("⚠️  Fallback mode: pdfplumber available")
        print("   PDF to Excel will use pdfplumber (still good results)")
    else:
        print("❌ Limited functionality: Missing dependencies")
        print("   PDF to Excel will use basic text extraction only")
    
    print("\nRecommendations:")
    if not java_ok:
        print("- Install Java JRE 8+ for best PDF table extraction")
    if not tabula_ok:
        print("- Install tabula-py: pip install tabula-py")
    if not pdfplumber_ok:
        print("- Install pdfplumber: pip install pdfplumber")

if __name__ == "__main__":
    main()