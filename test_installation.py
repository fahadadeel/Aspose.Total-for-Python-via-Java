#!/usr/bin/env python3
"""
Simple test script for Aspose Document Processor
"""

def test_imports():
    """Test basic imports"""
    print("🧪 Testing imports...")
    
    try:
        import sys
        print(f"✅ Python {sys.version.split()[0]} is working")
    except Exception as e:
        print(f"❌ Python issue: {e}")
        return False
    
    try:
        from main import AsposeDocumentProcessor
        print("✅ Main module imports successfully")
    except Exception as e:
        print(f"❌ Main module import failed: {e}")
        return False
    
    try:
        processor = AsposeDocumentProcessor()
        print("✅ Processor initialized")
    except Exception as e:
        print(f"❌ Processor initialization failed: {e}")
        print("   This is expected if no JAR files are present")
        return False
    
    try:
        formats = processor.get_supported_extensions()
        print(f"✅ Supported formats: {formats}")
    except Exception as e:
        print(f"❌ Could not get supported formats: {e}")
        return False
    
    try:
        processors = processor.list_processors()
        print(f"✅ Available processors: {list(processors.keys())}")
    except Exception as e:
        print(f"❌ Could not list processors: {e}")
        return False
    
    return True

def test_jar_detection():
    """Test JAR file detection"""
    print("\n🧪 Testing JAR detection...")
    
    import os
    from pathlib import Path
    
    plugins_dir = Path("plugins")
    jar_files = list(plugins_dir.glob("*.jar"))
    
    if jar_files:
        print(f"✅ Found {len(jar_files)} JAR files:")
        for jar in jar_files:
            print(f"   - {jar.name}")
        return True
    else:
        print("⚠️  No JAR files found in plugins/ directory")
        print("📥 Download Aspose.Cells from: https://downloads.aspose.com/cells/java")
        return False

def main():
    print("Aspose Document Processor - Test Script")
    print("======================================")
    
    # Test imports
    imports_ok = test_imports()
    
    # Test JAR detection
    jars_ok = test_jar_detection()
    
    print("\n📋 Summary:")
    print(f"   Imports: {'✅ OK' if imports_ok else '❌ Failed'}")
    print(f"   JAR files: {'✅ Found' if jars_ok else '⚠️  Missing'}")
    
    if imports_ok and jars_ok:
        print("\n🎉 All tests passed! Your installation is ready.")
    elif imports_ok:
        print("\n⚠️  Imports work but no JAR files found.")
        print("   Download Aspose.Cells JAR to complete setup.")
    else:
        print("\n❌ There are issues with your installation.")
        print("   Check the error messages above.")

if __name__ == "__main__":
    main()