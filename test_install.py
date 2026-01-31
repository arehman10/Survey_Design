#!/usr/bin/env python
"""Test if all dependencies are installed correctly."""

import sys

def test_imports():
    """Test all required imports."""
    print("Testing imports...")
    print("-" * 50)
    
    all_ok = True
    
    try:
        import streamlit as st
        print("✓ streamlit:", st.__version__)
    except ImportError as e:
        print(f"✗ streamlit: {e}")
        all_ok = False
    
    try:
        import pandas as pd
        print("✓ pandas:", pd.__version__)
    except ImportError as e:
        print(f"✗ pandas: {e}")
        all_ok = False
    
    try:
        import numpy as np
        print("✓ numpy:", np.__version__)
    except ImportError as e:
        print(f"✗ numpy: {e}")
        all_ok = False
    
    try:
        import cvxpy as cp
        print("✓ cvxpy:", cp.__version__)
    except ImportError as e:
        print(f"✗ cvxpy: {e}")
        all_ok = False
    
    try:
        import openpyxl
        print("✓ openpyxl:", openpyxl.__version__)
    except ImportError as e:
        print(f"✗ openpyxl: {e}")
        all_ok = False
    
    try:
        import matplotlib
        print("✓ matplotlib:", matplotlib.__version__)
    except ImportError as e:
        print(f"✗ matplotlib: {e}")
        all_ok = False
    
    print("-" * 50)
    
    if all_ok:
        print("\n✅ All dependencies installed correctly!")
        print("\nYou can now run the application:")
        print("  python run.py")
        print("\nOr:")
        print("  cd src")
        print("  streamlit run app.py")
    else:
        print("\n❌ Some dependencies are missing.")
        print("\nPlease install them:")
        print("  pip install -r requirements.txt")
    
    return all_ok

def test_python_version():
    """Test Python version."""
    print("\nChecking Python version...")
    print("-" * 50)
    
    version = sys.version_info
    print(f"Python {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("❌ Python 3.8 or higher is required")
        return False
    else:
        print("✅ Python version OK")
        return True

if __name__ == "__main__":
    print("=" * 50)
    print("Sample Allocation Optimizer - Installation Test")
    print("=" * 50)
    
    python_ok = test_python_version()
    imports_ok = test_imports()
    
    print("\n" + "=" * 50)
    if python_ok and imports_ok:
        print("Installation Status: ✅ READY")
        sys.exit(0)
    else:
        print("Installation Status: ❌ INCOMPLETE")
        sys.exit(1)
