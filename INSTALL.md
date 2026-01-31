# Installation & Quick Start Guide

## Three Ways to Run the Application

### Method 1: Simple Run (Recommended for Quick Start)

**Windows:**
```cmd
cd improved_solver
pip install -r requirements.txt
python run.py
```

**Mac/Linux:**
```bash
cd improved_solver
pip install -r requirements.txt
python run.py
```

This will automatically launch the Streamlit app in your browser.

---

### Method 2: Install as Package (Recommended for Development)

**Step 1: Install the package**
```bash
cd improved_solver
pip install -e .
```

**Step 2: Run the app**
```bash
cd src
streamlit run app.py
```

---

### Method 3: Direct Streamlit Run

**From the improved_solver directory:**
```bash
pip install -r requirements.txt
cd src
streamlit run app.py
```

If you get import errors, try this instead:
```bash
# From improved_solver root directory
export PYTHONPATH="${PYTHONPATH}:$(pwd)"  # Mac/Linux
# OR
set PYTHONPATH=%PYTHONPATH%;%CD%  # Windows

cd src
streamlit run app.py
```

---

## Troubleshooting Import Errors

### Error: "ModuleNotFoundError: No module named 'utils.file_handlers'"

**Solution 1** (Easiest):
```bash
# Use the run.py script
python run.py
```

**Solution 2**:
```bash
# Install as package
pip install -e .
cd src
streamlit run app.py
```

**Solution 3**:
```bash
# Run from the root directory with PYTHONPATH
# Windows:
set PYTHONPATH=%CD%
streamlit run src/app.py

# Mac/Linux:
PYTHONPATH=$(pwd) streamlit run src/app.py
```

---

## Verifying Installation

### Test 1: Check Python Version
```bash
python --version
# Should be 3.8 or higher
```

### Test 2: Check Dependencies
```bash
pip list | grep streamlit
pip list | grep cvxpy
pip list | grep pandas
```

### Test 3: Run Tests
```bash
pytest tests/ -v
```

If tests pass, everything is installed correctly!

---

## Common Issues

### Issue: "pip install fails"

**Solution:**
```bash
# Update pip first
python -m pip install --upgrade pip

# Install dependencies one by one
pip install numpy pandas
pip install cvxpy
pip install streamlit openpyxl
```

### Issue: "streamlit command not found"

**Solution:**
```bash
# Use python -m streamlit instead
python -m streamlit run src/app.py
```

### Issue: "No module named 'cvxpy'"

**Solution:**
```bash
pip install cvxpy
# If that fails:
pip install --upgrade pip setuptools wheel
pip install cvxpy
```

### Issue: Port already in use

**Solution:**
```bash
# Use a different port
streamlit run src/app.py --server.port=8502
```

---

## Quick Reference

### File Structure
```
improved_solver/
├── run.py                 ← Use this for easy launch
├── setup.py               ← For package installation
├── requirements.txt       ← Dependencies
├── src/
│   └── app.py            ← Main application
├── config/
│   └── settings.py       ← Configuration
├── optimization/
│   └── optimizer.py      ← Core optimization
├── utils/
│   ├── file_handlers.py  ← File I/O
│   └── session_manager.py
├── visualization/
│   ├── excel_generator.py
│   └── html_generator.py
├── tests/
│   └── test_optimizer.py ← Test suite
└── docs/
    └── README.md         ← Full documentation
```

### Launch Commands Summary

| Method | Command | Notes |
|--------|---------|-------|
| Easiest | `python run.py` | Recommended for first-time users |
| Package | `pip install -e . && cd src && streamlit run app.py` | Best for development |
| Direct | `streamlit run src/app.py` | Requires PYTHONPATH setup |

---

## Testing Your Installation

### Quick Test Script

Create a file called `test_install.py`:

```python
#!/usr/bin/env python
"""Test if all dependencies are installed correctly."""

import sys

def test_imports():
    """Test all required imports."""
    print("Testing imports...")
    
    try:
        import streamlit as st
        print("✓ streamlit")
    except ImportError as e:
        print(f"✗ streamlit: {e}")
        return False
    
    try:
        import pandas as pd
        print("✓ pandas")
    except ImportError as e:
        print(f"✗ pandas: {e}")
        return False
    
    try:
        import numpy as np
        print("✓ numpy")
    except ImportError as e:
        print(f"✗ numpy: {e}")
        return False
    
    try:
        import cvxpy as cp
        print("✓ cvxpy")
    except ImportError as e:
        print(f"✗ cvxpy: {e}")
        return False
    
    try:
        import openpyxl
        print("✓ openpyxl")
    except ImportError as e:
        print(f"✗ openpyxl: {e}")
        return False
    
    print("\n✅ All dependencies installed correctly!")
    return True

if __name__ == "__main__":
    success = test_imports()
    sys.exit(0 if success else 1)
```

Run it:
```bash
python test_install.py
```

---

## Getting Help

If you're still having issues:

1. **Check Python version**: Must be 3.8 or higher
2. **Use virtual environment**: Recommended to avoid conflicts
3. **Try the run.py script**: Easiest way to launch
4. **Check the full documentation**: See `docs/README.md`
5. **Run the test suite**: `pytest tests/ -v`

---

## Virtual Environment Setup (Recommended)

### Windows:
```cmd
cd improved_solver
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

### Mac/Linux:
```bash
cd improved_solver
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python run.py
```

This ensures a clean environment without conflicts.

---

## Success Checklist

- [ ] Python 3.8+ installed
- [ ] Dependencies installed (`pip install -r requirements.txt`)
- [ ] Test imports pass (`python test_install.py`)
- [ ] Can launch app (`python run.py`)
- [ ] Browser opens to http://localhost:8501
- [ ] Can upload sample data
- [ ] Can run optimization

If all checked, you're ready to go! 🎉

---

**Need more help?** See `docs/TROUBLESHOOTING.md` for detailed solutions.
