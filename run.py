"""
Launcher script for the Sample Allocation Optimizer.

This script ensures proper Python path setup and launches the Streamlit app.
"""

import sys
from pathlib import Path

# Add the project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

# Now run the Streamlit app
if __name__ == "__main__":
    import streamlit.web.cli as stcli
    
    app_path = project_root / "src" / "app.py"
    
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        "--server.port=8501",
        "--server.address=localhost"
    ]
    
    sys.exit(stcli.main())
