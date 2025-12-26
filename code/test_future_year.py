#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script to demonstrate year-independence of azuriraj_ucesnike.py
"""

import subprocess
import sys
from pathlib import Path

def test_script_with_different_folders():
    """Test the script with different tournament folder scenarios"""
    
    script_path = Path(__file__).parent / "azuriraj_ucesnike.py"
    
    test_cases = [
        # Test with relative path
        ("Turnir 2025", "Relative path test"),
        # Test with absolute path  
        (str(Path(__file__).parent.parent / "Turnir 2025"), "Absolute path test"),
        # Test without parameter (fallback mode)
        (None, "Fallback mode test")
    ]
    
    print("=== Testing Tournament Folder Parameter Handling ===\n")
    
    for i, (folder_param, description) in enumerate(test_cases, 1):
        print(f"Test {i}: {description}")
        print(f"Parameter: {folder_param or 'None (fallback)'}")
        
        # Build command
        if folder_param:
            cmd = [sys.executable, str(script_path), folder_param]
        else:
            cmd = [sys.executable, str(script_path)]
        
        try:
            # Run the script and capture output
            result = subprocess.run(cmd, 
                                  capture_output=True, 
                                  text=True, 
                                  cwd=script_path.parent.parent,
                                  timeout=30)
            
            print(f"Exit code: {result.returncode}")
            print("Output:")
            print(result.stdout)
            if result.stderr:
                print("Errors:")
                print(result.stderr)
                
        except subprocess.TimeoutExpired:
            print("Script timed out (probably waiting for file access)")
        except Exception as e:
            print(f"Error running script: {e}")
        
        print("-" * 50)

if __name__ == "__main__":
    test_script_with_different_folders()
