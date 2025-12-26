#!/usr/bin/env python3
"""
Test script for template creation functionality
"""

import sys
from pathlib import Path
import shutil

# Add the code directory to path so we can import our functions
sys.path.append(str(Path(__file__).parent))

from azuriraj_ucesnike import create_tournament_file_from_template, find_tournament_file

def test_template_creation():
    """Test the template creation functionality"""
    
    # Test with a hypothetical folder
    test_folder = Path("c:/Users/dejan/code/dejosub/sahovski-turniri/Turnir 2026")
    
    print("=== Template Creation Test ===")
    print(f"Test folder: {test_folder}")
    print(f"Folder exists: {test_folder.exists()}")
    
    # Check if tournament file exists
    existing_file = find_tournament_file(test_folder)
    print(f"Existing tournament file: {existing_file}")
    
    if not existing_file and test_folder.exists():
        print("\nTesting template creation...")
        new_file = create_tournament_file_from_template(test_folder)
        print(f"Created file: {new_file}")
        
        if new_file and Path(new_file).exists():
            print("✅ Template creation successful!")
            print(f"File size: {Path(new_file).stat().st_size} bytes")
        else:
            print("❌ Template creation failed!")
    else:
        print("Skipping template creation test (folder doesn't exist or file already exists)")
    
    # Test ID extraction
    test_cases = [
        "Turnir 2026",
        "Turnir 2027-Spring", 
        "Turnir Test",
        "Custom Tournament"
    ]
    
    print("\n=== ID Extraction Test ===")
    for folder_name in test_cases:
        if folder_name.startswith("Turnir "):
            tournament_id = folder_name[7:]  # Remove "Turnir " prefix
        else:
            tournament_id = folder_name
        print(f"'{folder_name}' → ID: '{tournament_id}' → File: 'Turnir {tournament_id}.xlsm'")

if __name__ == "__main__":
    test_template_creation()
