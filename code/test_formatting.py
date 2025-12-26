#!/usr/bin/env python3
import pandas as pd
import sys
from pathlib import Path

# Add a test participant to the participants file
participants_file = Path("Turnir 2025/Ucesnici 2025.xlsx")

try:
    # Read current participants
    df = pd.read_excel(participants_file)
    
    # Check if test participant already exists
    test_name = "Test Formatting Player"
    if test_name not in df['Име'].values:
        # Add test participant with payment
        new_row = {
            'Име': test_name,
            'Уплаћено учешће': 30.0,
            'Метод плаћања': None,
            'Сигурно долази': None,
            'Можда долази': None,
            'Гарнитура': None,
            'Сат': None
        }
        
        # Add the new row
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
        # Save the file (this might fail if file is open, but that's OK for testing)
        try:
            df.to_excel(participants_file, index=False)
            print(f"Added test participant: {test_name}")
        except:
            print("Could not save participants file (probably open in Excel)")
            print("Manually add a test participant with payment to test formatting")
    else:
        print(f"Test participant {test_name} already exists")

except Exception as e:
    print(f"Error: {e}")
    print("Could not modify participants file")

print("\nNow run the main script to test formatting preservation:")
print("python code/azuriraj_ucesnike.py \"Turnir 2025\"")
