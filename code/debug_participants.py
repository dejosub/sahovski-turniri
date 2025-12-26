#!/usr/bin/env python3
import pandas as pd

# Read the tournament file
df = pd.read_excel('Turnir 2025/Turnir 2025.xlsm')

print("Tournament file analysis:")
print(f"Total rows: {len(df)}")
print(f"First column name: '{df.columns[0]}'")
print()

# Check existing participants (using improved logic)
existing_participants = []
player_name_column = df.columns[0]

print("Checking first 5 rows for actual participants:")
for idx in range(min(5, len(df))):
    name = df.iloc[idx][player_name_column]
    if pd.notna(name):
        starts_with_player = str(name).startswith("Играч ")
        name_str = str(name)
        is_actual_player = (not starts_with_player and 
                           not name_str in ["Непар", "Унос играча", "Ресетовање", "Генериисање парова", 
                                          "Ажурирање табеле", "Касни почеци", "Паузирање"] and
                           not name_str.startswith("Играчи који не треба") and
                           not name_str.isdigit() and
                           len(name_str) > 3)
        
        print(f"Row {idx}: '{name}' - Placeholder: {starts_with_player}, Actual player: {is_actual_player}")
        
        if is_actual_player:
            existing_participants.append(name)

print(f"\nExisting participants (non-placeholder): {existing_participants}")
print(f"Total existing: {len(existing_participants)}")

# Check placeholders
placeholders = []
for idx, row in df.iterrows():
    name = row[player_name_column]
    if pd.notna(name) and str(name).startswith("Играч "):
        placeholders.append((idx, name))

print(f"\nPlaceholder slots: {placeholders[:10]}...")  # Show first 10
print(f"Total placeholders: {len(placeholders)}")
