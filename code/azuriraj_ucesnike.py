#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Chess Tournament Participant Update Script
Transfers paid participants from Ucesnici file to tournament bracket file
"""

import pandas as pd
import os
import sys
import time
import shutil
from pathlib import Path
try:
    import win32com.client
    EXCEL_COM_AVAILABLE = True
except ImportError:
    EXCEL_COM_AVAILABLE = False
    print("Warning: win32com not available. Install with: pip install pywin32")

def load_ratings_lookup(ratings_file):
    """Load player ratings from CSV file into a dictionary"""
    ratings = {}
    try:
        df = pd.read_csv(ratings_file, header=None, names=['name', 'rating'])
        for _, row in df.iterrows():
            ratings[row['name']] = int(row['rating'])
        print(f"Loaded {len(ratings)} player ratings from {ratings_file}")
    except Exception as e:
        print(f"Warning: Could not load ratings file {ratings_file}: {e}")
    return ratings

def get_paid_participants(ucesnici_file):
    """Get list of participants who have paid their fee"""
    try:
        df = pd.read_excel(ucesnici_file)
        # Check if payment column has any non-null values (indicating payment)
        paid_participants = []
        for _, row in df.iterrows():
            name = row['Име']
            payment_status = row['Уплаћено учешће']
            # Consider participant as paid if there's any value in payment column
            # Skip rows where name is NaN or "Укупно" (total row)
            if pd.notna(name) and name != "Укупно" and pd.notna(payment_status) and payment_status > 0:
                paid_participants.append(name)
        
        print(f"Found {len(paid_participants)} paid participants:")
        for participant in paid_participants:
            print(f"  - {participant}")
        
        return paid_participants
    except Exception as e:
        print(f"Error reading participants file {ucesnici_file}: {e}")
        return []

def find_tournament_file(folder_path):
    """Find the tournament file that starts with 'Turnir'"""
    folder = Path(folder_path)
    for file in folder.glob("Turnir*.xlsx"):
        return str(file)
    for file in folder.glob("Turnir*.xlsm"):
        return str(file)
    return None

def update_excel_via_com(tournament_file, updates, player_slots_end=30):
    """Update Excel file using COM interface (works even when file is open)"""
    if not EXCEL_COM_AVAILABLE:
        return False
    
    try:
        # Connect to Excel application
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True  # Make Excel visible so we can see what's happening
        xl.DisplayAlerts = False  # Suppress alerts
        
        # Try to find if the workbook is already open
        workbook = None
        file_name = Path(tournament_file).name
        
        for wb in xl.Workbooks:
            if wb.Name == file_name:
                workbook = wb
                print(f"Found open workbook: {file_name}")
                break
        
        # If not open, open it
        if workbook is None:
            workbook = xl.Workbooks.Open(tournament_file)
            print(f"Opened workbook: {tournament_file}")
        
        # Get the first worksheet
        worksheet = workbook.Worksheets(1)
        print(f"Working with worksheet: {worksheet.Name}")
        
        # Debug: Print first few rows and columns to understand structure
        print("Worksheet structure (first 5 rows, first 10 columns):")
        for row in range(1, 6):
            row_data = []
            for col in range(1, 11):
                cell_value = worksheet.Cells(row, col).Value
                row_data.append(str(cell_value)[:20] if cell_value else "")
            print(f"Row {row}: {row_data}")
        
        # Apply updates
        updates_applied = 0
        for row_idx, col_name, value in updates:
            print(f"\nProcessing update: row_idx={row_idx}, col_name='{col_name}', value='{value}'")
            
            # Find column index by name (check first row for headers)
            col_idx = None
            for col in range(1, worksheet.UsedRange.Columns.Count + 1):
                header_value = worksheet.Cells(1, col).Value
                if header_value and str(header_value).strip() == str(col_name).strip():
                    col_idx = col
                    print(f"Found column '{col_name}' at index {col_idx}")
                    break
            
            # If not found in first row, try to find by column position for player names
            if col_idx is None and col_name == "Unnamed: 0":
                col_idx = 1  # First column is player names
                print(f"Using first column for player names")
            
            if col_idx:
                # Excel uses 1-based indexing, pandas uses 0-based
                excel_row = row_idx + 2  # +1 for 0-based to 1-based, +1 for header row
                
                # Get current cell and its value
                current_cell = worksheet.Cells(excel_row, col_idx)
                current_value = current_cell.Value
                print(f"Current value at ({excel_row}, {col_idx}): '{current_value}'")
                
                # Use format copying approach to preserve formatting
                try:
                    # Find a reference cell with the same formatting (a placeholder cell)
                    reference_cell = None
                    for ref_row in range(1, player_slots_end + 2):  # +2 for Excel 1-based and header
                        ref_cell = worksheet.Cells(ref_row, col_idx)
                        ref_value = ref_cell.Value
                        if ref_value and str(ref_value).startswith("Играч "):
                            reference_cell = ref_cell
                            break
                    
                    # Update the cell value
                    current_cell.Value = value
                    
                    # Copy formatting from reference cell if found
                    if reference_cell:
                        try:
                            # Copy the format using Excel's built-in copy format functionality
                            reference_cell.Copy()
                            current_cell.PasteSpecial(Paste=-4122)  # xlPasteFormats
                            
                            # Clear the clipboard
                            xl.CutCopyMode = False
                            
                            # Restore the value (PasteSpecial might have changed it)
                            current_cell.Value = value
                            
                            print(f"Copied formatting from reference cell and updated value")
                        except Exception as format_error:
                            print(f"Warning: Could not copy formatting: {format_error}")
                    else:
                        print("Warning: No reference cell found for formatting")
                    
                except Exception as update_error:
                    print(f"Cell update failed: {update_error}")
                    current_cell.Value = value
                
                # Verify the update
                new_value = current_cell.Value
                print(f"Updated cell ({excel_row}, {col_idx}) from '{current_value}' to '{new_value}' (formatting preserved)")
                
                updates_applied += 1
            else:
                print(f"Warning: Could not find column '{col_name}'")
        
        if updates_applied > 0:
            # Save the workbook
            workbook.Save()
            print(f"Workbook saved successfully via COM ({updates_applied} updates applied)")
            return True
        else:
            print("No updates were applied")
            return False
        
    except Exception as e:
        print(f"COM update failed: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            if 'xl' in locals():
                xl.DisplayAlerts = True
        except:
            pass

def save_with_retry(df, file_path, max_retries=3, retry_delay=1):
    """Try to save Excel file with retries and different methods"""
    
    for attempt in range(max_retries):
        try:
            # Method 1: Direct save
            df.to_excel(file_path, index=False)
            print(f"Successfully saved {file_path} (attempt {attempt + 1})")
            return True
            
        except PermissionError as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            
            if attempt < max_retries - 1:
                print(f"Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
                retry_delay *= 2  # Exponential backoff
            
        except Exception as e:
            print(f"Unexpected error on attempt {attempt + 1}: {e}")
            break
    
    # Method 2: Try saving to temporary file then replace
    try:
        temp_file = str(Path(file_path).with_suffix('.tmp.xlsx'))
        df.to_excel(temp_file, index=False)
        
        # Wait a moment then replace
        time.sleep(0.5)
        shutil.move(temp_file, file_path)
        print(f"Successfully saved {file_path} via temporary file")
        return True
        
    except Exception as e:
        print(f"Temporary file method failed: {e}")
    
    return False

def update_tournament_file(tournament_file, paid_participants, ratings_lookup, default_rating=1400):
    """Update tournament file with paid participants and their ratings"""
    try:
        # Read the tournament file
        df = pd.read_excel(tournament_file)
        
        # Get existing participants and find player slot range
        existing_participants = set()
        player_name_column = df.columns[0]  # First column contains player names
        
        # Find the end of player slots (marked by "Непар")
        player_slots_end = len(df)
        for idx in range(len(df)):
            name = df.iloc[idx][player_name_column]
            if pd.notna(name) and str(name) == "Непар":
                player_slots_end = idx
                break
        
        print(f"Player slots range: 0 to {player_slots_end - 1}")
        
        # Check existing participants in player slots only
        for idx in range(player_slots_end):
            name = df.iloc[idx][player_name_column]
            if pd.notna(name) and not str(name).startswith("Играч "):
                # Filter out UI elements and only keep actual player names
                name_str = str(name)
                if (not name_str in ["Унос играча", "Ресетовање", "Генериисање парова", 
                                   "Ажурирање табеле", "Касни почеци", "Паузирање"] and
                    not name_str.startswith("Играчи који не треба") and
                    not name_str.isdigit() and
                    len(name_str) > 3):  # Actual names are longer than 3 characters
                    existing_participants.add(name)
        
        print(f"Found {len(existing_participants)} existing participants in tournament file")
        
        # Process each paid participant and collect updates
        updated_count = 0
        com_updates = []  # For COM method: (row_idx, col_name, value)
        failed_participants = []  # Track participants that couldn't be added
        
        for participant in paid_participants:
            if participant in existing_participants:
                print(f"  Skipping {participant} - already in tournament")
                continue
            
            # Find first placeholder slot (only in player slots range)
            placeholder_found = False
            for idx in range(player_slots_end):
                name = df.iloc[idx][player_name_column]
                if pd.notna(name) and str(name).startswith("Играч "):
                    # Replace placeholder with participant name
                    df.at[idx, player_name_column] = participant
                    
                    # Look up and set rating
                    rating = ratings_lookup.get(participant, default_rating)
                    if 'Pocetni poredak' in df.columns:
                        df.at[idx, 'Pocetni poredak'] = rating
                    
                    # Store updates for COM method
                    com_updates.append((idx, player_name_column, participant))
                    if 'Pocetni poredak' in df.columns:
                        com_updates.append((idx, 'Pocetni poredak', rating))
                    
                    print(f"  Added {participant} with rating {rating} (replacing {name})")
                    updated_count += 1
                    placeholder_found = True
                    break
            
            if not placeholder_found:
                print(f"  ERROR: No available placeholder slot for {participant}")
                print(f"  Tournament file only has {player_slots_end} player slots, all are filled.")
                failed_participants.append(participant)
        
        # Show dialog if some participants couldn't be added
        if failed_participants:
            error_msg = (f"Not enough placeholder slots in tournament file!\n\n"
                        f"Could not add {len(failed_participants)} participants:\n"
                        f"{', '.join(failed_participants)}\n\n"
                        f"Tournament file only has {player_slots_end} player slots.\n"
                        f"Please expand the tournament or remove some existing participants.")
            
            print(f"\nWARNING: {len(failed_participants)} participants could not be added due to insufficient slots")
            
            # Show error dialog
            try:
                import win32api
                win32api.MessageBox(0, error_msg, 'Insufficient Tournament Slots', 0x30)  # Warning icon
            except:
                pass  # If win32api not available, just print to console
        
        if updated_count > 0:
            print(f"\nAttempting to save {updated_count} participant updates...")
            
            # Method 1: Try COM interface (works with open files)
            if EXCEL_COM_AVAILABLE and com_updates:
                print("Trying COM interface (works with open files)...")
                if update_excel_via_com(tournament_file, com_updates, player_slots_end):
                    print(f"Successfully updated {updated_count} participants via COM")
                    return
                else:
                    print("COM method failed, trying alternative methods...")
            
            # Method 2: Try save with retry and temporary file
            print("Trying pandas save with retry...")
            if save_with_retry(df, tournament_file):
                print(f"Successfully updated {updated_count} participants")
                return
            
            # Method 3: Last resort - inform user with dialog
            error_msg = (f"All save methods failed for {tournament_file}\n\n"
                        "The data has been processed but could not be saved.\n\n"
                        "Please:\n"
                        "1. Close the Excel file\n"
                        "2. Run the script again\n\n"
                        "Or install pywin32 for better file handling:\n"
                        "pip install pywin32")
            
            print(f"\nAll save methods failed for {tournament_file}")
            print("The data has been processed but could not be saved.")
            print("Please:")
            print("1. Close the Excel file")
            print("2. Run the script again")
            print("Or install pywin32 for better file handling: pip install pywin32")
            
            # Show error dialog
            try:
                import win32api
                win32api.MessageBox(0, error_msg, 'Tournament Update Failed', 0x10)  # Error icon
            except:
                pass  # If win32api not available, just print to console
            
        else:
            print("\nNo updates were made to the tournament file")
            
    except Exception as e:
        print(f"Error updating tournament file {tournament_file}: {e}")

def main():
    """Main function to orchestrate the participant update process"""
    print("=== Chess Tournament Participant Update ===")
    
    # Check if tournament folder was passed as command line argument
    if len(sys.argv) > 1:
        tournament_folder = Path(sys.argv[1])
        print(f"Using tournament folder from parameter: {tournament_folder}")
    else:
        # Fallback: Get the script directory and project root
        script_dir = Path(__file__).parent
        project_root = script_dir.parent
        
        # Find the most recent Turnir folder
        current_year = None
        tournament_folder = None
        
        for folder in project_root.glob("Turnir *"):
            if folder.is_dir():
                year_part = folder.name.replace("Turnir ", "")
                if year_part.isdigit():
                    if current_year is None or int(year_part) > int(current_year):
                        current_year = year_part
                        tournament_folder = folder
        
        if current_year is None:
            print("Error: No tournament folder found (looking for 'Turnir YYYY' pattern)")
            return
        
        print(f"Using most recent tournament year: {current_year}")
    
    # Extract year from folder name for file naming
    folder_name = tournament_folder.name
    if "Turnir " in folder_name:
        current_year = folder_name.replace("Turnir ", "")
    else:
        print(f"Warning: Could not extract year from folder name '{folder_name}'")
        current_year = "UNKNOWN"
    
    # Define file paths
    ucesnici_file = tournament_folder / f"Ucesnici {current_year}.xlsx"
    project_root = tournament_folder.parent
    ratings_file = project_root / "Rating" / "all_ratings.csv"
    
    # Check if files exist
    if not ucesnici_file.exists():
        print(f"Error: Participants file not found: {ucesnici_file}")
        return
    
    if not ratings_file.exists():
        print(f"Warning: Ratings file not found: {ratings_file}")
        ratings_lookup = {}
    else:
        ratings_lookup = load_ratings_lookup(ratings_file)
    
    # Find tournament file
    tournament_file = find_tournament_file(tournament_folder)
    if not tournament_file:
        print(f"Error: No tournament file found in {tournament_folder}")
        return
    
    print(f"Using tournament file: {tournament_file}")
    
    # Get paid participants
    paid_participants = get_paid_participants(ucesnici_file)
    if not paid_participants:
        print("No paid participants found. Nothing to update.")
        return
    
    # Update tournament file
    update_tournament_file(tournament_file, paid_participants, ratings_lookup)
    
    print("\n=== Update Complete ===")

if __name__ == "__main__":
    main()
