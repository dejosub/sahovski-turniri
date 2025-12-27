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

def create_tournament_file_from_template(tournament_folder):
    """Create tournament file from template if it doesn't exist"""
    try:
        tournament_folder = Path(tournament_folder)
        folder_name = tournament_folder.name
        
        # Extract year/id from folder name (remove "Turnir " prefix)
        if folder_name.startswith("Turnir "):
            tournament_id = folder_name[7:]  # Remove "Turnir " prefix
        else:
            tournament_id = folder_name
        
        # Define paths
        project_root = tournament_folder.parent
        template_file = project_root / "Sabloni" / "Sahovski turnir sablon.xlsm"
        new_tournament_file = tournament_folder / f"Turnir {tournament_id}.xlsm"
        
        print(f"Template file: {template_file}")
        print(f"New tournament file: {new_tournament_file}")
        
        # Check if template exists
        if not template_file.exists():
            print(f"Error: Template file not found: {template_file}")
            return None
        
        # Simple file copy - no processing, just copy and rename
        shutil.copy2(template_file, new_tournament_file)
        print(f"Created tournament file from template: {new_tournament_file}")
        
        return str(new_tournament_file)
        
    except Exception as e:
        print(f"Error creating tournament file from template: {e}")
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
                break
        
        # If not open, open it
        if workbook is None:
            workbook = xl.Workbooks.Open(tournament_file)
        
        # Get the first worksheet
        worksheet = workbook.Worksheets(1)
        
        # Apply updates
        updates_applied = 0
        for row_idx, col_name, value in updates:
            # Find column index by name (check first row for headers)
            col_idx = None
            for col in range(1, worksheet.UsedRange.Columns.Count + 1):
                header_value = worksheet.Cells(1, col).Value
                if header_value and str(header_value).strip() == str(col_name).strip():
                    col_idx = col
                    break
            
            # If not found in first row, try to find by column position for player names
            if col_idx is None and col_name == "Unnamed: 0":
                col_idx = 1  # First column is player names
            
            if col_idx:
                # Excel uses 1-based indexing, pandas uses 0-based
                excel_row = row_idx + 2  # +1 for 0-based to 1-based, +1 for header row
                
                # Get current cell and its value
                current_cell = worksheet.Cells(excel_row, col_idx)
                
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
                        except:
                            pass  # Ignore formatting errors
                    
                except:
                    current_cell.Value = value
                
                updates_applied += 1
        
        if updates_applied > 0:
            # Save the workbook
            workbook.Save()
            return True
        else:
            return False
        
    except Exception as e:
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
            return True
            
        except PermissionError as e:
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
                retry_delay *= 2  # Exponential backoff
            
        except Exception as e:
            break
    
    # Method 2: Try saving to temporary file then replace
    try:
        temp_file = str(Path(file_path).with_suffix('.tmp.xlsx'))
        df.to_excel(temp_file, index=False)
        
        # Wait a moment then replace
        time.sleep(0.5)
        shutil.move(temp_file, file_path)
        return True
        
    except Exception as e:
        pass
    
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
        
        # Process each paid participant and collect updates
        updated_count = 0
        com_updates = []  # For COM method: (row_idx, col_name, value)
        failed_participants = []  # Track participants that couldn't be added
        
        for participant in paid_participants:
            if participant in existing_participants:
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
                    if 'Relativna snaga' in df.columns:
                        df.at[idx, 'Relativna snaga'] = rating
                    
                    # Store updates for COM method
                    com_updates.append((idx, player_name_column, participant))
                    if 'Relativna snaga' in df.columns:
                        com_updates.append((idx, 'Relativna snaga', rating))
                    
                    print(f"{participant} - {rating}")
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
            # Method 1: Try COM interface (works with open files)
            if EXCEL_COM_AVAILABLE and com_updates:
                if update_excel_via_com(tournament_file, com_updates, player_slots_end):
                    return
            
            # Method 2: Try save with retry and temporary file
            if save_with_retry(df, tournament_file):
                return
            
            # Method 3: Last resort - inform user
            print("Error: Could not save tournament file. Please close Excel and try again.")
        else:
            print("No new participants to transfer")
            
    except Exception as e:
        print(f"Error updating tournament file {tournament_file}: {e}")

def process_forbidden_pairs(tournament_file, paid_participants, project_root):
    """Process forbidden pairs and add them to tournament file"""
    try:
        # Find forbidden pairs Excel file
        forbidden_file = None
        for ext in ['.xlsx', '.xlsm']:
            test_file = project_root / f"Забрањени парови{ext}"
            if test_file.exists():
                forbidden_file = test_file
                break
        
        if not forbidden_file:
            print("No forbidden pairs Excel file found (looking for Забрањени парови.xlsx or .xlsm)")
            return
        
        # Read forbidden pairs from Excel
        forbidden_groups = []
        df = pd.read_excel(forbidden_file, header=None)  # No headers, just data
        
        for _, row in df.iterrows():
            # Get all non-null values from the row and trim whitespace
            names = []
            for value in row:
                if pd.notna(value) and str(value).strip():
                    names.append(str(value).strip())
            
            if len(names) >= 2:
                forbidden_groups.append(names)
        
        if not forbidden_groups:
            print("No forbidden pairs found in file")
            return
        
        # Generate pairs from paid participants
        forbidden_pairs = []
        paid_set = set(paid_participants)
        
        for group in forbidden_groups:
            # Find which names from this group are paid participants
            paid_in_group = [name for name in group if name in paid_set]
            
            # Generate all pairs from paid participants in this group
            for i in range(len(paid_in_group)):
                for j in range(i + 1, len(paid_in_group)):
                    pair = (paid_in_group[i], paid_in_group[j])
                    forbidden_pairs.append(pair)
        
        if not forbidden_pairs:
            print("No forbidden pairs among paid participants")
            return
        
        print(f"Found {len(forbidden_pairs)} forbidden pairs among paid participants:")
        for player1, player2 in forbidden_pairs:
            print(f"  {player1} - {player2}")
        
        # Add pairs to tournament file
        add_forbidden_pairs_to_tournament(tournament_file, forbidden_pairs)
        
    except Exception as e:
        print(f"Error processing forbidden pairs: {e}")

def add_forbidden_pairs_to_tournament(tournament_file, forbidden_pairs):
    """Add forbidden pairs to the tournament file's Zabranjeni parovi sheet"""
    try:
        if EXCEL_COM_AVAILABLE:
            # Use COM to add to specific sheet
            xl = win32com.client.Dispatch("Excel.Application")
            xl.Visible = True
            xl.DisplayAlerts = False
            
            # Try to find if workbook is open
            workbook = None
            file_name = Path(tournament_file).name
            
            for wb in xl.Workbooks:
                if wb.Name == file_name:
                    workbook = wb
                    break
            
            if workbook is None:
                workbook = xl.Workbooks.Open(tournament_file)
            
            # Find or create "Zabranjeni parovi" sheet
            sheet = None
            for ws in workbook.Worksheets:
                if ws.Name == "Zabranjeni parovi":
                    sheet = ws
                    break
            
            if sheet is None:
                print("Warning: 'Zabranjeni parovi' sheet not found in tournament file")
                return
            
            # Clear existing content (optional - you might want to append instead)
            sheet.UsedRange.Clear()
            
            # Add pairs
            for i, (player1, player2) in enumerate(forbidden_pairs, start=1):
                sheet.Cells(i, 1).Value = player1
                sheet.Cells(i, 2).Value = player2
            
            workbook.Save()
            print(f"Added {len(forbidden_pairs)} forbidden pairs to 'Zabranjeni parovi' sheet")
            
        else:
            print("Warning: COM not available, cannot update forbidden pairs sheet")
            
    except Exception as e:
        print(f"Error adding forbidden pairs to tournament file: {e}")

def main():
    """Main function to orchestrate the participant update process"""
    
    # Determine tournament folder
    if len(sys.argv) > 1:
        tournament_folder = Path(sys.argv[1])
    else:
        # Find the most recent tournament folder
        tournament_folder = find_most_recent_tournament()
        if not tournament_folder:
            print("Error: No tournament folder found")
            return
    
    # Check if tournament folder exists
    if not tournament_folder.exists():
        print(f"Error: Tournament folder does not exist: {tournament_folder}")
        return
    
    # Find participants file (only .xlsm files with macros)
    year = tournament_folder.name.split()[-1]
    ucesnici_file = tournament_folder / f"Ucesnici {year}.xlsm"
    
    if not ucesnici_file.exists():
        # Check if .xlsx exists and give specific error
        xlsx_file = tournament_folder / f"Ucesnici {year}.xlsx"
        if xlsx_file.exists():
            print(f"Error: Found {xlsx_file.name} but this script requires .xlsm file with macros.")
            print(f"Please rename {xlsx_file.name} to {ucesnici_file.name} and add the VBA macro.")
            return
        else:
            print(f"Error: Participants file not found: {ucesnici_file.name}")
            return
    
    # Load ratings lookup (relative to project root)
    project_root = tournament_folder.parent
    ratings_file = project_root / "Rating" / "all_ratings.csv"
    
    if not ratings_file.exists():
        print(f"Warning: Ratings file not found: {ratings_file}")
        ratings_lookup = {}
    else:
        ratings_lookup = load_ratings_lookup(ratings_file)
    
    # Find tournament file
    tournament_file = find_tournament_file(tournament_folder)
    if not tournament_file:
        print("Creating tournament file from template...")
        tournament_file = create_tournament_file_from_template(tournament_folder)
        if not tournament_file:
            print("Failed to create tournament file from template")
            return
    
    # Get paid participants
    paid_participants = get_paid_participants(ucesnici_file)
    if not paid_participants:
        print("No paid participants found.")
        return
    
    print(f"Found {len(paid_participants)} paid participants")
    
    # Update tournament file
    update_tournament_file(tournament_file, paid_participants, ratings_lookup)
    
    # Process forbidden pairs
    process_forbidden_pairs(tournament_file, paid_participants, project_root)

if __name__ == "__main__":
    main()
