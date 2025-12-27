#!/usr/bin/env python3
"""
Cleanup script to remove garbage values created by Range.Replace
"""

import win32com.client
from pathlib import Path

def cleanup_tournament_file():
    try:
        # Connect to Excel
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = True
        xl.DisplayAlerts = False
        
        # Find the open workbook
        tournament_file = "Turnir 2025.xlsm"
        workbook = None
        
        for wb in xl.Workbooks:
            if wb.Name == tournament_file:
                workbook = wb
                print(f"Found open workbook: {tournament_file}")
                break
        
        if not workbook:
            print("Tournament file not found open in Excel")
            return
        
        worksheet = workbook.Worksheets(1)
        
        # Clean up garbage values - replace unwanted "1381" and "###" with empty
        print("Cleaning up garbage values...")
        
        # Find the range to clean (avoid the actual participant data)
        # Clean from row 11 onwards (after the actual participants)
        cleanup_range = worksheet.Range("A11:Z100")  # Adjust range as needed
        
        # Remove "1381" values that shouldn't be there
        cleanup_range.Replace(
            What="1381",
            Replacement="",
            LookAt=2,  # xlPart - partial match
            SearchOrder=1,
            MatchCase=False
        )
        
        # Remove "###" symbols
        cleanup_range.Replace(
            What="###",
            Replacement="",
            LookAt=2,  # xlPart - partial match
            SearchOrder=1,
            MatchCase=False
        )
        
        # Save the workbook
        workbook.Save()
        print("Cleanup completed and file saved")
        
    except Exception as e:
        print(f"Cleanup failed: {e}")
    finally:
        try:
            xl.DisplayAlerts = True
        except:
            pass

if __name__ == "__main__":
    cleanup_tournament_file()
