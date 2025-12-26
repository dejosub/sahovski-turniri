# Chess Tournament Participant Update System

This system automates the transfer of paid participants from the registration file to the tournament bracket file.

## Files Created

1. **`azuriraj_ucesnike.py`** - Main Python script that handles the participant transfer
2. **`vba_button_code.vbs`** - VBA code to add to Excel for the "Azuriraj" button
3. **`README.md`** - This documentation file

## How It Works

1. **Registration**: Participants are listed in `Ucesnici YYYY.xlsx` with payment status
2. **Button Click**: User clicks "Azuriraj" button in Excel
3. **Folder Detection**: VBA macro detects current tournament folder and passes it to Python script
4. **Script Execution**: Python script processes files in the specified tournament folder
5. **Data Transfer**: Script transfers paid participants to tournament file in same folder
6. **Rating Lookup**: Script adds player ratings from `Rating/all_ratings.csv`

## Setup Instructions

### Step 1: Add VBA Button to Excel

1. Open `Turnir 2025/Ucesnici 2025.xlsx`
2. Press `Alt + F11` to open VBA Editor
3. In the Project Explorer, right-click on your workbook and select `Insert > Module`
4. Copy the code from `vba_button_code.vbs` and paste it into the new module
5. Close VBA Editor (`Alt + Q`)
6. Go to `Developer` tab (if not visible, enable it in Excel Options)
7. Click `Insert` and choose `Button (Form Control)`
8. Draw the button on your worksheet
9. In the "Assign Macro" dialog, select `AzurirajUcesnike`
10. Right-click the button and select "Edit Text", change it to "Azuriraj"

### Step 2: Install Dependencies (Optional but Recommended)

For best performance with open Excel files, install the required Python packages:

```bash
pip install -r code/requirements.txt
```

Or install manually:
```bash
pip install pandas openpyxl pywin32
```

### Step 3: Test the System

1. Open `Ucesnici 2025.xlsx` (can stay open now!)
2. Add a payment amount for a test participant in the "Уплаћено учешће" column
3. Click the "Azuriraj" button
4. Check the console output for results
5. Verify the participant was added to `Turnir 2025.xlsm`

## Script Features

### Folder Management
- ✅ **Accepts tournament folder as parameter** from VBA
- ✅ **Automatic fallback** to most recent tournament if no parameter provided
- ✅ **Year-independent** - works with any tournament year structure
- ✅ **Dynamic player slots** - Uses "Непар" marker to detect tournament size

### Participant Processing
- ✅ Reads payment status from "Уплаћено учешће" column
- ✅ Only processes participants with payment > 0
- ✅ Skips participants already in tournament file
- ✅ Finds first available "Играч *" placeholder slot

### Rating Management
- ✅ Looks up ratings from `Rating/all_ratings.csv`
- ✅ Uses default rating of 1400 if player not found
- ✅ Updates "Pocetni poredak" column with rating

### File Access Management
- ✅ **COM Interface** - Updates Excel files even when they're open
- ✅ **Formatting Preservation** - Maintains original cell formatting when updating content
- ✅ **Retry with Backoff** - Multiple save attempts with increasing delays
- ✅ **Temporary File Method** - Creates temp file then replaces original
- ✅ **Automatic Method Selection** - Tries best method first, falls back to alternatives

### Error Handling
- ✅ Checks if files exist before processing
- ✅ Handles file permission errors gracefully
- ✅ Provides clear error messages and instructions
- ✅ Multiple fallback strategies for file access issues

## Usage Notes

1. **File Access**: ✨ **NEW** - Script now works even with open Excel files! (requires pywin32)
2. **Payment Status**: Any positive number in "Уплаћено учешће" column indicates payment
3. **Placeholders**: Script replaces "Играч 1", "Играч 2", etc. with actual names
4. **Duplicates**: Script automatically skips participants already in tournament
5. **Ratings**: Default rating of 1400 is used if player not found in ratings file
6. **Year Independence**: Works with any year - VBA automatically passes correct folder path

## Troubleshooting

### "Permission denied" Error (Rare with new COM interface)
- Install pywin32 for better file handling: `pip install pywin32`
- If still failing, close Excel files and try again
- Make sure no other programs are using the Excel files

### "Python script not found" Error
- Verify the `code` folder is in the correct location
- Check that `azuriraj_ucesnike.py` exists in the `code` folder

### No Participants Found
- Check that payment amounts are entered as numbers (not text)
- Verify participant names are in the "Име" column
- Make sure there are no extra spaces in names

### Rating Not Found
- Check if player name exactly matches name in `Rating/all_ratings.csv`
- Script will use default rating of 1400 if no match found

## File Structure

```
sahovski-turniri/
├── code/
│   ├── azuriraj_ucesnike.py    # Main script
│   ├── vba_button_code.vbs     # VBA code for button
│   └── README.md               # This file
├── Rating/
│   └── all_ratings.csv         # Player ratings lookup
└── Turnir 2025/
    ├── Ucesnici 2025.xlsx      # Registration file
    └── Turnir 2025.xlsm        # Tournament bracket file
```

## Example Output

```
=== Chess Tournament Participant Update ===
Loaded 62 player ratings from Rating/all_ratings.csv
Using tournament file: Turnir 2025/Turnir 2025.xlsm
Found 1 paid participants:
  - Дејан Суботић
Found 13 existing participants in tournament file
  Added Дејан Суботић with rating 1502 (replacing Играч 1)

Successfully updated 1 participants in Turnir 2025.xlsm
=== Update Complete ===
```
