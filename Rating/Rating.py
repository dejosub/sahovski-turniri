import csv
import os
import shutil
from pathlib import Path
import pandas as pd
#pip install matplotlib
import matplotlib.pyplot as plt

k_factor = 40
initial_rating = 1400


def new_elo_rating(player_rating, opponent_rating, result):
    """
    Calculate the new Elo rating for a player.
    """
    expected_score = 1 / (1 + 10 ** ((opponent_rating - player_rating) / 400))
    new_rating = player_rating + k_factor * (result - expected_score)
    return round(new_rating)

def ensure_player_initial_rating(player, per_player_history):
    if player not in per_player_history:
        per_player_history[player] = [initial_rating]

def update_ratings(match_results, per_player_history):
    """
    Update Elo ratings for a set of players based on match results.
    """
    ratings = {}

    for player1, player2, result in match_results:
        ensure_player_initial_rating(player1, per_player_history)
        ensure_player_initial_rating(player2, per_player_history)

        player1_rating = ratings.get(player1, initial_rating)
        player2_rating = ratings.get(player2, initial_rating)

        ratings[player1] = new_elo_rating(player1_rating, player2_rating, result)
        per_player_history[player1].append(ratings[player1])
        
        # sum of all ratings for all players is constant
        # increase for the first player is equal to the decrease for the second player
        ratings[player2] = player2_rating - (ratings[player1] - player1_rating)
        per_player_history[player2].append(ratings[player2])

    return ratings

# Read match results from a series of CSV files
match_results = []
file_number = 1

while True:
    filename = f"..\\rezultati\\{file_number}.csv"
    if not os.path.exists(filename):
        break

    with open(filename, 'r', encoding='utf-8') as file:
        print(f"Processing {filename} ...")
        reader = csv.reader(file)
        for row in reader:
            # Skip the first column and read the rest
            player1, player2, result = row[1], row[2], float(row[3])
            match_results.append((player1, player2, result))

    file_number += 1

print("\n\n")

# Calculate and print the final ratings
per_player_history = {}
final_ratings = update_ratings(match_results, per_player_history)

# Generate history files for all players (keep existing behavior)
i = 1
for player, history in per_player_history.items():
    print(f"Processing history for player {i}", end="\r")

    with open(f"history\\{player}.csv", 'w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        for rating in history:
            writer.writerow([int(rating)])
    plt.figure(figsize=(10, 6))  # Width and height in inches
    plt.plot(history, marker='o')  # 'o' adds circle markers to each point
    plt.title(f'{player}')
    plt.ylabel('Rating')

    # Save the plot as a JPEG file
    plt.savefig(f'history\\{player}.jpg', format='jpg', dpi = 300)
    plt.close()
    i += 1

print("\n\n")

# Sort the ratings by their value in descending order
sorted_ratings = sorted(final_ratings.items(), key=lambda x: x[1], reverse=True)

def process_and_write_ratings(filename, sorted_ratings, players_to_hide_output=None, print_to_console=False):
    result = []
    for player, rating in sorted_ratings:
        if players_to_hide_output is None or player not in players_to_hide_output:
            result.append((player, int(rating)))
    
    if print_to_console:
        for i, (player, rating) in enumerate(result, start=1):
            print(f"{i}. {player}: {rating}")

    with open(filename, 'w', newline='', encoding='utf-8-sig') as file:
        writer = csv.writer(file)
        for i, (player, rating) in enumerate(result, start=1):
            writer.writerow([player, rating])

# Generate all_ratings.csv for all players (keep existing behavior)
process_and_write_ratings('all_ratings.csv', sorted_ratings, players_to_hide_output=None, print_to_console=False)

# Find the most recent tournament directory and generate tournament-specific files
def find_most_recent_tournament():
    """Find the most recent Turnir* directory"""
    base_path = Path('..')
    tournament_dirs = list(base_path.glob('Turnir*'))
    if not tournament_dirs:
        print("No tournament directories found")
        return None
    
    # Sort by directory name (assumes chronological naming)
    most_recent = sorted(tournament_dirs, key=lambda x: x.name)[-1]
    return most_recent

def get_tournament_participants(tournament_dir):
    """Get list of participants from tournament Excel file"""
    # Look for tournament Excel file
    tournament_files = list(tournament_dir.glob('Turnir*.xlsm')) + list(tournament_dir.glob('Turnir*.xlsx'))
    if not tournament_files:
        print(f"No tournament Excel file found in {tournament_dir}")
        return set()
    
    tournament_file = tournament_files[0]
    try:
        # Read the tournament file to get participant names
        df = pd.read_excel(tournament_file)
        # Get the first column which should contain player names
        player_names = df.iloc[:, 0].dropna().astype(str)
        # Filter out placeholder entries and empty cells
        participants = set()
        for name in player_names:
            name = name.strip()
            if name and not name.startswith('Играч') and name != 'Укупно':
                participants.add(name)
        return participants
    except Exception as e:
        print(f"Error reading tournament file {tournament_file}: {e}")
        return set()

# Find most recent tournament and generate tournament-specific files
tournament_dir = find_most_recent_tournament()
if tournament_dir:
    print(f"\nProcessing tournament: {tournament_dir.name}")
    
    # Get tournament participants
    participants = get_tournament_participants(tournament_dir)
    print(f"Found {len(participants)} participants in tournament")
    
    if participants:
        # Filter ratings for tournament participants only
        tournament_ratings = [(player, rating) for player, rating in sorted_ratings if player in participants]
        
        # Generate rating.csv in tournament directory
        rating_file = tournament_dir / 'rating.csv'
        process_and_write_ratings(str(rating_file), tournament_ratings, players_to_hide_output=None, print_to_console=True)
        print(f"\nGenerated {rating_file}")
        
        # Create rating_history folder in tournament directory
        history_dir = tournament_dir / 'rating_history'
        history_dir.mkdir(exist_ok=True)
        
        # Copy history files for tournament participants
        copied_count = 0
        for participant in participants:
            # Copy CSV history file
            src_csv = Path(f'history/{participant}.csv')
            if src_csv.exists():
                dst_csv = history_dir / f'{participant}.csv'
                shutil.copy2(src_csv, dst_csv)
                copied_count += 1
            
            # Copy JPG history file
            src_jpg = Path(f'history/{participant}.jpg')
            if src_jpg.exists():
                dst_jpg = history_dir / f'{participant}.jpg'
                shutil.copy2(src_jpg, dst_jpg)
        
        print(f"Copied {copied_count} history files to {history_dir}")
    else:
        print("No participants found in tournament file")
else:
    print("No tournament directory found")

print("\n\n")
