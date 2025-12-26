import csv
import os
#pip install matplotlib
import matplotlib.pyplot as plt

k_factor = 40
initial_rating = 1400

players_to_hide_output = {
    'Александар Јакшић',
    'Стеван Радаковић',
    'Предраг Тошић',
    'Емир Баручија',
    'Слободан Жарић',
    'Сергеј Хринко',
    'Звонко Николић',
    'Ануш',
    'Милош Грујић (*)',
    'Перо Остојић',
    'Мирко Спасојевић',
    'Марко Лакић',
    'Tim Hendon',
    'Саша Радојевић',
    'Аца Спасојевић',
    'Слободан Бојанић',
    'Саша Дожић',
    'Алекса Миловановић',
    'Перица Милошевић',
    'Далибор Марчета',
    'Nathan Main',
    'Eric Main',
    'Sagnik Sinha',
    'Ивица Колев',
    'Eli Main',
    'Кита Колева',
    'Cody Rose'
}


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

i = 1
for player, history in per_player_history.items():
    if player not in players_to_hide_output:
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

process_and_write_ratings('rating.csv', sorted_ratings, players_to_hide_output=players_to_hide_output, print_to_console=True)
process_and_write_ratings('all_ratings.csv', sorted_ratings, players_to_hide_output=None, print_to_console=False)

print("\n\n")
