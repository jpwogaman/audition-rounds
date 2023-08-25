"""Randomly selects excerpts for mock audition rounds, prints 1 page per round to default printer"""

import random
import tempfile
from typing import List
import os
import win32api
import win32print

def main():
    """entry point"""
    rounds = int(input("How many rounds? "))
    determine_rounds(rounds)

def dev_check():
    """print console if dev/true, print printer if prod/false"""
    return False

def number_of_excerpts(round_number):
    """determines the number of excerpts in a given round, with more excerpts in later rounds"""

    if round_number == 0:
        return random.randint(4, 6)
    if 0 < round_number <= 3:
        return random.randint(5, 9)
    return random.randint(6, 12)

def determine_rounds(rounds):
    """randomly determines the order of excerpts for each round, then prints them"""

    for i in range(rounds):

        number_of_excerpts_this_round = number_of_excerpts(i)
        selected_excerpts: List[str] = [f"Round {i+1}:\n"]
        if dev_check():
            print(f"Round {i+1}:\n")

        list_of_excerpts = [
            "Respighi, Pines of Rome",
            "Shostakovich, Festive Overture, Pt. 2",
            "Williams, Summon the Heroes, Solo",
            "Copland, Outdoor Overture",
            "Charlier, Etude No. 6",
            "Hindemith, Symphony Mvt. 2",
            "Elliot, British Eighth March",
            "Bernstein, West Side Story",
            "Stravinsky, Firebird, Pt. 1",
            "Stravinsky, Firebird, Pt. 2",
            "Barber, Commando March",
            "Hindemith, Symphony Mvt. 1",
            "Williams, Summon the Heroes, Pt. 1",
            "Williams, Summon the Heroes, Pt. 2",
            "Shostakovich, Festive Overture, Pt. 1",
            "Grafulla, Washington Grays March",
            "Smith, Festival Variations",
            "Ives, Variations on America",
            "Tchaikovsky, Symphony No. 4",
            "Butterfield, Taps",
            "Mussorgsky, Pictures at an Exhibition",
            "Williams, Hymn to the Fallen, Pt. 2",
            "Williams, Hymn to the Fallen, Pt. 1"
        ]

        for j in range(number_of_excerpts_this_round):
            excerpt = random.choice(list_of_excerpts)
            selected_excerpts.append(f"{j+1}. {excerpt}")
            list_of_excerpts.remove(excerpt)
            if dev_check():
                print(f"{j+1}. {excerpt}")

        selected_excerpts.append("\n")

        if dev_check():
            print("\n")
        else:
            txt_file = create_file("\n".join(selected_excerpts))
            print_action(txt_file)

def create_file(audition_rounds):
    """creates a .txt file with the audition round information"""

    filename = tempfile.mktemp(  # NOSONAR @SuppressWarnings("python:S5445")
        ".txt")
    with open(filename, "w", encoding='utf-8') as file:
        file.write(audition_rounds)
    return filename

def print_action(txt_file):
    """actually prints the file"""

    win32api.ShellExecute(  # pylint: disable=c-extension-no-member
        0,
        "print",
        txt_file,
        f'/d:"{win32print.GetDefaultPrinter()}"',  # pylint: disable=c-extension-no-member
        ".",
        0
    )

    os.remove(txt_file)

if __name__ == "__main__":
    main()
