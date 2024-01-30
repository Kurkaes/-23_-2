import json


import csv

INPUT_FILENAME = "input.csv"
OUTPUT_FILENAME = "output.json"


def task() -> None:
    with open(INPUT_FILENAME, 'r') as bimbim:
        reader = csv.DictReader(bimbim)
        with open(OUTPUT_FILENAME, 'w') as bambam:
            json.dump(list(reader), bambam, indent=4)


if __name__ == '__main__':

    task()
    with open(OUTPUT_FILENAME) as output_f:
        for line in output_f:
            print(line, end="")
