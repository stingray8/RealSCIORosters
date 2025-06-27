import pulp
import time
import copy
import pandas as pd
import openpyxl
import numpy as np
import copy
import math
from tabulate import tabulate
from functions import *
from main import *
file_path = 'Compute Rosters.xlsx'


team_df = pd.read_excel(file_path, sheet_name="Team")

team_info = data_frame_to_np(team_df)
schedule = team_info[1][7]


def remove_empty(array):
    z = 0
    while z < len(array):
        if array[z] == []:
            array.pop(z)
            z -= 1


time_conflicts = pd.read_excel(file_path, sheet_name=schedule)
time_conflicts = time_conflicts.replace('nan', np.nan)
time_conflicts = data_frame_to_np(time_conflicts)
time_conflicts = [[item for item in row if not pd.isna(item)] for row in time_conflicts]
time_conflicts = [[item.title() for item in row if not "Unnamed" in item] for row in time_conflicts]
time_conflicts = [tuple(_) for _ in time_conflicts]
print(time_conflicts)


test_roster = data_frame_to_np(pd.read_excel(file_path, sheet_name="Test Roster")).tolist()
test_roster = [[item for item in row if not pd.isna(item)] for row in test_roster]
test_roster = [[item.title() for item in row if not "Unnamed" in item] for row in test_roster]

if len(test_roster) != 15:
    raise Exception("Not enough people. Make sure to use person to event roster, not event to person")


def check_person(row):
    max_count = 0
    for conflict in time_conflicts:
        count = 0
        for event in conflict:
            if event in row:
                count += 1
        max_count = max(max_count, count)
    return max_count


print("--------")
mistake = False
for person in test_roster:
    if not check_person(person) <= 1:
        print_red("mistake " + str(person))
        mistake = True
if mistake:
    print_red("Mistakes found")
else:
    print("No scheduling mistakes found")


events_with_three_people_df = pd.read_excel(file_path, sheet_name="3 Person Events")
events_with_three_people = data_frame_to_np(events_with_three_people_df).tolist()
events_with_three_people = tuple(event[0] for event in events_with_three_people)

people_event = {}
event_people = {}
for row in test_roster:
    person_events = []
    for i in range(1, len(row)):
        if pd.isna(row[i]) or "Unnamed" in row[i]:
            continue
        person_events.append(row[i])
        if row[i] not in event_people:
            event_people[row[i]] = [row[0]]
        else:
            event_people[row[i]].append(row[0])

    people_event[row[0]] = person_events

most_events = 4
for key in people_event:
    if len(people_event[key]) > most_events:
        print_red(key + " has " + str(len(people_event[key])) + " events")

for key in event_people:
    print(event_people[key])
    if len(event_people[key]) < 2:
        print_red(key + " doesn't have enough people")
    elif len(event_people[key]) == 3 and key not in events_with_three_people:
        print_red(key + " has " + str(len(event_people[key])) + " people")
    elif len(event_people[key]) != 2:
        if not (len(event_people[key]) == 3 and key in events_with_three_people):
            print_red(key + " has " + str(len(event_people[key])) + " people")

with open("output.txt", "w"):
    pass

print('people event', people_event)
result = list(people_event.items())
get_results(result, event_assignments=event_people)

