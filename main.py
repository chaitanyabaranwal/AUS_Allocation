import pandas as pd
from random import sample, shuffle, choice
from collections import defaultdict

# Get all the speaker names with their organisation
speaker_list = []
speaker_sheet = pd.read_excel("HL_Speakers.xlsx", sheet_name="Afternoon")

for i in speaker_sheet.index:
    if not pd.isna(speaker_sheet['Name'][i]):
        if ('\n' in speaker_sheet['Name'][i]):
            for name in speaker_sheet['Name'][i].split('\n'):
                speaker_det = '{} ({})'.format(name, speaker_sheet['Organisation'][i])
                speaker_list.append(speaker_det)
        else:
            speaker_det = '{} ({})'.format(speaker_sheet['Name'][i], speaker_sheet['Organisation'][i])
            speaker_list.append(speaker_det)

# Generate a list of participants with their four preferences
participant_group_map = {}
participant_prefs = {}

participant_sheet = pd.read_excel("Final_Preferences.xlsx", sheet_name="Afternoon")
for i in participant_sheet.index:
    name = participant_sheet['Name'][i]
    group = participant_sheet['Group Number'][i]
    participant_group_map[name] = group
    participant_prefs[name] = [participant_sheet['First'][i], participant_sheet['Second'][i], participant_sheet['Third'][i], participant_sheet['Fourth'][i]]

participant_list = list(participant_group_map.keys())

# Create a speaker dictionary with the possible allocation number
speaker_allocation_dict = {}
ratio = 170 // len(speaker_list)
for i in speaker_list:
    speaker_allocation_dict[i] = (ratio + 1) * 2

# Init final participant allocations
final_allocations = defaultdict(list)

# Try assigning from the four preferences
shuffle(participant_list)
for _ in range(2):
    for participant in participant_list:
        if (speaker_allocation_dict[participant_prefs[participant][0]] > 0) and (participant_prefs[participant][0] not in final_allocations[participant]):
            final_allocations[participant].append(participant_prefs[participant][0])
            speaker_allocation_dict[participant_prefs[participant][0]] -= 1
        elif (speaker_allocation_dict[participant_prefs[participant][1]] > 0) and (participant_prefs[participant][1] not in final_allocations[participant]):
            final_allocations[participant].append(participant_prefs[participant][1])
            speaker_allocation_dict[participant_prefs[participant][1]] -= 1
        elif (speaker_allocation_dict[participant_prefs[participant][2]] > 0) and (participant_prefs[participant][2] not in final_allocations[participant]):
            final_allocations[participant].append(participant_prefs[participant][2])
            speaker_allocation_dict[participant_prefs[participant][2]] -= 1
        elif (speaker_allocation_dict[participant_prefs[participant][3]] > 0) and (participant_prefs[participant][3] not in final_allocations[participant]):
            final_allocations[participant].append(participant_prefs[participant][3])
            speaker_allocation_dict[participant_prefs[participant][3]] -= 1

# If some participants not assigned fully (meaning preferences not worked out), random bomb
shuffle(participant_list)
for _ in range(2):
    for participant in participant_list:
        while len(final_allocations[participant]) < 2:
            speaker = choice(speaker_list)
            if (speaker_allocation_dict[speaker] > 0) and (speaker not in final_allocations[participant]):
                final_allocations[participant].append(speaker)
                speaker_allocation_dict[speaker] -= 1

# Convert to list to add stuff
for participant in participant_list:
    final_allocations[participant].insert(0, participant_group_map[participant])

# Final allocations available!
final_allocations = pd.DataFrame.from_dict(final_allocations, orient='index')
final_allocations.to_excel('Allocations_afternoon.xlsx')
