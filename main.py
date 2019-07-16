import pandas as pd
from random import sample, shuffle, choice
from collections import defaultdict

# Get all the speaker names with their organisation
speaker_list = []
speaker_sheet = pd.read_excel("HL_Speakers.xlsx", sheet_name="Master Speakers' List")

for i in speaker_sheet.index:
    if not pd.isna(speaker_sheet['Name(s)'][i]):
        if ('\n' in speaker_sheet['Name(s)'][i]):
            for name in speaker_sheet['Name(s)'][i].split('\n'):
                speaker_det = '{} ({})'.format(name, speaker_sheet['Organisation'][i])
                speaker_list.append(speaker_det)
        else:
            speaker_det = '{} ({})'.format(speaker_sheet['Name(s)'][i], speaker_sheet['Organisation'][i])
            speaker_list.append(speaker_det)

# Generate a placeholder list of participants with their four preferences
participant_list = []
participant_prefs = {}

for i in range(1, 251):
    name = 'Test Participant #{}'.format(i)
    participant_list.append(name)
    participant_prefs[name] = sample(speaker_list, 4)

# Create a speaker dictionary with the possible allocation number
speaker_allocation_dict = {}
ratio = len(participant_prefs) // len(speaker_list)
for i in speaker_list:
    speaker_allocation_dict[i] = (ratio + 1) * 2

# Init final participant allocations
final_allocations = defaultdict(set)

# Try assigning from the four preferences
shuffle(participant_list)
for _ in range(2):
    for participant in participant_list:
        if (speaker_allocation_dict[participant_prefs[participant][0]] > 0) and (participant_prefs[participant][0] not in final_allocations[participant]):
            final_allocations[participant].add(participant_prefs[participant][0])
            speaker_allocation_dict[participant_prefs[participant][0]] -= 1
        elif (speaker_allocation_dict[participant_prefs[participant][1]] > 0) and (participant_prefs[participant][1] not in final_allocations[participant]):
            final_allocations[participant].add(participant_prefs[participant][1])
            speaker_allocation_dict[participant_prefs[participant][1]] -= 1
        elif (speaker_allocation_dict[participant_prefs[participant][2]] > 0) and (participant_prefs[participant][2] not in final_allocations[participant]):
            final_allocations[participant].add(participant_prefs[participant][2])
            speaker_allocation_dict[participant_prefs[participant][2]] -= 1
        elif (speaker_allocation_dict[participant_prefs[participant][3]] > 0) and (participant_prefs[participant][3] not in final_allocations[participant]):
            final_allocations[participant].add(participant_prefs[participant][3])
            speaker_allocation_dict[participant_prefs[participant][3]] -= 1

# If some participants not assigned fully (meaning preferences not worked out), random bomb
shuffle(participant_list)
for _ in range(2):
    for participant in participant_list:
        while len(final_allocations[participant]) < 2:
            speaker = choice(speaker_list)
            if (speaker_allocation_dict[speaker] > 0) and (speaker not in final_allocations[participant]):
                final_allocations[participant].add(speaker)
                speaker_allocation_dict[speaker] -= 1

# Final allocations available!
print(sum(speaker_allocation_dict.values()))