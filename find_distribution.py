import pandas as pd
from collections import defaultdict

# Get all the speaker names with their organisation
speaker_list = []
speaker_sheet = pd.read_excel("HL_Speakers.xlsx", sheet_name="Afternoon")

for i in speaker_sheet.index:
    speaker_det = '{} ({})'.format(speaker_sheet['Name'][i], speaker_sheet['Organisation'][i])
    speaker_list.append(speaker_det)

# Create a speaker dictionary with the possible allocation number
speaker_allocation_dict = {}
ratio = 170 // len(speaker_list)
for i in speaker_list:
    speaker_allocation_dict[i] = [ratio + 1, ratio + 1]

# Reduce the counts of speakers depending on how much they have been chosen
speaker_pref = pd.read_excel("Allocations_afternoon.xlsx", sheet_name="Sheet1")
for i in speaker_pref.index:
    speaker_allocation_dict[speaker_pref[1][i]][0] -= 1
    speaker_allocation_dict[speaker_pref[2][i]][1] -= 1

for k, v in speaker_allocation_dict.items():
    speaker_allocation_dict[k][0] = ratio + 1 - (speaker_allocation_dict[k][0])
    speaker_allocation_dict[k][1] = ratio + 1 - (speaker_allocation_dict[k][1])

final_number = pd.DataFrame.from_dict(speaker_allocation_dict, orient='index')
final_number.to_excel('Speaker_Quotas_Afternoon.xlsx')