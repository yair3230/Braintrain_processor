import os
import json

filename = '.\\expected_results\\Psycho_2025_a.json'
folder = filename + '_chunks'
if not os.path.isdir(folder):
    os.mkdir(folder)
with open(filename) as openfile:
    content = json.load(openfile)
chunks = [[]]
for q in content:
    if len(chunks[-1]) >= 5:
        chunks.append([])
    chunks[-1].append(q)
for index, chunk in enumerate(chunks):

    with open(folder + f'\\{index}.json', 'w') as openfile:
        json.dump(chunk, openfile, indent=4)
