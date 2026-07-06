# msa_alignment_identity_excel.py
# Author: Hagar Ibrahim
#
# Generates:
# 1) Residue-number multiple sequence alignment table
# 2) Percent identity matrix
# Both saved in ONE Excel file with clear PDB ID headers

import sys
import re
from collections import OrderedDict
from itertools import combinations
from openpyxl import Workbook


# --------------------------------------------------
# Input handling
# --------------------------------------------------
if len(sys.argv) > 1:
    alignment_file = sys.argv[1]
else:
    alignment_file = "alignment.txt"

default_start = 1

# --------------------------------------------------
# Read alignment (FASTA + CLUSTAL compatible)
# --------------------------------------------------
sequences = OrderedDict()
seq_id = None

with open(alignment_file, "r") as f:
    for line in f:
        line = line.strip()
        if not line or line.startswith("CLUSTAL") or line.startswith("*"):
            continue

        if line.startswith(">"):
            seq_id = line[1:].split()[0]
            sequences.setdefault(seq_id, [])
        elif re.match(r"^[A-Za-z0-9_\-|]+\s+[A-Z\-]+", line):
            parts = line.split()
            seq_id = parts[0]
            sequences.setdefault(seq_id, [])
            sequences[seq_id].append(parts[1])

# Join fragments
for sid in sequences:
    sequences[sid] = "".join(sequences[sid])

if not sequences:
    raise ValueError("No sequences detected in alignment file")

# --------------------------------------------------
# Normalize lengths
# --------------------------------------------------
max_len = max(len(seq) for seq in sequences.values())
for sid in sequences:
    sequences[sid] = sequences[sid].ljust(max_len, "-")

seq_ids = list(sequences.keys())

# --------------------------------------------------
# Build residue-number MSA table
# --------------------------------------------------
counters = {sid: default_start for sid in seq_ids}
msa_table = []

for i in range(max_len):
    row = []
    for sid in seq_ids:
        aa = sequences[sid][i]
        if aa == "-":
            row.append("-")
        else:
            row.append(f"{counters[sid]}_{aa}")
            counters[sid] += 1
    msa_table.append(row)

# --------------------------------------------------
# Percent identity matrix
# --------------------------------------------------
identity_matrix = [[None] * len(seq_ids) for _ in seq_ids]

for i, s1 in enumerate(seq_ids):
    for j, s2 in enumerate(seq_ids):
        if i == j:
            identity_matrix[i][j] = 100.0
        else:
            matches = 0
            compared = 0
            for a, b in zip(sequences[s1], sequences[s2]):
                if a == "-" and b == "-":
                    continue
                compared += 1
                if a == b:
                    matches += 1
            identity_matrix[i][j] = round((matches / compared) * 100, 2) if compared else 0.0

# --------------------------------------------------
# Write to ONE Excel file (two sheets)
# --------------------------------------------------
wb = Workbook()

# Sheet 1: MSA table
ws_msa = wb.active
ws_msa.title = "MSA_Table"
ws_msa.append(seq_ids)  # PDB IDs as headers

for row in msa_table:
    ws_msa.append(row)

# Sheet 2: Percent identity
ws_pid = wb.create_sheet("Percent_Identity")
ws_pid.append([""] + seq_ids)

for sid, row in zip(seq_ids, identity_matrix):
    ws_pid.append([sid] + row)

# Save file
output_file = "MSA_and_Percent_Identity.xlsx"
wb.save(output_file)

print(f"✅ Excel file created: {output_file}")
