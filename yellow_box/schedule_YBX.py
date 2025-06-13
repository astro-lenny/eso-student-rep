#!/usr/bin/env python3

import random
from datetime import datetime, timedelta
import pandas as pd
import calendar
import argparse
import os

def load_people(filepath):
    people = []
    with open(filepath, 'r') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            parts = line.split(',')
            if len(parts) >= 2:
                last = parts[0].strip()
                first = parts[1].strip()
                people.append((first, last))
    return people

# def create_schedule(people, start_date, filename='schedule.xlsx'):
#     # Align to next Monday if not already a Monday
#     start_date += timedelta(days=(0 - start_date.weekday()) % 7)

#     # Estimate number of 2-week blocks needed (3 people per block)
#     num_blocks = len(people) // 3 + (1 if len(people) % 3 != 0 else 0)

#     # Create 2-week non-overlapping date blocks
#     blocks = []
#     d = start_date
#     for _ in range(num_blocks * 2):  # buffer
#         blocks.append((d, d + timedelta(days=6)))
#         d += timedelta(weeks=1)

#     def week_overlaps(date_range, skip_start, skip_end):
#         start, end = date_range
#         return not (end < skip_start or start > skip_end)

#     # Define skip periods
#     christmas_start = datetime(start_date.year, 12, 24)
#     christmas_end = datetime(start_date.year, 12, 30)
#     following_week_start = christmas_end + timedelta(days=1)
#     following_week_end = following_week_start + timedelta(days=6)

#     # Filter out skip weeks
#     blocks = [b for b in blocks if not (
#         week_overlaps(b, christmas_start, christmas_end) or
#         week_overlaps(b, following_week_start, following_week_end)
#     )]

#     blocks = blocks[:num_blocks]

#     # Shuffle and assign 3 people per block
#     random.shuffle(people)
#     schedule = []
#     for block in blocks:
#         if len(people) >= 3:
#             group = [people.pop() for _ in range(3)]
#         else:
#             group = random.sample(schedule[-1][1], 3) if schedule else random.sample(people, 3)
#         schedule.append((block, group))

#     # Prepare DataFrame
#     df = pd.DataFrame([{
#         'Week': f"{start.strftime('%a, %d %b')} - {end.strftime('%d %b')}",
#         'Person 1': group[0][0] + ' ' + group[0][1],
#         'Person 2': group[1][0] + ' ' + group[1][1],
#         'Person 3': group[2][0] + ' ' + group[2][1],
#     } for (start, end), group in sorted(schedule, key=lambda x: x[0])])

#     # Save Excel file
#     base, ext = os.path.splitext(filename)
#     filename = f"{base}_{start_date.strftime('%Y-%m-%d')}{ext}"
#     if os.path.exists(filename):
#         backup_name = f"{base}_{start_date.strftime('%Y%m%d')}_backup{random.randint(1000,9999)}{ext}"
#         os.rename(filename, backup_name)

#     df.to_excel(filename, index=False, engine='openpyxl')

def create_schedule(people, start_date, year=2025, filename='schedule.xlsx'):
    # Generate weekly slots starting from given start_date
    weeks = []
    d = start_date
    while d.year == year:
        weeks.append((d, d + timedelta(days=6)))
        d += timedelta(weeks=1)

    # Skip Christmas and the following week
    def week_overlaps(date_range, skip_start, skip_end):
        start, end = date_range
        return not (end < skip_start or start > skip_end)

    christmas_start = datetime(year, 12, 24)
    christmas_end = datetime(year, 12, 30)
    skip_weeks = []
    for w in weeks:
        if week_overlaps(w, christmas_start, christmas_end) or \
           week_overlaps(w, christmas_end + timedelta(days=1), christmas_end + timedelta(days=7)):
            skip_weeks.append(w)

    weeks = [w for w in weeks if w not in skip_weeks]

    # Group weeks into pairs (2 weeks per group)
    paired_weeks = [weeks[i:i+2] for i in range(0, len(weeks), 2) if i+1 < len(weeks)]

    # Shuffle and assign 3 people per pair of weeks
    random.shuffle(people)
    schedule = []
    for pair in paired_weeks:
        if len(people) >= 3:
            group = [people.pop() for _ in range(3)]
        else:
            assigned_people = sum((g for _, g in schedule), [])
            group = random.sample(assigned_people, 3)

        for week in pair:
            schedule.append((week, group))

    # Prepare DataFrame with one name per column
    df = pd.DataFrame([{
        'Week': f"{start.strftime('%a, %d %b')} - {end.strftime('%d %b')}",
        'Person 1': f"{group[0][0]} {group[0][1]}",
        'Person 2': f"{group[1][0]} {group[1][1]}",
        'Person 3': f"{group[2][0]} {group[2][1]}",
    } for (start, end), group in sorted(schedule, key=lambda x: x[0])])

    # Adjust filename based on start date
    base, ext = os.path.splitext(filename)
    filename = f"{base}_{start_date.strftime('%Y-%m-%d')}{ext}"

    # Backup old file if it exists
    if os.path.exists(filename):
        backup_name = f"{base}_{start_date.strftime('%Y%m%d')}_backup{random.randint(1000,9999)}{ext}"
        os.rename(filename, backup_name)

    # Save to Excel
    df.to_excel(filename, index=False, engine='openpyxl')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-names', type=str, required=True, help='Path to file with names')
    parser.add_argument('-start', type=str, default=datetime.now().strftime('%Y-%m-%d'), help='Start date (YYYY-MM-DD), defaults to today')
    parser.add_argument('-out', type=str, default='schedule.xlsx', help='Output Excel filename')
    args = parser.parse_args()

    start_date = datetime.strptime(args.start, '%Y-%m-%d')
    people = load_people(args.names)
    create_schedule(people, start_date, filename=args.out)