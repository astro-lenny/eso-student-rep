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



# def create_schedule(people, start_date, year=2025, filename='schedule.xlsx'):
#     # Generate weekly slots starting from given start_date
#     weeks = []
#     d = start_date
#     while d.year == year:
#         weeks.append((d, d + timedelta(days=6)))
#         d += timedelta(weeks=1)

#     # Skip Christmas and the following week
#     def week_overlaps(date_range, skip_start, skip_end):
#         start, end = date_range
#         return not (end < skip_start or start > skip_end)

#     christmas_start = datetime(year, 12, 24)
#     christmas_end = datetime(year, 12, 30)
#     skip_weeks = []
#     for w in weeks:
#         if week_overlaps(w, christmas_start, christmas_end) or \
#            week_overlaps(w, christmas_end + timedelta(days=1), christmas_end + timedelta(days=7)):
#             skip_weeks.append(w)

#     weeks = [w for w in weeks if w not in skip_weeks]

#     # Group weeks into pairs (2 weeks per group)
#     paired_weeks = [weeks[i:i+2] for i in range(0, len(weeks), 2) if i+1 < len(weeks)]

#     # Shuffle and assign 3 people per pair of weeks
#     random.shuffle(people)
#     schedule = []
#     for pair in paired_weeks:
#         if len(people) >= 3:
#             group = [people.pop() for _ in range(3)]
#         else:
#             assigned_people = sum((g for _, g in schedule), [])
#             group = random.sample(assigned_people, 3)

#         for week in pair:
#             schedule.append((week, group))

#     # Prepare DataFrame with one name per column
#     df = pd.DataFrame([{
#         'Week': f"{start.strftime('%a, %d %b')} - {end.strftime('%d %b')}",
#         'Person 1': f"{group[0][0]} {group[0][1]}",
#         'Person 2': f"{group[1][0]} {group[1][1]}",
#         'Person 3': f"{group[2][0]} {group[2][1]}",
#     } for (start, end), group in sorted(schedule, key=lambda x: x[0])])

#     # Adjust filename based on start date
#     base, ext = os.path.splitext(filename)
#     filename = f"{base}_{start_date.strftime('%Y-%m-%d')}{ext}"

#     # Backup old file if it exists
#     if os.path.exists(filename):
#         backup_name = f"{base}_{start_date.strftime('%Y%m%d')}_backup{random.randint(1000,9999)}{ext}"
#         os.rename(filename, backup_name)

#     # Save to Excel
#     df.to_excel(filename, index=False, engine='openpyxl')

def create_schedule(people, start_date, filename='schedule.xlsx'):
    # --- 1. build calendar weeks for the year starting at start_date
    weeks = []
    d = start_date
    num_blocks = len(people)  # 1 block per person
    weeks = []
    d = start_date
    while len(weeks) < num_blocks * 2:   # 2 weeks per block
        weeks.append((d, d + timedelta(days=6)))
        d += timedelta(weeks=1)

    # 1a. remove the week before christmas + two weeks after
    years = sorted({w[0].year for w in weeks} | {w[1].year for w in weeks})
    skip_ranges = []
    for year in years:
        christmas = datetime(year, 12, 24)
        skip_ranges.append((christmas - timedelta(days=7), christmas - timedelta(days=1)))   # week before
        skip_ranges.append((christmas, christmas + timedelta(days=6)))                      # week of
        skip_ranges.append((christmas + timedelta(days=7), christmas + timedelta(days=13)))  # week after
    def skip_week(week):
        s, e = week
        return any(not (e < a or s > b) for a, b in skip_ranges)
    weeks = [w for w in weeks if not skip_week(w)]

    # --- 2. randomize people
    random.shuffle(people)
    print(people)
    # --- 3. assign 1 person per 2-week block
    schedule = []
    for i, week in enumerate(weeks):
        person = people[(i // 2) % len(people)]   # same person for 2 consecutive weeks
        schedule.append({
            'Year': week[0].year,  # year of the week start
            'Week': f"{week[0].strftime('%a, %d %b')} - {week[1].strftime('%d %b')}",
            'Person': f"{person[0]} {person[1]}" if isinstance(person, (list, tuple)) else str(person)
        })

    df = pd.DataFrame(schedule)
    print(df)

    # save to excel
    base, ext = os.path.splitext(filename)
    filename = f"{base}_{start_date.strftime('%Y-%m-%d')}{ext}"
    if os.path.exists(filename):
        backup_name = f"{base}_{start_date.strftime('%Y%m%d')}_backup{random.randint(1000,9999)}{ext}"
        os.rename(filename, backup_name)

    df.to_excel(filename, index=False, engine='openpyxl')
    return df

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-names', type=str, required=True, help='Path to file with names')
    parser.add_argument('-start', type=str, default=datetime.now().strftime('%Y-%m-%d'), help='Start date (YYYY-MM-DD), defaults to today')
    parser.add_argument('-out', type=str, default='schedule.xlsx', help='Output Excel filename')
    args = parser.parse_args()

    start_date = datetime.strptime(args.start, '%Y-%m-%d')
    people = load_people(args.names)
    create_schedule(people, start_date, filename=args.out)