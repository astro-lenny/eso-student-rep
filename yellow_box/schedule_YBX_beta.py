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

def create_schedule(people, start_month, year=2025, filename='schedule.xlsx'):
    # Generate start date aligned to Monday
    month_number = list(calendar.month_name).index(start_month.capitalize())
    start_date = datetime(year, month_number, 1)
    start_date += timedelta(days=(0 - start_date.weekday()) % 7)

    # Estimate number of 2-week blocks needed (3 people per block)
    num_blocks = len(people) // 3 + (1 if len(people) % 3 != 0 else 0)

    # Create 2-week non-overlapping date blocks
    blocks = []
    d = start_date
    for _ in range(num_blocks * 2):  # extra buffer for potential skipped weeks
        blocks.append((d, d + timedelta(days=13)))
        d += timedelta(weeks=2)

    def week_overlaps(date_range, skip_start, skip_end):
        start, end = date_range
        return not (end < skip_start or start > skip_end)

    # Define skip periods (Christmas and the following week)
    christmas_start = datetime(year, 12, 24)
    christmas_end = datetime(year, 12, 30)
    following_week_start = christmas_end + timedelta(days=1)
    following_week_end = following_week_start + timedelta(days=6)

    # Filter out skipped blocks
    blocks = [b for b in blocks if not (
        week_overlaps(b, christmas_start, christmas_end) or
        week_overlaps(b, following_week_start, following_week_end)
    )]

    # Adjust actual number of usable blocks
    blocks = blocks[:len(people) // 3 + (1 if len(people) % 3 != 0 else 0)]

    # Shuffle and assign 3 people per block
    random.shuffle(people)
    schedule = []
    for block in blocks:
        if len(people) >= 3:
            group = [people.pop() for _ in range(3)]
        else:
            group = random.sample(schedule[-1][1], 3) if schedule else random.sample(people, 3)
        schedule.append((block, group))

    # Prepare DataFrame
    df = pd.DataFrame([{
        'Week': f"{start.strftime('%a, %d %b')} - {end.strftime('%d %b')}",
        'Assigned': ', '.join(f"{p[0]} {p[1]}" for p in group)
    } for (start, end), group in sorted(schedule, key=lambda x: x[0])])

    # Add month and year to output filename
    base, ext = os.path.splitext(filename)
    filename = f"{base}_{start_month.lower()}{year}{ext}"

    # If file exists, rename it with random backup name
    if os.path.exists(filename):
        backup_name = f"{base}_{start_month.lower()}{year}_backup{random.randint(1000,9999)}{ext}"
        os.rename(filename, backup_name)

    # Save to Excel
    df.to_excel(filename, index=False, engine='openpyxl')



if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-names', type=str, required=True, help='Path to file with names and emails')
    parser.add_argument('-start', type=str, default=calendar.month_name[datetime.now().month], help='Start month (e.g., april)')
    parser.add_argument('-out', type=str, default='schedule.xlsx', help='Output Excel filename')
    args, unknown = parser.parse_known_args()

    # Compute default year based on current date and start month
    now = datetime.now()
    start_month_num = list(calendar.month_name).index(args.start.capitalize())
    if start_month_num == 0:
        raise ValueError("Invalid month name.")
    default_year = now.year if start_month_num >= now.month else now.year + 1

    # Add year with default logic
    parser.add_argument('-year', type=int, default=default_year, help=f'Year (default: {default_year})')
    args = parser.parse_args()

    people = load_people(args.names)
    create_schedule(people, args.start, year=args.year, filename=args.out)