#!/usr/bin/env python3
# coding: utf-8

import pandas as pd
import argparse
import re
import unicodedata
from pathlib import Path
from datetime import datetime, date, timedelta

CUSTOM_TIME_RANGE_EXPR = r'^@\[(\d+:\d+ \d+:\d+)\]'
BEGIN_HOUR = '9:00'
BEGIN_ROW = 6
DATE_COLUMN = 'Date'

def normalize_string(string):
    return ''.join((c for c in unicodedata.normalize('NFD', string) if unicodedata.category(c) != 'Mn'))

def convert_time(time):
    h, m = map(int, time.split(':'))
    unit = 'AM' if h < 12 or h == 24 else 'PM'
    h = h%12 if h%12 != 0 else 12
    return f'{h}:{m:02} {unit}'
    
def parse_time_range(time_range):
    return map(convert_time, time_range.split())


def parse_file(path):
    if not Path(path).exists():
        raise argparse.ArgumentTypeError('invalid file path')
    return Path(path)

def rename_columns(df):
    day_index = -1
    hours_rows = df.iloc[2]
    columns = []
    for time_range in hours_rows:
        if pd.isna(time_range):
            columns.append(None)
            continue
        begin, end = parse_time_range(time_range)
        if BEGIN_HOUR in begin:
            day_index += 1
        columns.append((day_index, begin, end))
    columns[1] = DATE_COLUMN
    df.columns = columns

def extract_timetable(df, groups):
    permissions_mapper = {
        'algo': groups.algo,
        'reseaux': groups.net,
        'gla': groups.gla,
        'pfa': groups.pfa,
        'sys': groups.sys,
        'log': groups.log,
        'compilation': groups.comp,
        'o&a': groups.oa,
    }
    timetable = {
        'Subject': [],
        'Start Date': [],
        'End Date': [],
        'Start Time': [],
        'End Time': []
    }

    def update_current_week(line_index):
        date = df.loc[line_index, DATE_COLUMN]
        if pd.isna(date):
            return None
        date = re.sub(r'\d+$', lambda x: f'20{x.group(0)}', date)
        return datetime.strptime(date, '%d/%m/%Y')

    def update_timetable(activity, date, begin, end):
        nonlocal timetable
        timetable['Subject'].append(activity)
        timetable['Start Date'].append(date)
        timetable['End Date'].append(date)
        timetable['Start Time'].append(begin)
        timetable['End Time'].append(end)

    def is_concern_by_activity(activity):
        nonlocal permissions_mapper
        content = normalize_string(activity.lower())
        if not 'tp' in content and not 'td' in content:
            return True
        for key, group in permissions_mapper.items():
            if key in content:
                if not group:
                    print(f'Warning: We found an activity but you don\'t provide any groups ({key})')
                    return False
                return group in content

    def parse_row(row_items):
        for key, value in row_items:
            if pd.isna(key) or key == DATE_COLUMN:
                continue
            day_index, begin, end = key
            if pd.isna(value):
                continue
            activity = value
            if not is_concern_by_activity(activity):
                continue
            if match := re.match(CUSTOM_TIME_RANGE_EXPR, activity):
                    activity = activity.replace(match.group(0), '')
                    begin, end = parse_time_range(match.group(1))
            date = (current_week + timedelta(days=day_index)).strftime("%m/%d/%y")
            update_timetable(activity, date, begin, end)

    current_week = None
    for line_index in range(BEGIN_ROW, len(df)):
        current_week = update_current_week(line_index) or current_week
        if current_week:
            parse_row(df.loc[line_index].iteritems())
    return timetable


parser = argparse.ArgumentParser(description="Quickly convert your student timetable into google calendar events")
parser.add_argument('file',type=parse_file,  help='student timetable in .xlsx')
parser.add_argument('--algo', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='algorithm group')
parser.add_argument('--log', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='logical group')
parser.add_argument('--pfa', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='pfa group')
parser.add_argument('--gla', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='gla group')
parser.add_argument('--net', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='network group')
parser.add_argument('--sys', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='system group')
parser.add_argument('--comp', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='compilation group')
parser.add_argument('--oa', type=str, choices=['g1', 'g2', 'g3', 'g4', 'g5', 'g6', 'g7', 'g8'], help='O&A group')
parser.add_argument('--output', type=str, default='timetable.csv', help='.csv output file')


def main():
    args = parser.parse_args()
    used_keys = filter(lambda x: args.__getattribute__(x), ['algo', 'log', 'pfa', 'gla', 'net', 'sys', 'comp', 'oa'])
    a = ', '.join(f'{k}={args.__getattribute__(k)}' for k in used_keys)
    print(f'Converting \'{args.file}\' into google calendar events with :\n{a}', end='\n\n')
    df = pd.read_excel(args.file)
    rename_columns(df)
    timetable = extract_timetable(df, args)
    pd.DataFrame(timetable).to_csv(args.output, index=False, header=True)
    print(f'✅ Your \'{args.output}\' file is ready to be imported on google calendar.')

if __name__ == '__main__':
    main()
