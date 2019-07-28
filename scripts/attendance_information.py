#! /usr/bin/env python
'''
 attendance_information.py "Attendance_2\February 2018 Attendance with Demographics.xlsx" "Dallas" "Tarrant" "Collin"

import os
# TARGET =
import openpyxl
from collections import Counter
wb = openpyxl.load_workbook(TARGET)
# event reporting
import json

CATAGORY_COL = 'C'
stats = {}
for sheet in wb:
    stats[sheet.title] = Counter([cell.value for cell in sheet[CATAGORY_COL]])


with open('event_count.json', 'w') as file:
    file.write(json.dumps(stats))

Column Y  is “Sexual Orientation”
For sorting add additional field of “Bereaved” for Participant Type from Column P.

'''

import os.path as osp
from common import *
from collections import OrderedDict
from datetime import datetime
from math import isnan
import argparse
from glob import glob
from itertools import chain

parser = argparse.ArgumentParser()
parser.add_argument('workbooks',action='store',nargs='+')
parser.add_argument('--sheets','-s',action='store',nargs='*')
parser.add_argument('--glob','-g',action='store_true')


class BadPersonExcpetion(Exception):
    pass

class person:
    """
    person:
        criteria: column N "Participant Status" MUST BE "Participant"

        has fields:
            first_name: col "G" ("First Name")
            last_name: col "H" ("Last Name")
            type: col "O" ("Participant Type")
            gender: col "P" ("Gender")
            age_group: col "Q" ("Age Group")
            ethnicity: col "R" ("Ethnicity")
            insurance: col "S" ("Insurance")
            employment: col "T" ("Type of Employment")
            income: col "U" ("Income")
            cancer: col "V" ("Primary Cancer")
            sexual_orientation: col "Y" ("Sexual Orentation")
            earliest attended event: earliest(col "A" ("Attendance Date"))
    """
    good_keys = {'pid', 'first_name', 'last_name', 'date', 'clubhouse',
                 'volunteer', 'participant_stat', 'participant_type',
                 'gender', 'age_group', 'ethnicity', 'insurance',
                 'employment', 'income', 'cancer','sexual_orientation'
                 }

    def __init__(self, record):
        self.info = {key:(record.get(key,'N/A') or 'N/A') for key in self.good_keys}
        if not isinstance(self.info['age_group'],int):
            # print('fixing_age for {}'.format(self.full_name) )
            # print('was {}, now nan'.format(self.info['age_group']) )
            self.info['age_group'] = float('nan')
        if not isinstance( self.info['date'],datetime):
             self.info['date'] = datetime(self.info['date'])

    @property
    def full_name(self):
        return '_'.join(str(self.info[x]) for x in ('first_name','last_name','pid'))

    def __eq__(self,other):
        return hash(self) == hash(other)

    def __hash__(self):
        return hash(self.full_name)

    def check_date(self,record):
        if self.info['date'] > record['date']:
            # print('')
            # print('common record for {}'.format(self.full_name))
            # print('replacing date {} with newer value {}'.format(self.info['date'] ,record['date'] ))
            self.info['date'] = record['date']

    def __str__(self):
        return ','.join(str(self.info[key]) for key in title_list if key in self.info.keys())

title_mapping = {
                 'date': "A" ,
                 'clubhouse':"E",
                 'first_name':  "H",
                 'last_name': "I",
                 'pid':"J",
                 'volunteer':'L',
                 'participant_stat':"P",
                 'participant_type': "Q",
                 'gender':  "R",
                 'age_group': "S" ,
                 'ethnicity': "T" ,
                 'insurance':  "U" ,
                 'employment':  "V" ,
                 'income':  "W" ,
                 'cancer':  "X" ,
                 'sexual_orientation': "Y",
                 }
title_list =('pid', 'first_name', 'last_name', 'date', 'clubhouse','volunteer', 'participant_stat', 'participant_type', 'gender',
'age_group', 'ethnicity', 'insurance', 'employment', 'income', 'cancer','sexual_orientation')

def parse_row(row,col_mapping):
    x = {}
    # print('')
    for key,val in col_mapping.items():
        # print(key,row[col2num(val)-1].value,sep=":")
        x[key] = row[col2num(val)-1].value
    # print('')
    # x ={key: row[col2num(val)-1].value for key,val in col_mapping.items()}
    return x


def print_dict_with_persons(dict):
        for _type,_ppl in dict.items():
            print(_type)
            for person in _ppl:
                print(person.info)


if __name__ == '__main__':
    import sys
    args = parser.parse_args()
    # print(vars(args))
    # sys.exit()
    # paths = sys.argv[1:]
    persons = []
    if args.glob:
        print('wildcarded')
        args.workbooks = chain.from_iterable([glob(osp.join(TEST_DIR,wb)) for wb in args.workbooks])
    for path in args.workbooks:
        print(path)
        wb= get_workbook(FakeArgs(path,None))[0]
        rets = {}
        for sheet in args.sheets:
            print(sheet)
            try:
                ws = wb.get_sheet_by_name(sheet)
                for row in ws:
                    try:
                        record = parse_row(row,title_mapping)
                        per = person(record)
                        index = next((i for i,x in enumerate(persons) if x == per), None)
                        if index is not None:
                            # print('found person {}'.format(per.full_name))
                            # print('record date:{}'.format(per.info['date']))
                            # print('found person in persons: {}'.format(persons[index].full_name))
                            # print('persons date:{}'.format(persons[index].info['date']))
                            # print(persons[index])
                            # print(record)
                            persons[index].check_date(record)

                        else:
                            persons.append(per)
                    except BadPersonExcpetion:
                        pass
                    except Exception as e:
                        print('throwing row:{} unable to parse...'.format(row[0].row))
                        pass
                rets[ws] = persons
            except Exception as e:
                print(e)
                print(row)
                print(record)
                print(per)
            wb.close()


    # # dump records to file
    # for place,persons in rets.items():
    # for persons in rets.values():
        # print(place)
    place = 'results'
    month = osp.dirname(path)
    path = osp.join(TEST_DIR,month)
    with open(osp.join(osp.dirname(path),'{}_{}.csv'.format(month,place)),'w') as outfile:
        outfile.write('{}\n'.format(','.join(str(key) for key in title_list)))
        for person in persons:
            pstr = str(person)
            outfile.write('{}\n'.format(pstr))
    print('\n\n')


    # persons = [person for person in persons if person.info['date'].month == 2]

    # # analysis of good records:
    participant_types = set(person.info['participant_type'] for person in persons)
    # print(participant_types)


    split_by_participant_type = {participant_type:[person for person in persons if person.info['participant_type'] == participant_type]
                                    for participant_type in participant_types}

    # print('split_by_participant_type:')
    # print_dict_with_persons(split_by_participant_type)


    def split_participants_by_age(_participant_type):
        temp = split_by_participant_type[_participant_type]
        return {
               '>=18': [person for person in temp if person.info['age_group'] >= 18],
               '<18':[person for person in temp if person.info['age_group'] < 18],
               'Unknown':[person for person in temp if isnan(person.info['age_group'] )]
              }

    # Living With Cancer
    split_cancer_by_age = split_participants_by_age('Living with Cancer')

    # support Person
    support_people = split_participants_by_age('Support Person')

    # Survivor
    survivors = split_participants_by_age('Survivor')

    # print('split_cancer_by_age:')
    # print_dict_with_persons(split_cancer_by_age)
    # print('support_people:')
    # print_dict_with_persons(support_people)
    # print('survivors:')
    # print_dict_with_persons(survivors)

    # volunteers
    volunteers = {'volunteers':[person for person in persons if person.info['volunteer'].lower() == 'true']}
    # print('volunteers:')
    # print_dict_with_persons(volunteers)

    def split_by_clubhouse(dict_of_persons):
        clubhouses = {person.info['clubhouse'] for person in persons}
        # print(clubhouses)
        for category,_persons in dict_of_persons.items():
            dict_of_persons[category] = {clubhouse:[person for person in _persons if person.info['clubhouse'] == clubhouse] for clubhouse in clubhouses}

    for name,split in [('split_cancer_by_age',split_cancer_by_age),
                       ('support_people',support_people),
                       ('survivors',survivors),
                       ('volunteers',volunteers)
                       ]:

        split_by_clubhouse(split)
        # print(name)
        # print(split)

    with open(osp.join(osp.dirname(path),'{}_{}.csv'.format('stats','8')),'w') as outfile:
        for name,split in [('split_cancer_by_age',split_cancer_by_age),
                               ('support_people',support_people),
                               ('survivors',survivors),
                               ('volunteers',volunteers)
                              ]:
            outfile.write('{}\n'.format(name))
            print(name)
            for cat,clubhouse_split in split.items():
                outfile.write('{}\n'.format(cat))
                print('\tcat:{}'.format(cat))
                for clubhouse,_persons in clubhouse_split.items():
                    outfile.write('{}\n'.format(clubhouse))
                    print('\t\t{}: '.format(clubhouse), end='')
                    print(len(_persons))
                    outfile.write('{}\n'.format(','.join(entry for entry in title_list)))
                    for person in _persons:
                        # print(person)
                        pepstr = str(person)
                        outfile.write('{}\n'.format(pepstr))
    print('\n\n')


    wb.close()
    sys.exit(0)
