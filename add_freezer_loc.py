__author__ = 'Thomas Antonacci'

"""Add freezer location to given work order frag file"""

import smartsheet

import csv
import os
import sys
import glob
from datetime import datetime, timedelta
import argparse
import subprocess
import xlrd
from string import Template

API_KEY = os.environ.get('SMRT_API')

if API_KEY is None:
    sys.exit('Api key not found')

smart_sheet_client = smartsheet.Smartsheet(API_KEY)
smart_sheet_client.errors_as_exceptions(True)

def create_folder(new_folder_name, location_id,location_tag):

    if location_tag == 'f':
        response = smart_sheet_client.Folders.create_folder_in_folder(str(location_id), new_folder_name)

    elif location_tag == 'w':
        response = smart_sheet_client.Workspaces.create_folder_in_workspace(str(location_id), new_folder_name)

    elif location_tag == 'h':
        response = smart_sheet_client.Home.create_folder(new_folder_name)

    return response

def create_workspace_home(workspace_name):
    # create WRKSP command
    workspace = smart_sheet_client.Workspaces.create_workspace(smartsheet.models.Workspace({'name': workspace_name}))
    return workspace

def get_sheet_list(location_id, location_tag):
    #Read in all sheets for account
    if location_tag == 'a':
        ssin = smart_sheet_client.Sheets.list_sheets(include="attachments,source,workspaces",include_all=True)
        sheets_list = ssin.data

    elif location_tag == 'f' or location_tag == 'w':
        location_object = get_object(str(location_id), location_tag)
        sheets_list = location_object.sheets

    return sheets_list

def get_folder_list(location_id, location_tag):

    if location_tag == 'f' or location_tag == 'w':
        location_object = get_object(str(location_id), location_tag)
        folders_list = location_object.folders

    elif location_tag == 'a':
        folders_list = smart_sheet_client.Home.list_folders(include_all=True)

    return folders_list

def get_workspace_list():
    # list WRKSPs command
    read_in = smart_sheet_client.Workspaces.list_workspaces(include_all=True)
    workspaces = read_in.data
    return workspaces

def get_object(object_id, object_tag):

    if object_tag == 'f':
        obj = smart_sheet_client.Folders.get_folder(str(object_id))
    elif object_tag == 'w':
        obj = smart_sheet_client.Workspaces.get_workspace(str(object_id))
    elif object_tag == 's':
        obj = smart_sheet_client.Sheets.get_sheet(str(object_id))
    return obj

mmddyy = datetime.now().strftime('%m%d%y')

#get files and set up lists and dicts
freeze_loc_files = glob.glob('*Freezer Loc*')
frag_files = glob.glob('*_Frag_Plate_*.csv')
new_files = []

freeze_loc_dict = {}
found_dict = {}
not_found_dict = {}
barcd_fnd = []


#Pair freezer locs with barcode
for file in freeze_loc_files:
    with open(file, 'r') as ff:

        fr_read = csv.DictReader(ff, delimiter = ',')
        ff_header = fr_read.fieldnames
        for line in fr_read:
            freeze_loc_dict[line['Barcode']] = line['Freezer_Loc']


#find barcode in frag file fills in row in new file.
for file in frag_files:
    loc_added = False
    with open(file, 'r') as fragfile, open(file.replace('.csv','_new'), 'w') as newfragfile:
        frg_read = csv.DictReader(fragfile, delimiter = ',')
        outfile_header_list = frg_read.fieldnames + ['Freezer_Loc']

        new_files.append(newfragfile.name)
        newfragwrt = csv.DictWriter(newfragfile,delimiter = ',',fieldnames=outfile_header_list)
        newfragwrt.writeheader()

        frg_header = frg_read.fieldnames
        for line in frg_read:
            if line['Barcode'] in freeze_loc_dict:
                loc_added = True
                line_dict = {}
                for head in outfile_header_list[:-1]:
                    line_dict[head] = line[head]
                line_dict[outfile_header_list[-1]] = freeze_loc_dict.pop(line['Barcode'])
                newfragwrt.writerow(line_dict)

    if loc_added:
        barcd_fnd.append(newfragfile.name)

#id what barcodes not found: Need output
for key in freeze_loc_dict.keys():
    not_found_dict[key] = freeze_loc_dict.pop(key)

#removes unused files: Need output
for file in new_files:
    if file not in barcd_fnd:
        os.remove(file)

#Rename files to .csv's
for file in barcd_fnd:
    os.rename(file, file.replace('_new','.csv'))

#Make processed freezer files dir if needed
if not os.path.exists('processed/freezer_loc_files'):
    os.makedirs('processed/freezer_loc_files')

#Move processed freezer location files; adds processed date to file name
for file in freeze_loc_files:
    os.rename(file, 'processed/freezer_loc_files/' + file.split('.')[0] +'_' + mmddyy + '_' + file.split('.')[1])