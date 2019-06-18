__author__ = 'Thomas Antonacci'

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

smart_sheet_client = smartsheet.Smartsheet('attvod4mfscdwx3658bwf1sgbb')
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

"""Read in Dilution Drop off and generate plate and tracking sheets as well as upload to/ update Smartsheet

TODO:   Get DDO sheet and convert to csv ---------
        Read in info and begin sort --------
        assign to plate file------
        update original sheets with location
        update/upload to smartsheet
        """

mmddyy = datetime.now().strftime('%m%d%y')

#os.rename('dilution_drop_off.xls', 'dilution_drop_off.excel')
excel_files = glob.glob('*.excel')
print(excel_files)
for file in excel_files:

    xls_file = '{}.xls'.format(file.split('.excel')[0])
    csv_file = '{}.csv'.format(file.split('.excel')[0])
    os.rename(file, xls_file)

    with open(csv_file, 'w') as fh:
        csv_write = csv.writer(fh)

        xls_openf = xlrd.open_workbook(xls_file)
        xls_sheet = xls_openf.sheet_by_index(0)

        for rownum in range(xls_sheet.nrows):
            csv_write.writerow(xls_sheet.row_values(rownum))

#get csv files... need more formatting information

csv_file = 'dilution_drop_off.csv'

# Create file and read in different headings for pipelines
with open('pipelines_file.csv', 'r+') as pipe_f:
    wgs_pipe_list = []
    exome_pipe_list = []

    pipe_read = csv.DictReader(pipe_f,delimiter=',')
    pipe_head = pipe_read.fieldnames
    for line in pipe_read:
        if line['Pipeline'] == 'e':
            exome_pipe_list.append(line['Name'])
        elif line['Pipeline'] == 'w':
            wgs_pipe_list.append(line['Name'])

#open and read in csv
with open(csv_file,'r+') as rf, open('others.{}.csv'.format(mmddyy),'w') as otherf:
    next(rf)
    ddo_dict = csv.DictReader(rf,delimiter=',')
    header = ddo_dict.fieldnames

    outfile_header_list = ['Source BC', 'Content_Desc','Total_DNA (ng)','Volume (ul)','Outgoing Queue Work Order','Outgoing Queue Work Order Pipeline','Outgoing Queue Work Order Description']

    other_file_d = csv.DictWriter(otherf, delimiter=',',fieldnames= outfile_header_list)

    #break up dict; sort by pipeline

    wgs_count = 0
    exome_count = 0
    wgs_plate_count = 0
    exome_plate_count = 0

    exomef = open('exome_plates_temp_file','w')
    exome_file_d = csv.DictWriter(exomef, delimiter=',', fieldnames=outfile_header_list)
    exome_woids = []

    wgsf = open('wgs_plates_temp_file','w')
    wgs_file_d = csv.DictWriter(wgsf, delimiter=',', fieldnames=outfile_header_list)
    wgs_woids = []


    for line in ddo_dict:

        if exome_count == 96:
            exome_filename = 'Fragmentation_Plate_Exome_' + str(exome_plate_count) + '_' + '_'.join(exome_woids) + '.csv'
            exomef.close()
            os.rename('exome_plates_temp_file',exome_filename)
            exomef = open('exome_plates_temp_file','w')
            exome_file_d = csv.DictWriter(exomef, delimiter=',', fieldnames=outfile_header_list)
            exome_woids = []
            exome_count = 0
            exome_plate_count += 1

        if wgs_count == 96:
            wgs_filename = 'Fragmentation_Plate_WGS_' + str(wgs_plate_count) + '_' + '_'.join(wgs_woids) + '.csv'
            wgsf.close()
            os.rename('wgs_plates_temp_file',wgs_filename)
            wgsf = open('wgs_plates_temp_file'.format(wgs_plate_count),'w')
            wgs_file_d = csv.DictWriter(wgsf, delimiter=',', fieldnames=outfile_header_list)
            wgs_woids = []
            wgs_count = 0
            wgs_plate_count += 1

        line_dict = {}
        if line['Outgoing Queue Work Order Pipeline'] in wgs_pipe_list:

            for head in outfile_header_list:
                line_dict[head] = line[head]
            wgs_file_d.writerow(line_dict)
            if line['Outgoing Queue Work Order'] not in wgs_woids:
                wgs_woids.append(line['Outgoing Queue Work Order'])
            wgs_count += 1

        elif line['Outgoing Queue Work Order Pipeline'] in exome_pipe_list:

            for head in outfile_header_list:
                line_dict[head] = line[head]
            exome_file_d.writerow(line_dict)
            if line['Outgoing Queue Work Order'] not in exome_woids:
                exome_woids.append(line['Outgoing Queue Work Order'])
            exome_count += 1
        else:

            pipeline_in = input('Is this WGS, Exome, or other: {}\nEnter w,e, or o: '.format(line['Outgoing Queue Work Order Pipeline']))
            if pipeline_in == 'w':
                for head in outfile_header_list:
                    line_dict[head] = line[head]
                wgs_file_d.writerow(line_dict)
                wgs_count += 1
                wgs_pipe_list.append(line['Outgoing Queue Work Order Pipeline'])

            elif pipeline_in == 'e':
                for head in outfile_header_list:
                    line_dict[head] = line[head]
                exome_file_d.writerow(line_dict)
                exome_woids.append(line['Outgoing Queue Work Order'])
                exome_count += 1
                exome_pipe_list.append(line['Outgoing Queue Work Order Pipeline'])

            elif pipeline_in == 'o':
                other_file_d.writerow(line)
            else:
                input('Is this WGS, Exome, or other: {}\nPlease enter either w,e, or o: '.format(line['Outgoing Queue Work Order Pipeline']))
exome_pipe_list
if exome_count != 0:
    exomef.close()
    exome_filename = 'Fragmentation_Plate_Exome_' + str(exome_plate_count) + '_' + '_'.join(exome_woids) + '.csv'
    os.rename('exome_plates_temp_file',exome_filename)
if wgs_count != 0:
    wgsf.close()
    wgs_filename = 'Fragmentation_Plate_WGS_' + str(wgs_plate_count) + '_' +  '_'.join(wgs_woids) + '.csv'
    os.rename('wgs_plates_temp_file',wgs_filename)

with open('pipelines_file.csv', 'w') as pipe_f:
    pipe_dict_write = csv.DictWriter(pipe_f,delimiter=',',fieldnames=pipe_head)
    pipe_dict_write.writeheader()
    pipe_dict = {'Name' : None,'Pipeline':None}

    for name in wgs_pipe_list:
        pipe_dict['Pipeline'] = 'w'
        pipe_dict['Name'] = name
        pipe_dict_write.writerow(pipe_dict)

    for name in exome_pipe_list:
        pipe_dict['Pipeline'] = 'e'
        pipe_dict['Name'] = name
        pipe_dict_write.writerow(pipe_dict)

print('debug')
