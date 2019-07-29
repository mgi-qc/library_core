__author__ = 'Thomas Antonacci'


import csv
import os
import sys
import glob
from datetime import datetime
import xlrd
import webbrowser
import subprocess

mmddyy = datetime.now().strftime('%m%d%y')

excel_files = glob.glob('*.excel')

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

#get csv files...
csv_files = glob.glob('dilution_drop_off*.csv')
xls_files = glob.glob('dilution_drop_off*.xls')

if len(csv_files) == 0:
    print('Dilution files not found!')
    sys.exit()


# Create file and read in different headings for pipelines
outfile_header_list = ['Barcode', 'Source BC', 'Content_Desc','Total_DNA (ng)','Volume (ul)','Outgoing Queue Work Order','Outgoing Queue Work Order Pipeline','Outgoing Queue Work Order Description','BC_link','WO_link']


#Read in pipeline file and load pipeline lists
while True:
    try:
        with open('pipelines_file.csv', 'r+') as pipe_f:
            wgs_pipe_list = []
            exome_pipe_list = []
            other_pipe_list = []

            pipe_read = csv.DictReader(pipe_f,delimiter=',')
            pipe_head = pipe_read.fieldnames
            for line in pipe_read:
                if line['Pipeline'] == 'e':
                    exome_pipe_list.append(line['Name'])
                elif line['Pipeline'] == 'w':
                    wgs_pipe_list.append(line['Name'])
                elif line['Pipeline'] == 'o':
                    other_pipe_list.append(line['Name'])
        break
    except:
        print('Pipeline File not found\nBuilding pipelines_file.csv...')
        with open('pipelines_file.csv', 'w') as pipes:
            pipes.write('Name,Pipeline')


#Determine Pipeline, update pipeline lists if needed, and return tag
def check_pipeline(line):
    if line['Outgoing Queue Work Order Pipeline'] in exome_pipe_list:
        return 'e'
    elif line['Outgoing Queue Work Order Pipeline'] in wgs_pipe_list:
        return 'w'
    elif line['Outgoing Queue Work Order Pipeline'] in other_pipe_list:
        return 'o'
    else:

        pipeline_in = input('Is this WGS, Exome, or other: {}\nEnter w,e, or o: '.format(line['Outgoing Queue Work Order Pipeline']))

        while True:
            if pipeline_in == 'w':
                wgs_pipe_list.append(line['Outgoing Queue Work Order Pipeline'])
                return  'w'
            elif pipeline_in == 'e':
                exome_pipe_list.append(line['Outgoing Queue Work Order Pipeline'])
                return 'e'
            elif pipeline_in == 'o':
                other_pipe_list.append(line['Outgoing Queue Work Order Pipeline'])
                return 'o'
            else:
                pipeline_in = input('Is this WGS, Exome, or other: {}\nPlease enter either w,e, or o: '.format(line['Outgoing Queue Work Order Pipeline']))


#Check if woid is current, existing, or new
def check_woid(curr_woid, line, count_dict):
    if line['Outgoing Queue Work Order'] == curr_woid:
        return 'c'
    elif line['Outgoing Queue Work Order'] in count_dict:
        return 'ex'
    else:
        return 'n'


#adds given line to file given by writer object and updates the sample count in dict
def add_line_to_file(line, writer,count_dict,barcode_dict):
    line_dict = {}
    for head in outfile_header_list[:-2]:
        line_dict[head] = line[head]
    line_dict['BC_link'] = 'https://imp-lims.gsc.wustl.edu/entity/barcode/' + line['Barcode']
    #https://imp-lims.gsc.wustl.edu/entity/barcode/4v1Sqn
    line_dict['WO_link'] = 'https://imp-lims.gsc.wustl.edu/entity/setup-work-order/' + line['Outgoing Queue Work Order'].replace('.0','')
    #https://imp-lims.gsc.wustl.edu/entity/setup-work-order/2858611
    barcode_dict[line['Barcode']] = line['Outgoing Queue Work Order']

    writer.writerow(line_dict)
    count_dict[line['Outgoing Queue Work Order']] += 1


#closes given old file and, new temp file, updates the count and plate dict
def close_old_open_new(old_file, old_woid, old_pipe, new_woid, new_pipe, count_dict, reason):
    if reason == 'ini':
        new_file = open('temp_plate_file', 'w')
        count_dict[new_woid] = 0
        dict_writer = csv.DictWriter(new_file, delimiter=',', fieldnames=outfile_header_list)
        dict_writer.writeheader()

        return dict_writer,new_file

    else:
        filename = old_woid.replace('.0','') + '_' + str(count_dict[old_woid]) + '_Frag_Temp_' + old_pipe + '_'  + mmddyy + '.csv'
        old_file.close()
        os.rename(old_file.name, filename)

        if reason == 'new':
            new_file = open('temp_plate_file', 'w')
            count_dict[new_woid] = 0
            dict_writer = csv.DictWriter(new_file, delimiter=',', fieldnames=outfile_header_list)
            dict_writer.writeheader()
            return dict_writer, new_file

        elif reason == 'existing':
            ex_filename = new_woid.replace('.0','') + '_' + str(count_dict[new_woid]) +  '_Frag_Temp_' + new_pipe + '_' + mmddyy + '.csv'
            new_file = open(ex_filename, 'a')
            dict_writer = csv.DictWriter(new_file, delimiter=',',fieldnames=outfile_header_list)
            return dict_writer, new_file


#Close last file after all samples read it
def term_file(old_file, old_woid, old_pipe, count_dict):
    filename = old_woid.replace('.0','') + '_' + str(count_dict[old_woid]) + '_Frag_Temp_' + old_pipe + '_' + mmddyy + '.csv'
    old_file.close()
    os.rename(old_file.name, filename)


"""MAIN"""

#Initialize counting dicts and line variables
count_dict = {}
barcode_dict = {}
ini = True
current_wo = None
current_pipe = None

#Create Frag files and fill dicts
with open('others.{}.csv'.format(mmddyy), 'w') as otherf:

    other_file_d = csv.DictWriter(otherf, delimiter=',', fieldnames=outfile_header_list)

    for file in csv_files:

        with open(file,'r') as rf:

            next(rf)
            ddo_dict = csv.DictReader(rf, delimiter=',')
            header = ddo_dict.fieldnames

            for line in ddo_dict:

                #Initialize on first line
                if ini == True:

                    if check_pipeline(line) == 'e':
                         writer,o_file = close_old_open_new(None,None,None,line['Outgoing Queue Work Order'],'Exome',count_dict,'ini')
                         add_line_to_file(line, writer, count_dict,barcode_dict)
                         current_wo = line['Outgoing Queue Work Order']
                         current_pipe = 'Exome'
                    elif check_pipeline(line) == 'w':
                         writer,o_file = close_old_open_new(None, None, None, line['Outgoing Queue Work Order'], 'WGS', count_dict, 'ini')
                         add_line_to_file(line, writer, count_dict,barcode_dict)
                         current_wo = line['Outgoing Queue Work Order']
                         current_pipe = 'WGS'
                    elif check_pipeline(line) == 'o':
                        add_line_to_file(line,other_file_d,count_dict,barcode_dict)
                        current_wo = line['Outgoing Queue Work Order']
                        current_pipe = 'Other'
                    ini = False
                else:
                    #Check WO -> Check pipeline -> add to appropriate file/ will also open and close files as necessary
                     if check_woid(current_wo, line, count_dict) == 'c':
                        add_line_to_file(line,writer,count_dict,barcode_dict)

                     elif check_woid(current_wo,line,count_dict) == 'n':
                        if check_pipeline(line) == 'e':
                            writer, o_file = close_old_open_new(o_file,current_wo,current_pipe,line['Outgoing Queue Work Order'],'Exome',count_dict,'new')
                            add_line_to_file(line,writer,count_dict,barcode_dict)
                            current_wo = line['Outgoing Queue Work Order']
                            current_pipe = 'Exome'

                        elif check_pipeline(line) == 'w':
                            writer, o_file = close_old_open_new(o_file,current_wo,current_pipe,line['Outgoing Queue Work Order'],'WGS',count_dict,'new')
                            add_line_to_file(line,writer,count_dict,barcode_dict)
                            current_wo = line['Outgoing Queue Work Order']
                            current_pipe = 'WGS'

                        elif check_pipeline(line) == 'o':
                            add_line_to_file(line,other_file_d,count_dict,barcode_dict)

                     elif check_woid(current_wo,line,count_dict) == 'ex':
                         if check_pipeline(line) == 'e':
                             writer, o_file = close_old_open_new(o_file, current_wo, current_pipe,line['Outgoing Queue Work Order'], 'Exome', count_dict, 'existing')
                             add_line_to_file(line, writer, count_dict,barcode_dict)
                             current_wo = line['Outgoing Queue Work Order']
                             current_pipe = 'Exome'

                         elif check_pipeline(line) == 'w':
                             writer, o_file = close_old_open_new(o_file, current_wo, current_pipe,line['Outgoing Queue Work Order'], 'WGS', count_dict, 'existing')
                             add_line_to_file(line,writer,count_dict,barcode_dict)
                             current_wo = line['Outgoing Queue Work Order']
                             current_pipe = 'WGS'

                         elif check_pipeline(line) == 'o':
                             add_line_to_file(line, other_file_d, count_dict,barcode_dict)
    #Close last Frag Temp file
    term_file(o_file,current_wo,current_pipe,count_dict)

frag_files = glob.glob('*_Frag_Temp_*_*.csv')


#Open Freezer location file for work orders
for order in count_dict.keys():
    wo_bcs = []
    for bc in barcode_dict.keys():
        if barcode_dict[bc] == order:
            wo_bcs.append(bc)
    freezerURL = 'https://imp-lims.gsc.wustl.edu/gsc/report/barcode/results?report_type=freezer_loc&override_cache='
    for bc in wo_bcs:
        freezerURL += '&barcode=' + bc
    print('Work Order: ' + order.replace('.0',''))
    print(freezerURL + '\n')
    webbrowser.get('chrome').open_new_tab(freezerURL)


#update pipeline file with any changes to pipeline lists
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

    for name in other_pipe_list:
        pipe_dict['Pipeline'] = 'o'
        pipe_dict['Name'] = name
        pipe_dict_write.writerow(pipe_dict)


#Make processed directory if needed
if not os.path.exists('processed/dilution_drop_off_files/'):
    os.makedirs('processed/dilution_drop_off_files/')


#move dilution drop off files to processed directory; adds processed date to end of file
for file in csv_files:
    os.rename(file, 'processed/dilution_drop_off_files/' + file.split('.')[0] +'_' + mmddyy + '.' + file.split('.')[1])
for file in xls_files:
    os.rename(file, 'processed/dilution_drop_off_files/' + file.split('.')[0] +'_' + mmddyy + '.' + file.split('.')[1])

while True:
    cont = input('Are you ready to continue? (y/n): ')
    if cont == 'y':
        subprocess.run(['python3.7', 'add_freezer_loc.py'])
        break
    elif cont == 'n':
        print('Please move the freezer loc files to the RB_Dropoff directory and enter "y" to continue')