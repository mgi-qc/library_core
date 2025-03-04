__author__ = 'Thomas Antonacci'

"""
Sort by pipeline-------
Sort by WO-------
Sort by freezer loc-------

build plates of 96

List of sample objects -> plates as many lists (output as plate files with WO and number samples in them)

Push to smart sheet
"""

import smartsheet
import csv
import os
import sys
import glob
from datetime import datetime
import webbrowser
import time
import copy
import math
import operator

#Smartsheet client ini
API_KEY = os.environ.get('SMRT_API')

if API_KEY is None:
    sys.exit('Api key not found')

smart_sheet_client = smartsheet.smartsheet.Smartsheet(API_KEY)
smart_sheet_client.errors_as_exceptions(True)

mmddyy = datetime.now().strftime('%m%d%y')

#Misc Functions
def is_int(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


#SmartSheet Functions:
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

#Classes for plate building
class sample:

    def __init__(self,name,loc,work_order,pipe,bc):
        self.name = name
        self.sourcebc = bc
        self.loc = loc
        self.work_order = work_order
        self.pipe = pipe
        self.plate = loc.split(' ')[-3]

class plate:

    def __init__(self):
        self.name = ''
        self.samples = []
        self.wo = []
        self.pipe = ''

    def set_name(self, name):
        self.name = name
        return None

    def set_samples(self, list_of_samples):
        self.samples = list_of_samples
        return None

    def set_work_orders(self, list_of_wos):
        self.wo = list_of_wos
        return None

    def set_pipe(self,pipe):
        if self.pipe is None:
            self.pipe = pipe
            return None
        elif self.pipe == pipe:
            return
        else:
            print('Pipe already set for Plate {}!'.format(self.name))
            return None

    def add_sample(self, sampl):
        if type(sampl) is list:
            for samp in sample:
                self.samples.append(samp)
                return None
        else:
            self.samples.append(sampl)
            return None

    def add_wo(self, workorder):
        if workorder in self.wo:
            return
        else:
            self.wo.append(workorder)
            return

    def remove_sample(self, sample):
        if type(sample) is list:
            for samp in sample:
                self.samples.remove(samp)
                return None
        else:
            self.samples.remove(sample)
            return None

    def clear_plate(self):
        self.name = None
        self.samples = None
        self.wo = None
        return None

class work_order:

    def __init__(self):
        self.name = ''
        self.in_plates = []
        self.out_plates = []
        self.samples = []
        self.pipe = None

    def set_name(self,name):
        self.name = name
        return

    def addInPlate(self,plate):
        self.in_plates.append(plate)
        return None

    def addOutPlate(self,plate):
        self.out_plates.append(plate)
        return None

    def add_sample(self, sampl):
        if type(sampl) is list:
            for samp in sample:
                self.samples.append(samp)
                return None
        else:
            self.samples.append(sampl)
            return None

#Functions and variables for sorting

capacity = 96

def add_to_bins(item,bins):
    for b in bins:
        bin_value = 0
        for t in b:
            bin_value += wo_dict[t]
        if bin_value + wo_dict[item] < capacity:
            b.append(item)
            return
    bins.append([])
    add_to_bins(item)
    return


def ffd(usable_list,bins):
    for item in usable_list:
        add_to_bins(item,bins)

"""
Read in Files and load objects(samples, plates and work orders) and lists
---------------------------------------------------
"""

#get frag_files
wo_frag_files = glob.glob('*Frag_Temp*.csv')
if not len(wo_frag_files) > 0:
    sys.exit('No Frag_Temp files found!')
wgs_files = []
exome_files = []

total_samples = 0

#get total samples
for file in wo_frag_files:
    num = int(file.split('_')[1])
    total_samples += num

#ini
samples_master = []
count = 0

#Divide pipes
for file in wo_frag_files:
    if 'WGS' in file:
        wgs_files.append(file)
    elif 'Exome' in file:
        exome_files.append(file)


#Read in the samples
for file in wgs_files:
    with open(file,'r') as rf:
        file_reader = csv.DictReader(rf, delimiter = ',')
        header = file_reader.fieldnames

        for line in file_reader:
            try:
                line['Freezer_Loc']
            except KeyError:
                sys.exit('Freezer Location not found in {}'.format(file))
            samples_master.append(sample(line['Barcode'],line['Freezer_Loc'],line['Outgoing Queue Work Order'] ,'WGS',line['Source BC']))
            count += 1

for file in exome_files:
    with open(file, 'r') as rf:
        file_reader = csv.DictReader(rf, delimiter=',')
        header = file_reader.fieldnames

        for line in file_reader:
            try:
                line['Freezer_Loc']
            except KeyError:
                sys.exit('Freezer Location not found in {}'.format(file))
            samples_master.append(sample(line['Barcode'], line['Freezer_Loc'], line['Outgoing Queue Work Order'], 'Exome',line['Source BC']))
            count += 1

"""
Get FFPE samples based on either file or BC
---------------------------------------------------
"""

FFPE_samples = []
print('Opening work order samples page in IMP...')
time.sleep(1.5)

#Open samples page URL for work orders
for file in wo_frag_files:
        wo = file.split('_')[0]
        FFPE_url = 'https://imp-lims.ris.wustl.edu/entity/setup-work-order/{}?setup_name={}&perspective=Sample'.format(wo.replace('.0',''),wo.replace('.0',''))
        webbrowser.get('chrome').open_new_tab(FFPE_url)

FFPE_present = input('Are there FFPE samples present? (y/n): ')
if FFPE_present == 'y':
    print('Would you like to:\n1. Enter comma separated list of barcodes\n2. Import Samples sheets from IMP(WIP)')
    while True:
        option = input('Please enter 1 or 2: ')
        if is_int(option) and 0 < int(option) <= 1:
            option = int(option)
            break
        elif is_int(option) and int(option) == 2:
            print('Option 2 is under construction; use option 1.')

    if option == 1:
        FFPE_list = input('Please enter the barcodes as a comma separated list(No white space): ')

        for barcode in FFPE_list.split(','):
            found = False
            for samp in samples_master:
                if samp.sourcebc == barcode:
                    FFPE_samples.append(samp)
                    found = True
            if not found and not FFPE_list == '':
                print('{} not found!'.format(barcode))
    if option == 2:
        exit('How did you get here...?')

"""
Build incoming plates
---------------------------------------------------
"""

plates_in_master = []
plate_in_count = 0

for samp in samples_master:
    if len(plates_in_master) == 0:
        plates_in_master.append(plate())
        plates_in_master[plate_in_count].name = samp.loc.split(' ')[-3]
        plates_in_master[plate_in_count].pipe = samp.pipe
        plates_in_master[plate_in_count].add_sample(samp)
        plates_in_master[plate_in_count].add_wo(samp.work_order)
        plate_in_count += 1
    else:
        found = False
        for plt in plates_in_master:
            if samp.loc.split(' ')[-3] == plt.name:
                plt.add_sample(samp)
                found = True
        if found == False:
            plates_in_master.append(plate())
            plates_in_master[plate_in_count].name = samp.loc.split(' ')[-3]
            plates_in_master[plate_in_count].pipe = samp.pipe
            plates_in_master[plate_in_count].add_sample(samp)
            plates_in_master[plate_in_count].add_wo(samp.work_order)
            plate_in_count += 1

#Build WO's (sort plates according to WO)
WO_master = []
wo_count = 0

for plt in plates_in_master:
    if len(WO_master) == 0:
        for wo in plt.wo:
            WO_master.append(work_order())
            WO_master[wo_count].name = wo
            WO_master[wo_count].addInPlate(plt)
            for samp in plt.samples:
                if WO_master[wo_count].name == samp.work_order:
                    WO_master[wo_count].add_sample(samp)
            wo_count += 1

    else:

        for wo in plt.wo:
            found = False
            for order in WO_master:
                if order.name == wo:
                    found = True
                    order.addInPlate(plt)
                    for samp in plt.samples:
                        if samp.work_order == wo:
                            order.add_sample(samp)

            if not found:
                WO_master.append(work_order())
                WO_master[wo_count].name = wo
                WO_master[wo_count].addInPlate(plt)
                for samp in plt.samples:
                    if WO_master[wo_count].name == wo:
                        WO_master[wo_count].add_sample(samp)
                wo_count += 1


"""
Presorting plates for Bin packing
-------------------------------------------------------------------
"""

#Temp Plates for building output:
temp_plates_master = []

#Plates bound for sort algorithm
Exome_sortable_plates = []
WGS_sortable_plates = []

easy_plates = []
hard_plates = []

#Outgoing plates
Exome_out_plates = []
WGS_out_plates = []

"""
get easy plates
    i.e. get plates of 96 samples with same WO as these can be handed off with no extra work
        """

for plt in plates_in_master:

    if len(plt.wo) == 1 and len(plt.samples) == 96:
        easy_plates.append(plt)
        temp_plates_master.append(plt)
    elif len(plt.wo) > 1:
        hard_plates.append(plt)
    else:
        temp_plates_master.append(plt)

#Break up plates with more than one work order
"""
Hard Plate --> WO1 plate |}> Temp plates
           --> WO2 plate |

"""
new_plates = []
new_count = 0
plt_pos = 0
for plt in hard_plates:
    wo = 'ini'
    for samp in plt.samples:
        found = False
        if samp.work_order == wo:
            new_plates[plt_pos].add_sample(samp)
            found = True
        elif wo == 'ini':
            new_plates.append(plate())
            new_plates[new_count].wo.append(samp.work_order)
            new_plates[new_count].pipe = samp.pipe
            new_plates[new_count].add_sample(samp)
            wo = samp.work_order
            new_count += 1
            found = True
        else:
            for nplt in new_plates:
                if samp.work_order in nplt.wo:
                    plt_pos = new_plates.index(nplt)
                    new_plates[plt_pos].add_sample(samp)
                    wo = samp.work_order
                    found = True
            if not found:
                new_plates.append(plate())
                new_plates[new_count].wo.append(samp.work_order)
                new_plates[new_count].pipe = samp.pipe
                new_plates[new_count].add_sample(samp)
                wo = samp.work_order
                plt_pos = new_count
                new_count += 1

#Add new formed plates
for plt in new_plates:
    temp_plates_master.append(plt)

#Sort temp plates into pipes
for plt in temp_plates_master:
    if plt.pipe == 'Exome':
        if plt in easy_plates:
            Exome_out_plates.append(plt)
        else:
            Exome_sortable_plates.append(plt)
    elif plt.pipe == 'WGS':
        if plt in easy_plates:
            WGS_out_plates.append(plt)
        else:
            WGS_sortable_plates.append(plt)


#Reconcile any partial plates with same wos
plt_del = []
for plt in Exome_sortable_plates:
    for pltc in Exome_sortable_plates:
        if plt.wo == pltc.wo and pltc not in plt_del and plt not in plt_del and pltc is not plt:
            for samp in plt.samples:
                pltc.samples.append(samp)
                plt_del.append(plt)

for plt in WGS_sortable_plates:
    for pltc in Exome_sortable_plates:
        if plt.wo == pltc.wo and pltc not in plt_del and plt not in plt_del and pltc is not plt:
            for samp in plt.samples:
                pltc.samples.append(samp)
                plt_del.append(plt)

#Delete reconciled plates:
for plt in plt_del:
    temp_plates_master.remove(plt)


#Break up temp plates w/ more than 96 samples
full_plates = []
for plt in Exome_sortable_plates:
    if len(plt.samples) > 96:
        count = 0

        for samp in plt.samples:
            if count == 96:
                full_plates.append(plt)
                Exome_sortable_plates.append(plate())
                Exome_sortable_plates[len(Exome_sortable_plates)-1].wo = plt.wo
                Exome_sortable_plates[len(Exome_sortable_plates) - 1].name = samp.loc.split(' ')[-3]
                Exome_sortable_plates[len(Exome_sortable_plates) - 1].pipe = 'Exome'
                temp_plates_master.append([len(Exome_sortable_plates) - 1])
                Exome_sortable_plates[len(Exome_sortable_plates) - 1].add_sample(samp)
            elif count > 96:
                Exome_sortable_plates[len(Exome_sortable_plates) - 1].add_sample(samp)
            else:
                count += 1

for plt in WGS_sortable_plates:
    if len(plt.samples) > 96:
        count = 0

        for samp in plt.samples:
            if count == 96:
                full_plates.append(plt)
                WGS_sortable_plates.append(plate())
                WGS_sortable_plates[len(WGS_sortable_plates) - 1].wo = plt.wo
                WGS_sortable_plates[len(WGS_sortable_plates) - 1].name = samp.loc.split(' ')[-3]
                WGS_sortable_plates[len(WGS_sortable_plates) - 1].pipe = 'WGS'
                temp_plates_master.append(WGS_sortable_plates[len(WGS_sortable_plates) - 1])
                WGS_sortable_plates[len(WGS_sortable_plates) - 1].add_sample(samp)
            elif count > 96:
                WGS_sortable_plates[len(WGS_sortable_plates) - 1].add_sample(samp)
            else:
                count += 1


#Add full plates to outgoing and remove from sortable and temp
for plt in full_plates:
    temp_plates_master.remove(plt)
    if plt.pipe == 'WGS':
        WGS_sortable_plates.remove(plt)
    elif plt.pipe == 'Exome':
        Exome_sortable_plates.remove(plt)


usable_keys = []


"""
Sort Exome
--------------------------------
"""
wo_dict = {}
usable_list = []
total = 0
plates = []

#Populate Dictionary
for plt in Exome_sortable_plates:
    wo_dict[plt.wo[0]] = len(plt.samples)
    total += len(plt.samples)

#Calc lower bound
low_bnd = total/capacity

#Load plates list with min needed
while len(plates) <= low_bnd:
    plates.append([])

#Sort Dict by # of samples
sorted_tuples = sorted(wo_dict.items(),key=operator.itemgetter(1), reverse=True)

#Load list of wo's
for item in sorted_tuples:
    usable_list.append(item[0])

#first fit decreasing method
if len(usable_list) > 0:
    ffd(usable_list,plates)


#Load Outgoing plates based on sol found by FFD
for plt in plates:
    Exome_out_plates.append(plate())
    Exome_out_plates[len(Exome_out_plates) - 1].pipe = 'Exome'
    for order in plt:
        for plte in Exome_sortable_plates:
            if order in plte.wo:
                Exome_out_plates[len(Exome_out_plates) - 1].add_wo(order)
                for samp in plte.samples:
                    Exome_out_plates[len(Exome_out_plates) - 1].add_sample(samp)

"""
Sort WGS
------------------------------------
"""
usable_list = []
wo_dict = {}
total = 0
plates.clear()

#Populate Dictionary
for plt in WGS_sortable_plates:
    wo_dict[plt.wo[0]] = len(plt.samples)
    total += len(plt.samples)

#Calc lower bound
low_bnd = total/capacity

#Load plates list with min needed
while len(plates) <= low_bnd and low_bnd != 0:
    plates.append([])

#Sort Dict by # of samples
sorted_tuples = sorted(wo_dict.items(),key=operator.itemgetter(1), reverse=True)

#Load list of wo's
for item in sorted_tuples:
    usable_list.append(item[0])

#first fit decreasing method
if len(usable_list) > 0:
    ffd(usable_list,plates)

#Load Outgoing plates based on sol found by FFD
for plt in plates:
    WGS_out_plates.append(plate())
    WGS_out_plates[len(WGS_out_plates) - 1].pipe = 'WGS'
    for order in plt:
        for plte in WGS_sortable_plates:
            if order in plte.wo:
                WGS_out_plates[len(WGS_out_plates) - 1].add_wo(order)
                for samp in plte.samples:
                    WGS_out_plates[len(WGS_out_plates) - 1].add_sample(samp)


"""
Choose plates to go out or keep
---------------------------------------------------
"""


hold_list = []
small_plates = []
for plt in WGS_out_plates:
    if len(plt.samples) < 60:
        small_plates.append(plt)

for plt in Exome_out_plates:
    if len(plt.samples) < 60:
        small_plates.append(plt)

if len(small_plates) == 0:
    print('All plates over 60.')
else:
    print('-' * 50)
    print('{:>22}{:>39}'.format('Work Order(s)','# of Samples'))
    count = 1
    for plt in small_plates:
        print('Plate {}: {:<39} {}'.format(count,' '.join(plt.wo).replace('.0',''),len(plt.samples)))
        count += 1

    while True:
        plt_in = input('Enter the plate numbers you want to hold.\n(Enter 0 to exit or if you wish to run all): ')

        if is_int(plt_in) and 0 < int(plt_in) < count and not plt_in == '':
            hold_list.append(small_plates[int(plt_in) - 1])
        elif int(plt_in) == 0:
            break
        else:
            print('Enter a number between 0 and {}'.format(count))

for plt in hold_list:
    if plt in WGS_out_plates:
        WGS_out_plates.remove(plt)
    elif plt in Exome_out_plates:
        Exome_out_plates.remove(plt)

"""
Build output files
----------------------------------------------------------
"""

plate_files = []
plate_count = 1
FFPE_list = []

for plt in WGS_out_plates:
    plate_file = '{}_{}_Frag_Plate_{}_{}_{}.csv'.format('_'.join(plt.wo).replace('.0',''),len(plt.samples),plate_count,plt.pipe,mmddyy)
    plate_count += 1
    with open(plate_file, 'w') as of:
        plate_writer = csv.DictWriter(of,delimiter = ',', fieldnames=header)
        plate_writer.writeheader()
        plate_files.append(plate_file)
        for wo in plt.wo:
            for file in wo_frag_files:
                if wo.replace('.0','') in file:
                    with open(file, 'r') as ff:
                        frag_reader = csv.DictReader(ff,delimiter = ',')
                        for line in frag_reader:
                            for samp in plt.samples:
                                if line['Barcode'] == samp.name:
                                    plate_writer.writerow(line)
                                    if samp in FFPE_samples:
                                        if plate_file not in FFPE_list:
                                            FFPE_list.append(plate_file)


for plt in Exome_out_plates:
    plate_file = '{}_{}_Frag_Plate_{}_{}_{}.csv'.format('_'.join(plt.wo).replace('.0',''),len(plt.samples),plate_count,plt.pipe,mmddyy)
    plate_count += 1
    with open(plate_file, 'w') as of:
        plate_writer = csv.DictWriter(of,delimiter = ',', fieldnames=header)
        plate_writer.writeheader()
        plate_files.append(plate_file)
        for wo in plt.wo:
            for file in wo_frag_files:
                if wo.replace('.0','') in file:
                    with open(file, 'r') as ff:
                        frag_reader = csv.DictReader(ff,delimiter = ',')
                        for line in frag_reader:
                            for samp in plt.samples:
                                if line['Barcode'] == samp.name:
                                    plate_writer.writerow(line)
                                    if samp in FFPE_samples:
                                        if plate_file not in FFPE_list:
                                            FFPE_list.append(plate_file)


for file in FFPE_list:
    plate_files.remove(file)
    os.rename(file, file.replace('Frag_Plate','Frag_Plate_FFPE'))
    plate_files.append(file.replace('Frag_Plate','Frag_Plate_FFPE'))

"""
Build Held plate files
-----------------------------------------------------------
"""

FFPE_list = []

for plt in hold_list:
    for order in plt.wo:
        samp_count = 0
        for file in wo_frag_files:
            if order.replace('.0','') in file:
                with open(file, 'r') as ff:
                    frag_reader = csv.DictReader(ff, delimiter = ',')
                    filename = 'hold_'+file
                    with open(filename, 'w') as off:
                        frag_writer = csv.DictWriter(off, delimiter = ',', fieldnames=header)
                        frag_writer.writeheader()
                        for line in frag_reader:
                            for samp in plt.samples:
                                if line['Barcode'] == samp.name:
                                    frag_writer.writerow(line)
                                    samp_count += 1
                                    if samp in FFPE_samples:
                                        if plate_file not in FFPE_list:
                                            FFPE_list.append(plate_file)
        new_filename = filename.split('_')
        new_filename[2] = str(samp_count)
        new_filename = '_'.join(new_filename)
        os.rename(filename, new_filename)


for file in FFPE_list:
    os.rename(file, file.replace('Frag_Plate','FFPE_Frag_Plate'))

"""
Push to SmartSheet
------------------------------------------------------------
"""

workspaces = get_workspace_list()

for workspace in workspaces:
    if workspace.name == 'Production Pipeline':
        prod_workspace = workspace

prod_folders = get_folder_list(prod_workspace.id, 'w')

for folder in prod_folders:
    if folder.name == 'OPG':
        OPG_folder = folder

OPG_folders = get_folder_list(OPG_folder.id, 'f')

for folder in OPG_folders:
    if folder.name == 'lib_core':
        lib_core_folder = folder

lib_core_folder = get_object(lib_core_folder.id, 'f')

for sheet in lib_core_folder.sheets:
    if sheet.name == 'plate_assignment_sheet':
        assgn_sheet = sheet

assgn_sheet = get_object(assgn_sheet.id, 's')
num_rows = len(assgn_sheet.rows) + 1

assgn_sheet_col_ids = {}

for col in assgn_sheet.columns:
    assgn_sheet_col_ids[col.title] = col.id

for file in plate_files:
    imported_sheet = smart_sheet_client.Folders.import_csv_sheet(
      lib_core_folder.id,           # folder_id
      file,
      file,  # sheet_name
      header_row_index=0
    ).data

    imported_sheet = get_object(imported_sheet.id, 's')

    # get FFPE tag
    if 'FFPE' in file:
        FFPE = True
    else:
        FFPE = False

    new_row = smartsheet.smartsheet.models.Row()
    new_row.to_bottom = True



    new_row.cells.append({'column_id': assgn_sheet_col_ids['Plate File Name'], 'value': file})
    new_row.cells.append({'column_id': assgn_sheet_col_ids['Link to Plate Sheet'], 'value': file, 'hyperlink' : {"sheetId" : imported_sheet.id}})
    new_row.cells.append({'column_id': assgn_sheet_col_ids['No. of Samples'], 'value': file.split('_')[1]})
    new_row.cells.append({'column_id': assgn_sheet_col_ids['FFPE Flag'], 'value': FFPE})
    new_row.cells.append({'column_id': assgn_sheet_col_ids['Task'], 'formula': '=IF([Fragmentation Complete]{} = 1, IF([Lib Construction Complete]{} = 1, IF([QC Complete]{} = 1, "Complete", "QC"), "Lib Construction"), "Fragmentation")'.format(num_rows,num_rows,num_rows)})

    response = smart_sheet_client.Sheets.add_rows(assgn_sheet.id, [new_row])
    response = response.data

    """
    need row and column id's to create reference ---
    uses the same row/column as start/finish
    append column to import sheet
    update rows with the new cross sheet reference
    """

    xref = smartsheet.models.CrossSheetReference({
    'name': 'plt_assgn_ref',
    'source_sheet_id': assgn_sheet.id,
    'start_row_id': response[0].id,
    'end_row_id': response[0].id,
    'start_column_id': assgn_sheet_col_ids['Task'],
    'end_column_id': assgn_sheet_col_ids['Task']
    })

    smart_sheet_client.Sheets.create_cross_sheet_reference(
        imported_sheet.id, xref)

    new_column = smartsheet.smartsheet.models.Column({
        'title': 'Status',
        'type': 'TEXT_NUMBER',
        'index': len(imported_sheet.columns) + 1
    })

    col_response = smart_sheet_client.Sheets.add_columns(imported_sheet.id,[new_column]).data[0]

    import_rows = []
    for row in imported_sheet.rows:

        up_row = smartsheet.smartsheet.models.Row()
        up_row.id = row.id
        for cel in row.cells:
            up_row.cells.append(cel)

        new_cell = smartsheet.smartsheet.models.Cell()
        new_cell.column_id = col_response.id
        new_cell.formula = '={plt_assgn_ref}'

        up_row.cells.append(new_cell)
        import_rows.append(up_row)

    smart_sheet_client.Sheets.update_rows(imported_sheet.id, import_rows)

    num_rows += 1



"""
Move Files to correct places
------------------------------------------------------------
"""


if not os.path.exists('processed/frag_temps/'):
    os.makedirs('processed/frag_temps/')

for file in wo_frag_files:
    os.rename(file, 'processed/frag_temps/' + file)

hold_list = glob.glob('hold*')

for file in hold_list:
    os.rename(file, file.replace('hold_',''))

