#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Script name:         main.py
# Description:         Code for reading VCF files
# Invocation example:  Reading VCF file and filtering contacts based on tags:
#                           python3 src/main.py -vcf_file_path "/Users/anthony/MEGA/DOCUMENTS/Programmation/Python/MyPythonProjects/VCFParser/tests/ContactsTest.vcf" -tag_list 1m02 2m02 3m02 -logic_op "|"
#                      Accessing useful help messages:
#                           python3 src/main.py -h
# Author:              Anthony Guinchard
# Version:             0.1
# Creation date:       2023-06-29
# Modification date:   2023-08-13
# Working:             ✅


# Required packages

import argparse
import base64
import glob
import json
import logging
import math as m
import os
import re
import sys
from datetime import datetime

import openpyxl
import pandas as pd
from openpyxl.drawing.image import \
    Image  # requires Pillow to be installed to fetch image objects
from tqdm import tqdm

# Setting current working directory

# Getting the path leading to the current working directory
project_path = os.getcwd()
print(f"project_path: {project_path}")  # printing "project_path"
# Getting the path leading to the currently executing script
script_path = sys.path[0]
print(f"script_path: {script_path}")    # printing "script_path"
# os.chdir(project_path)                # setting the current working directory based on the path leading to the current working directory
# setting the current working directory based on the path leading to the
# currently executing script
os.chdir(script_path)


# Initializations

DEBUG_MODE = True
BUSY_CONTACTS_BACKUP_PATH = '/Users/anthony/Library/CloudStorage/Dropbox/Applications/BusyContactsBackups/'
JSON_FILE_PATH = script_path+'/json'
OUTPUT_FILE_PATH = script_path+'/out'
PICTURE_FILE_PATH = script_path+'/pictures_temp'
JSON_FILE_NAME = '/contacts_dict.json'
EXCEL_WORKBOOK_NAME = '/contacts.xlsx'
EXCEL_WORKSHEET_NAME = 'Sheet1'

# Conversion table of remainders to "Excel base 26" equivalent
conversion_table = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E',
                    6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J',
                    11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O',
                    16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T',
                    21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y',
                    26: 'Z'}


# Functions

def get_timestamp():
    """
    Get current timestamp.
    """
    now = datetime.now()
    return now.strftime("%Y-%m-%d_%H-%M-%S")


def print_and_log(message):
    """
    Print and log debug message
    """
    print(f'{get_timestamp()}: {message}')
    logging.info(message)


def get_latest_busy_contacts_vcf():
    """
    List all ".babu" files in BUSY_CONTACTS_BACKUP_PATH.
    """
    print_and_log('Getting latest VCF backup file of BusyContacts app.')
    babu_path_list = glob.glob(BUSY_CONTACTS_BACKUP_PATH + "*.babu")
    babu_path_list_sorted = sorted(babu_path_list)
    babu_path_latest = babu_path_list_sorted[-1]+'/'
    babu_path_latest_content = os.listdir(babu_path_latest)
    for item in babu_path_latest_content:
        item_path = babu_path_latest+item+'/'
        if os.path.isdir(item_path):
            backup_dir_path = item_path
            break
    latest_busy_contacts_vcf = backup_dir_path+'Contacts.vcf'

    return latest_busy_contacts_vcf


def parse_arguments(latest_busy_contacts_vcf):
    """
    Read arguments from a command line.
    """

    print_and_log('Parsing arguments.')

    parser = argparse.ArgumentParser(description="This program intends to parse VCF file\
    content.\
    Enjoy!")

    if DEBUG_MODE:

        parser.add_argument(
            '-vcf_file_path',
            metavar='/path/to/your/vcf/file/vcf_file.vcf',
            type=str,
            default='/Users/anthony/MEGA/DOCUMENTS/Programmation/Python/MyPythonProjects/VCFParser/tests/ContactsTest.vcf',
            help='VCF file path to read content from.'
        )
        parser.add_argument(
            '-tag_list',
            metavar='my_tag_1 my_tag_2 my_tag_3',
            nargs='+',
            type=str,
            default=['EPFL'],  # ['1m02', '2m02', '3m02']
            help='List of tags identifying contacts of interest. If left empty, this means that all contacts have to be extracted from VCF file independently of their tags.'
        )
        parser.add_argument(
            '-logic_op',
            metavar='&',
            type=str,
            default='|',
            help='Logical operator symbol being either "&" or "|" indicating whether the user wants to extract contacts meeting respectively all tags or only at least one from the list of tags.'
        )

    else:

        parser.add_argument(
            '-vcf_file_path',
            metavar='/path/to/your/vcf/file/vcf_file.vcf',
            type=str,
            default=latest_busy_contacts_vcf,
            help='VCF file path to read content from.'
        )
        parser.add_argument(
            '-tag_list',
            metavar='my_tag_1 my_tag_2 my_tag_3',
            nargs='+',
            type=str,
            default=[],  # empty list of tags
            help='List of tags identifying contacts of interest. If left empty, this means that all contacts have to be extracted from VCF file independently of their tags.'
        )
        parser.add_argument(
            '-logic_op',
            metavar='&',
            type=str,
            default='|',
            help='Logical operator symbol being either "&" or "|" indicating whether the user wants to extract contacts meeting respectively all tags or only at least one from the list of tags.'
        )

    args = parser.parse_args()

    return args


def initialize_logging():
    """
    Initialize log file.
    """
    logging.basicConfig(
        filename=f'{script_path}/app.log',
        filemode='w',
        encoding='utf-8',
        format='%(asctime)s: %(message)s',
        datefmt='%y-%m-%d_%H-%M-%S',
        level=logging.INFO
    )


def reset_field_counts(field_dict_list):
    """
    Reset field count.
    """
    for field_dict in field_dict_list:
        for key in list(field_dict.keys()):
            field_dict[key]['count'] = 0


def add_contact(contacts_dict, contact_dict, contact_id):
    """
    Add contact dictionary in nested global contacts dictionary.
    """
    contacts_dict[contact_id] = contact_dict
    contact_id = contact_id+1
    listening_to_data = False
    return contacts_dict, contact_dict, contact_id, listening_to_data


def decimalToExcelBase26(decimal):
    """
    Convert decimal value to "Excel base 26" value.
    """
    excel_base_26 = ''
    while (decimal > 0):
        remainder = decimal % 26
        excel_base_26 = conversion_table[remainder] + excel_base_26
        decimal = decimal // 26

    return excel_base_26


def main(args):
    """
    Main function.
    """

    # Printing arguments
    print(f'args.vcf_file_path: {args.vcf_file_path}')
    print(f'args.tag_list: {args.tag_list}')
    print(f'args.logic_op: {args.logic_op}')

    # Extracting contact data between "FN" and "REV" fields

    contacts_dict = {}
    contact_id = 0

    BEGIN_SYMBOL = 'N:'
    END_SYMBOL = 'END:'

    standard_field_file = open(JSON_FILE_PATH+'/standard_fields.json')
    standard_field_dict = json.load(standard_field_file)
    standard_field_file.close()
    standard_field_list = list(standard_field_dict.keys())

    custom_field_file = open(JSON_FILE_PATH+'/custom_fields.json')
    custom_field_dict = json.load(custom_field_file)
    custom_field_file.close()
    custom_field_list = list(custom_field_dict.keys())

    social_profile_field_file = open(
        JSON_FILE_PATH+'/social_profile_fields.json')
    social_profile_field_dict = json.load(social_profile_field_file)
    social_profile_field_file.close()
    social_profile_field_list = list(social_profile_field_dict.keys())

    field_dict_list = [standard_field_dict,
                       custom_field_dict, social_profile_field_dict]

    listening_to_data = False

    key = BEGIN_SYMBOL
    key_previous = BEGIN_SYMBOL

    value = 'Name'
    value_previous = 'Name'
    last_symbol = ''

    print_and_log('Reading VCF file...')

    with open(args.vcf_file_path, mode='r') as vcf_:
        for line in tqdm(vcf_):

            if line.startswith(BEGIN_SYMBOL):
                # Listening to data and creating a new entry in contact_dict
                listening_to_data = True
                contact_dict = {}
                key = BEGIN_SYMBOL
                key_previous = BEGIN_SYMBOL
                reset_field_counts(field_dict_list)

            if line.startswith(END_SYMBOL) and listening_to_data:
                # Checking list of tags at the very end of current contact
                if (len(args.tag_list) > 0):
                    # If we want to filter contacts based on tags and if the
                    # current contact has no "Tags" field, then we skip this contact...
                    if 'Tags' not in contact_dict:
                        listening_to_data = False
                        continue
                    else:
                        # Checking tags (for filtering contacts, i.e. skipping some
                        # contacts if tags are not met) ONLY when list of tags has
                        # been completely gathered for current contact
                        # Note: If no tag is listed in "args.tag_list", then no
                        # contact at all gets filtered out and an Excel spreadsheet
                        # with all contacts stored in VCF file is generated.
                        tag_value = standard_field_dict['CATEGORIES']['value']
                        contact_tags = contact_dict[tag_value].split(',')
                        if args.logic_op == '&':
                            if not all(tag in contact_tags for tag in args.tag_list):
                                listening_to_data = False
                                continue
                            else:
                                contacts_dict, contact_dict, contact_id, listening_to_data = add_contact(
                                    contacts_dict, contact_dict, contact_id)
                                continue
                        if args.logic_op == '|':
                            if not any(tag in contact_tags for tag in args.tag_list):
                                listening_to_data = False
                                continue
                            else:
                                contacts_dict, contact_dict, contact_id, listening_to_data = add_contact(
                                    contacts_dict, contact_dict, contact_id)
                                continue
                else:
                    contacts_dict, contact_dict, contact_id, listening_to_data = add_contact(
                        contacts_dict, contact_dict, contact_id)
                    continue

            if listening_to_data:

                # Contact name
                if line.startswith(BEGIN_SYMBOL):
                    names = line.split(':')[-1]
                    last_name = names.split(';')[0]
                    first_name = names.split(';')[1]
                    middle_name_1 = names.split(';')[2]
                    middle_name_2 = names.split(';')[3]
                    if len(last_name) > 0:
                        contact_dict['Last Name'] = last_name
                    if len(first_name) > 0:
                        contact_dict['First Name'] = first_name
                    if len(middle_name_1) > 0:
                        contact_dict['Middle Name 1'] = middle_name_1
                    if len(middle_name_2) > 0:
                        contact_dict['Middle Name 2'] = middle_name_2

                # Built-in fields
                if line.startswith(tuple(standard_field_list)):

                    # Getting the key, its value, saving the data at this line in contact, and then moving to next line
                    key = re.split(':|;', line)[0]
                    value = standard_field_dict[key]['value']
                    standard_field_dict[key]['count'] = standard_field_dict[key]['count']+1
                    count = standard_field_dict[key]['count']
                    if count > 1:
                        if count == 2:
                            contact_dict[value +
                                         '_1'] = contact_dict.pop(value)
                        value = value+'_'+str(count)
                    if key == 'NOTE':
                        data = line.split('NOTE:')[-1]
                    else:
                        data = line.split(':')[-1]
                    if data.startswith(';;'):  # this is sometimes the case for ADR
                        data = data[2:]
                    # Replacing unwanted characters in case field is not a "Note"
                    if key != 'NOTE' and line[0] != ' ':
                        data = data.replace('\n', '').replace(
                            '\\', '').replace(';;', ', ').replace(';', ', ')
                    if key == 'ADR' and data.startswith(', '):
                        data = data[2:]
                    if key == 'NOTE':
                        # TODO: Refactor with below same block already used!
                        # Removing ending line break finishing line of length at least 72
                        if (len(data) - 72) <= 0 and data[-1:] == '\n':
                            data = data[:-1]
                        # Getting last symbol
                        last_symbol = data[-1:]  # last_symbol = data[-1]
                        # Translating elements
                        data = data.replace('\xa0', '')
                        data = data.replace('\u200b', '')
                        # data = data.replace('\\\n', '\n')
                        data = data.replace('\\n', '\n')
                        data = data.replace('\\', '')
                    if key == 'PHOTO':
                        data = data.split(',')[-1]
                    if key == 'BDAY':
                        if data.startswith('--'):
                            data = data[2:4]+'-'+data[-2:]
                        else:
                            data = data[:-4]+'-'+data[-4:-2]+'-'+data[-2:]
                    contact_dict[value] = data
                    key_previous = key
                    value_previous = value

                # Custom fields
                if 'X-CUSTOM' in line and any(custom_field in line for custom_field in custom_field_list):
                    for custom_field in custom_field_list:
                        if line.find(custom_field) >= 0:
                            key = custom_field
                            break
                    value = custom_field_dict[key]['value']
                    custom_field_dict[key]['count'] = custom_field_dict[key]['count']+1
                    count = custom_field_dict[key]['count']
                    if count > 1:
                        if count == 2:
                            contact_dict[value +
                                         '_1'] = contact_dict.pop(value)
                        value = value+'_'+str(count)
                    data = line.split('-=+=-')[-1].replace('\n', '')
                    contact_dict[value] = data
                    key_previous = key
                    value_previous = value

                # Social profile fields
                if 'X-SOCIALPROFILE' in line and any(social_profile_field in line for social_profile_field in social_profile_field_list):
                    key = line.split(';')[1].split(':')[0].replace('TYPE=', '')
                    value = social_profile_field_dict[key]['value']
                    social_profile_field_dict[key]['count'] = social_profile_field_dict[key]['count']+1
                    count = social_profile_field_dict[key]['count']
                    if count > 1:
                        if count == 2:
                            contact_dict[value +
                                         '_1'] = contact_dict.pop(value)
                        value = value+'_'+str(count)
                    pref_to_remove = line.split(';')[-1].split(':')[0]+':'
                    data = line.split(
                        ';')[-1].replace(pref_to_remove, '').replace('\n', '')
                    contact_dict[value] = data
                    key_previous = key
                    value_previous = value

                # Unfinished lines (this seems to be only possible for Address,
                # Note, Picture, Tags and social profile fields)
                if (key_previous == 'ADR' or key_previous == 'NOTE' or
                    key_previous == 'PHOTO' or key_previous == 'CATEGORIES' or
                        key_previous in social_profile_field_list) and line[0] == ' ':
                    if key_previous == 'NOTE':
                        # Getting rid of first space at beginning of line and replacing:
                        line = line[1:]
                        # Removing eventual ending line break finishing line of length up to 75
                        if (len(line) - 75) <= 0 and line[-1:] == '\n':
                            line = line[:-1]
                        # Recreating original line break if cut in half in VCF file
                        first_symbol = line[0:2]
                        if last_symbol == '\\' and first_symbol == 'n ':
                            line = '\n'+line[1:]
                        # Getting last symbol
                        last_symbol = line[-1:]  # last_symbol = line[-1]
                        # Translating elements
                        line = line.replace('\xa0', '')
                        line = line.replace('\u200b', '')
                        # line = line.replace('\\\n', '\n')
                        line = line.replace('\\n', '\n')
                        line = line.replace('\\', '')
                    if key_previous == 'ADR' or key_previous == 'PHOTO' or key_previous == 'CATEGORIES' or key_previous in social_profile_field_list:
                        # Getting rid of first space at beginning of line and replacing "\n" with ""
                        line = line[1:].replace('\n', '')
                    # Appending unfinished line to previous line(s)
                    contact_dict[value_previous] = contact_dict[value_previous]+line

                # Checking tags (for filtering contacts, i.e. skipping some
                # contacts if tags are not met) ONLY when list of tags has
                # been completely gathered for current contact
                # Note: If no tag is listed in "args.tag_list", then no
                # contact at all gets filtered out and an Excel spreadsheet
                # with all contacts stored in VCF file is generated.
                # if not listening_to_tags and (len(args.tag_list) > 0):
                #     tag_value = standard_field_dict['CATEGORIES']['value']
                #     contact_tags = contact_dict[tag_value].split(',')
                #     if args.logic_op == '&':
                #         if not all(tag in contact_tags for tag in args.tag_list):
                #             listening_to_data = False
                #             continue
                #     if args.logic_op == '|':
                #         if not any(tag in contact_tags for tag in args.tag_list):
                #             listening_to_data = False
                #             continue

    # Converting contacts_dict dictionary to contacts_df DataFrame
    print_and_log('Converting dictionary to DataFrame.')
    contacts_df = pd.DataFrame.from_dict(contacts_dict, orient="index")

    # Reordering DataFrame rows in chronological index order
    print_and_log('Reordering DataFrame\'s rows.')
    contacts_df.sort_index(axis=0, inplace=True)

    # Reordering DataFrame columns
    print_and_log('Reordering DataFrame\'s columns.')

    column_list = contacts_df.columns.tolist()
    ordered_column_list = []

    # Ordering the first columns
    first_columns_list = ['First Name', 'Middle Name 1', 'Middle Name 2', 'Last Name', 'Nickname', 'Organization', 'Profession',
                          'Birthday', 'Gender', 'Nationality', 'Tags', 'Note', 'Picture']
    for col in first_columns_list:
        if col in column_list:
            ordered_column_list.append(col)

    # Ordering the last columns
    ordered_column_set = set(ordered_column_list)
    last_columns_diff = sorted(
        [x for x in column_list if x not in ordered_column_set])

    # Merging duplicated last columns (i.e., 'Address' into 'Address_1', 'Email' into 'Email_1', etc.)
    last_columns_diff_no_dupl = []
    i = 0
    while i < (len(last_columns_diff)-1):
        col = last_columns_diff[i]
        col_next = last_columns_diff[i+1]
        if col_next.startswith(col):
            # Merging col and col_next
            df_col = contacts_df[col]
            df_col_next = contacts_df[col_next]
            df_merge = df_col.combine_first(df_col_next)
            # Inserting df_merge in col_next
            contacts_df[col_next] = df_merge
            # Removing col
            contacts_df.drop(col, axis=1, inplace=True)
            # Updating last_columns_diff_no_dupl list
            last_columns_diff_no_dupl.append(col_next)
            i = i+2
            while (i < len(last_columns_diff)) and (col in last_columns_diff[i]):
                last_columns_diff_no_dupl.append(last_columns_diff[i])
                i = i+1
        else:
            last_columns_diff_no_dupl.append(col)
            i = i+1

    # Applying new column order
    ordered_column_list = ordered_column_list + last_columns_diff_no_dupl
    contacts_df = contacts_df[ordered_column_list]

    # Saving transposed DataFrame to JSON file
    print_and_log('Saving DataFrame to pretty-printed JSON file.')
    with open(OUTPUT_FILE_PATH+JSON_FILE_NAME, 'w', encoding='utf-8') as file:
        contacts_df.T.to_json(file, force_ascii=False, indent=2)

    # Creating a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(
        OUTPUT_FILE_PATH+EXCEL_WORKBOOK_NAME, engine="xlsxwriter")

    # Saving contacts_df to contacts.xlsx
    print_and_log('Saving DataFrame to Excel spreadsheet.')
    contacts_df.to_excel(writer, sheet_name=EXCEL_WORKSHEET_NAME, index=True)

    # Setting up Excel worksheet pivot table
    print_and_log('Setting up Excel worksheet pivot table.')

    # Getting XlsxWriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets[EXCEL_WORKSHEET_NAME]
    # Getting the dimensions of the DataFrame
    (max_row, max_col) = contacts_df.shape
    # Creating a list of column headers, to use in "add_table()"
    column_settings = [{"header": "Index"}]+[{"header": column}
                                             for column in contacts_df.columns]
    # Adding the Excel table structure (Pandas will add the data)
    if max_row > 0:
        worksheet.add_table(0, 0, max_row, max_col,
                            {"columns": column_settings})
    else:
        logging.warning(
            'WARNING - "contact_df" is empty. Must have at least one data row in in "add_table()"')
    # Closing the Pandas Excel writer and output the Excel file
    writer.close()

    # Handling pictures
    print_and_log('Handling contact pictures.')

    photo_value = standard_field_dict['PHOTO']['value']
    if photo_value in contacts_df:

        # Creating pictures
        for i, pic_b64encode in tqdm(enumerate(contacts_df[photo_value])):
            if isinstance(pic_b64encode, str):
                # Decoding picture
                pic_b64decode = base64.b64decode(pic_b64encode)
                # Creating writable image
                pic_png = open(f'{PICTURE_FILE_PATH}/{i}.png', 'wb')
                # Writing image
                pic_png.write(pic_b64decode)

        # Inserting pictures in Excel spreadsheet
        # Opening existing workbook
        workbook = openpyxl.load_workbook(OUTPUT_FILE_PATH+EXCEL_WORKBOOK_NAME)
        # Opening existing worksheet
        worksheet = workbook[EXCEL_WORKSHEET_NAME]
        # Getting column index in workbook
        photo_col_idx = ordered_column_list.index(photo_value)+2
        photo_col_idx_excel = decimalToExcelBase26(photo_col_idx)
        # Iterating over image paths
        for pic_path in glob.glob(f'{PICTURE_FILE_PATH}/*.png'):
            pic_idx = int(pic_path.split('/')[-1].replace('.png', ''))
            pic = Image(pic_path)
            pic.width = 20
            pic.height = 20
            cell_address = photo_col_idx_excel+str(pic_idx+2)
            worksheet[cell_address] = ''  # erasing current cell content
            worksheet.add_image(pic, cell_address)
        # Saving workbook
        workbook.save(OUTPUT_FILE_PATH+EXCEL_WORKBOOK_NAME)

        # Deleting picture files once workbook has been saved
        for pic_path in glob.glob(f'{PICTURE_FILE_PATH}/*.png'):
            os.remove(pic_path)


# Main program
if __name__ == '__main__':
    initialize_logging()
    latest_busy_contacts_vcf = get_latest_busy_contacts_vcf()
    args = parse_arguments(latest_busy_contacts_vcf)
    main(args)


# TODO:
# - Clean and refactor wherever possible
# - Test with total dataset to extracct 1m02, etc.
# - Generate GIF and nice cover picture
# - Update GitHub
