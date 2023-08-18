# vcf-parser<!-- omit in toc -->

Python program for parsing VCF files and generating an Excel spreadsheet with contact data.

## Table of contents

- [Table of contents](#table-of-contents)
- [1. Description](#1-description)
- [2. Getting started](#2-getting-started)
  - [2.1 Dependencies](#21-dependencies)
  - [2.2 Installing](#22-installing)
  - [2.3 Executing program](#23-executing-program)
- [3. Version history](#3-version-history)

<!-- toc -->

## 1. Description

Python project consisting in parsing VCF files (typically used in _BusyContacts_ macOS app for instance) and generating a summary Excel spreadsheet with contact data under the form of a pivot table. Contacts can be filtered according to their category tag(s) attributed in the original contact management app. The program currently takes 3 inputs: The path of the VCF file to parse (by default, the latest backup of _BusyContacts_ macOS app is taken into account), the filtering tags for characterizing the contacts to filter and the logical operator being either "&" (in case one wants to filter out all contacts precisely presenting all the filtering tags) or "|" (for filtering out all contacts presenting at least one of the filtering tags). Note that if no filtering tag is provided at all, a spreadsheet containing all contacts stored in the VCF file is generated, no matter the tag(s) of the contacts.

## 2. Getting started

### 2.1 Dependencies

- Tested on macOS Ventura version 13.4
- Python 3.10.0

### 2.2 Installing

`pip install -r requirements.txt`

### 2.3 Executing program

- To access useful help messages, type following Terminal command at the root of
  the project:
  
  `python3 src/main.py -h`

- To run the script, simply type for instance following command from the root of the project:

  `python3 src/main.py -vcf_file_path "absolute/path/to/Contacts.vcf" -tag_list my_tag_1 my_tag_2 my_tag_3 -logic_op "|"`

## 3. Version history

- 0.1
  - Initial release
