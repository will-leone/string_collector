"""
VERSIONS

[V1] Will Leone 2 DEC 2018:
    - Purpose: Locate all occurrences of a given list of strings in any file
      within a chosen directory.
    - Details:
      - Outputs a multi-sheet Excel workbook.
         - Each sheet is named after a string in the provided list.
         - A given string's sheet contains all instances where the string
           occurred (including file, line #, and the line itself).
         - For SAS files, only full variable definition statements (even those
           spanning >1 line) are included. (This is intentional.)
      - Code is intended to work on very large files (>1 GB), in part to
        encourage code re-use for other projects
      - Supported file extensions: any beginning with .sas, .sql, .py, .csv,
        and .txt.
      - If searching for a variable in a SAS file, extract all statements
        defining the variable and provide the lines spanned by each statement

[V2] TENTATIVE: will add support for PDF and MS Office documents.

INSTRUCTIONS
Open Anaconda Prompt and type the following. Press Enter
after each step.
- activate saspy_env
      # replace saspy_env with a Python virtual environment
- cd /d X:\Data_Governance\Projects\PDW Data Dictionary\Code
      # Change directory to wherever this code is located
- python
- import string_collector
- X:\Data_Governance\Projects\PDW Data Dictionary\Code
      # Folder to be searched (when prompted)
- sas_BUFFER   ~ thisvaluedoesn'texist ~ Non_null
      # Strings to search for (when prompted)

"""

import os
import numpy as np
import pandas as pd  # implicit xlsxwriter dependency
import datetime


def dir_search():

    start = datetime.datetime.utcnow()

    print(("Copy the directory path in which you want to search. (CTRL+SHIFT "
           + "and right-click on the directory, then Copy as Path). Paste and "
           + "press enter."))

    _mypath = input()
    mypath = _mypath.strip(" \"\'")
    print('Search directory is: ', mypath)

    print(('Enter the string you want to search for in ', mypath, '. You can '
           + 'search for multiple strings by separating them with ~ (tilde).'))
    _list_target_strings = input()
    list_target_strings = _list_target_strings.split('~')
    for k, object in enumerate(list_target_strings):
        list_target_strings[k] = object.strip()
    timestamp = start.strftime('%d.%m.%Y')
    results_filename = (('scan_' + ','.join(list_target_strings).upper()
        + '_' +  timestamp + '.xlsx'))
    try:  # could also use if os.is_file() to check
        os.remove(results_filename)  # removes prior version if file exists
    except FileNotFoundError:
        pass

    def str_scan(target_dir, file, my_str):
        with open(os.path.join(target_dir,file), 'r') as f:
            occurrences = dict()
            if '.sas' in file:
                sas_buffer = str()
            else:  # not necessary, but helps readability and encourages review of logic flow
                pass
            try:
                for linenum, line in enumerate(f, start=1):
                    if '.sas' in file:
                        sas_buffer += line.strip()
                    else:
                        pass
                    if my_str.lower() in line.lower():
                        if (file, my_str, linenum) not in occurrences:
                            occurrences[(file, my_str, linenum)] = line.strip()
                        else:
                            pass
                        if '.sas' in file and sas_buffer:  # duck typing
                            if '=' in sas_buffer:
                                occurrences[(file, my_str, linenum)] = sas_buffer
                            else:
                                del occurrences[(file, my_str, linenum)]  # false positive
                    else:
                        continue
                    if '.sas' in file and ';' in sas_buffer:  # second condition only checked if first is true
                        sas_buffer = str()
            except UnicodeDecodeError:
                pass
        occurrences_list = list()
        for key, value in occurrences.items():
            occurrences_list.append((key[0], key[1], key[2], value))
        print('\n\nNew occurrences list:')  # QA
        for item in occurrences_list:
            print(item)
        print('\n\n')
        return occurrences_list

    data_box = list()
    for target_string in list_target_strings:
        print(f"\nBEGINNING SCAN FOR {target_string}.\n")
        dir_files = list()
        for foo in os.listdir(mypath):
            if any(extension in foo for extension in
                    ['.sas', '.sql', '.py', '.csv', '.txt']):
                new_scan = str_scan(mypath, foo, target_string)
                if new_scan:  # duck typing: executes iff new_scan is not empty
                    dir_files.extend(new_scan)
                print(f"Scanned {foo} for instances of {target_string}.")
            else:
                continue

        final_array = np.array(dir_files)
        df = pd.DataFrame(final_array, columns=['file', 'string'
                                                , 'linenum', 'line'])
        data_box.append((target_string, df))

    destination = "C:\\Users\\" + os.getlogin() + "\\Downloads"
    os.chdir(destination)

    with pd.ExcelWriter(results_filename, engine='xlsxwriter') as writer:
        for (string_foo, df_foo) in data_box:
            df_foo.to_excel(writer, sheet_name=string_foo, index=False)
        writer.save()

    print("\nThe results have been saved to your Downloads folder at "
          , f"{destination}.\n")

    end = datetime.datetime.utcnow()
    print("\nThis script ran in ", (end - start).total_seconds(), " seconds.\n")

    return

dir_search()
raise SystemExit  # properly exit Python once the script executes
