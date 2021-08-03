""" Run with at least Python 3.8, pip install xlsxwriter if needed """

import os
import pandas as pd

# Specify path to .csv Reports
path = "C:\\path\\to\\folder\\containing\\tufin_reports\\"

# Iterate over .csv files in a path
files = [x for x in os.listdir(path=path) if x.endswith(".csv")]

# Adjust encoding if needed
encoding_files = "windows-1252"
number = 0

dataframe_list = list()

while (number := number + 1) < len(files):
    for f in files:
        file = path + f

        # Reading files
        main_frame = pd.read_csv(file, encoding=encoding_files)

        # Getting a row number of column set for rules
        rules = main_frame.index[main_frame["Tufin Object lookup results"] == "Device name"].tolist()

        # Changing main columns to the identified column for rules
        main_frame.columns = main_frame.iloc[rules[0]]

        # Dropping each row until founded columns + 1 - it's self, what leaves me with a new assigned columns & rules
        # only
        main_frame = main_frame.drop(main_frame.index[0:rules[0] + 1])

        for index, row in main_frame.iterrows():
            # Appending to list for further creation of a DataFrame + Lowering the case for consistent checks
            dataframe_list.append(row.str.lower())

# Creating a new DataFrame from a list
new_frame = pd.DataFrame(dataframe_list)

# Removing duplicate Rules from a DataFrame
new_frame = new_frame.drop_duplicates(keep='first')

# Writing to all rules to "Rules" sheet
writer = pd.ExcelWriter(path + "Parsed_Rules.xlsx", engine='xlsxwriter')
new_frame.to_excel(writer, sheet_name="All Rules", startrow=0)

# Color Definition
colors = {
    'PASS': '\033[92m',     # GREEN
    'WARNING': '\033[93m',  # YELLOW
    'FAIL': '\033[91m',     # RED
    'RESET': '\033[0m'      # RESET COLOR
}

# Checks summary list
checks_summary = list()

# Checks must be "lowercase"
# Any Check at Service Field
# If Palo-Alto FW is present check Application field value - If No value or any is present then proceed
any_srv = new_frame.loc[
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Application Identity'].isnull() == True)
    |
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Application Identity'] == 'any')
    ]

if not any_srv.empty:
    any_srv.to_excel(writer, sheet_name="Any Service", startrow=0)
    checks_summary.append(f'FAIL - {len(any_srv["Service"])} Rules with Object "Any" were found in "Service" | "Application Identity" Fields')
else:
    checks_summary.append('PASS - No Object "Any" found in "Service" | "Application Identity" Fields')
# Any Check at Source Field
# If Palo-Alto FW is present check From zone field value - If No value or any is present then proceed
any_src = new_frame.loc[
    (new_frame['Source'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Source negated'] == 'false') &
    (new_frame['From zone'].isnull() == True)
    |
    (new_frame['Source'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Source negated'] == 'false') &
    (new_frame['From zone'] == 'any')
    ]

if not any_src.empty:
    any_src.to_excel(writer, sheet_name="Any Source", startrow=0)
    checks_summary.append(f'FAIL - {len(any_src["Service"])} Rules with Object "Any" were found in "Source" | "From Zone" Fields')
else:
    checks_summary.append('PASS - No Object "Any" found in "Source" | "From Zone" Fields')

# Any Check at Destination Field
# If Palo-Alto FW is present check To zone field value - If No value or any is present then proceed
any_dst = new_frame.loc[
    (new_frame['Destination'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Destination negated'] == 'false') &
    (new_frame['To zone'].isnull() == True)
    |
    (new_frame['Destination'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Destination negated'] == 'false') &
    (new_frame['To zone'] == 'any')
    ]

if not any_dst.empty:
    any_dst.to_excel(writer, sheet_name="Any Destination", startrow=0)
    checks_summary.append('FAIL - {len(any_dst["Destination"])} Rules with Object "Any" were found in "Destination" | "To Zone" Fields')
else:
    checks_summary.append('PASS - No Object "Any" found in "Destination" | "To Zone" Fields')

# Disabled Rules check
disabled_rules = new_frame.loc[new_frame['Rule status'] == 'disabled']

if not disabled_rules.empty:
    disabled_rules.to_excel(writer, sheet_name="Disabled rules", startrow=0)
    checks_summary.append(f'FAIL - {len(disabled_rules["Rule status"])} "Disabled" Rules were found')
else:
    checks_summary.append('PASS - No "Disabled" Rules were found')

# Reject rules check
reject_rules = new_frame.loc[
    (new_frame['Action'] == 'reject') &
    (new_frame['Rule status'] == 'enabled')
    ]

if not reject_rules.empty:
    reject_rules.to_excel(writer, sheet_name="Reject rules", startrow=0)
    checks_summary.append(f'FAIL - {len(reject_rules["Action"])} "Reject rules" Rules were found')
else:
    checks_summary.append('PASS - No "Reject" Object was found in "Action" Field')

# No Log rules
no_log_rules = new_frame.loc[
    (new_frame['Track'] == 'none') &
    (new_frame['Rule status'] == 'enabled')
    ]

if not no_log_rules.empty:
    no_log_rules.to_excel(writer, sheet_name="No Log rules", startrow=0)
    checks_summary.append(f'FAIL - {len(no_log_rules["Track"])} "No Log rules" Rules were found')
else:
    checks_summary.append('PASS - Logs are on for All Rules')

# Crossed Rules check
crossed_rules = new_frame.loc[
    (new_frame['Source'].isin(new_frame['Destination'])) &
    (new_frame['Rule status'] == 'enabled')]

# There may appear rules with the same Source & Destination but different Protocols,
# So, this one keeps rules with the same protocols
crossed_rules = crossed_rules[crossed_rules.duplicated(['Service'], keep=False)]

# Sorting rules by Service - for easier view
crossed_rules = crossed_rules.sort_values('Service')

if not crossed_rules.empty:
    crossed_rules.to_excel(writer, sheet_name="Crossed Rules", startrow=0)
    checks_summary.append(f'FAIL - {len(crossed_rules["Service"])} "Crossed Rules" Rules were found')
else:
    checks_summary.append('PASS - No "Crossed" Rules were found')

# Un-Safe Protocols rules
# Add as you wish to the list
unsafe_protocols = [
    'smb',
    'smbv1',
    'microsoft-ds',
    'telnet',
    'ftp',
    'http',
    'remote_desktop_protocol',
    'rdp',
    'sshv1'
]

unsafe_srv = new_frame.loc[
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service'].isin(unsafe_protocols)) |
    (new_frame['Application Identity'].isin(unsafe_protocols))
    ]

# Sorting rules by Service - for easier view
unsafe_srv = unsafe_srv.sort_values('Service')

if not unsafe_srv.empty:
    unsafe_srv.to_excel(writer, sheet_name="Un-Safe Protocols", startrow=0)
    checks_summary.append(f'FAIL - {len(unsafe_srv["Service"])} "Un-Safe Protocols" Rules were found')
else:
    checks_summary.append('PASS - No "Un-Safe" Protocols were found in "Service" | "Application Identity" Fields')

# Worst Rules - Presence of combination of multiple checks on one rule
# Example:
# From "Any" Zone & "Any" Source traffic may proceed to "Any" Zone & "Any" Destination on Any Service | Application
"""Needs Scripting"""


# Printing-out summary to console
for check in checks_summary:
    if check.startswith("PASS"):
        print(colors.get('PASS') + check + colors.get('RESET'))
    elif check.startswith("FAIL"):
        print(colors.get('FAIL') + check + colors.get('RESET'))
    else:
        pass

writer.save()
