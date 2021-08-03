""" Run with at least Python 3.8, pip install xlsxwriter if needed """

import os
import pandas as pd

# Specify path to .csv Reports
path = "C:\\path\\to\\folder\\containing\\tufin_reports\\"

# Iterate over .csv files in a path
files = [x for x in os.listdir(path=path) if x.endswith(".csv")]

# Adjust encoding if needed
encoding_files = "windows-1255"
number = 0

dataframe_list = list()

# Resolving "SettingWithCopyWarning"
# https://pandas.pydata.org/pandas-docs/stable/user_guide/indexing.html#returning-a-view-versus-a-copy
# SettingWithCopyWarning - Appeared when a .drop empty columns method was called before .to_excel method call
pd.set_option('mode.chained_assignment', None)

while (number := number + 1) < len(files):
    for f in files:
        file = path + f

        # Reading files
        main_frame = pd.read_csv(file, encoding=encoding_files)

        # Getting a row number of column set for rules
        # rules = main_frame.index[main_frame["Tufin Object lookup results"] == "Device name"].tolist()
        rules = main_frame[main_frame.isin(["Device name"]).any(axis="columns")]
        rules = rules.index.tolist()

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
    'PASS': '\033[92m',  # GREEN
    'WARNING': '\033[93m',  # YELLOW
    'FAIL': '\033[91m',  # RED
    'RESET': '\033[0m'  # RESET COLOR
}

# Checks summary list
checks_summary = list()


# Main Checks function
def check(data_frame: pd.DataFrame, sheet_name: str, column: list, pass_msg: str, fail_msg: str):
    if not data_frame.empty:
        data_frame.dropna(how='all', axis=1, inplace=True)
        data_frame.to_excel(writer, sheet_name=sheet_name, startrow=0)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        colorize = ['blue', 'green', 'red']

        # for s, d in zip(data_frame['Source'], data_frame['Destination']):
        #     if any(data_frame.column == s):
        #         finding_position = list(data_frame).index(column[0]) + 1
        #         cell_format = workbook.add_format({'bold': True, 'font_color': 'blue'})
        #         worksheet.set_column(first_col=finding_position, last_col=finding_position, cell_format=cell_format)
        #     if data_frame[data_frame['Destination'].str.contains(s)]:
        #         finding_position = list(data_frame).index(column[1]) + 1
        #         cell_format = workbook.add_format({'bold': True, 'font_color': 'green'})
        #         worksheet.set_column(first_col=finding_position, last_col=finding_position, cell_format=cell_format)

        if len(column) == 3:
            for col, color in zip(column, colorize):
                finding_position = list(data_frame).index(col) + 1
                cell_format = workbook.add_format({'bold': True, 'font_color': color})
                worksheet.set_column(first_col=finding_position, last_col=finding_position, cell_format=cell_format)
        elif len(column) < 2:
            finding_position = list(data_frame).index(column[0]) + 1
            cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
            worksheet.set_column(first_col=finding_position, last_col=finding_position, cell_format=cell_format)

        checks_summary.append(fail_msg)
    else:
        checks_summary.append(pass_msg)


# Checks must be "lowercase"
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

disabled_rules = new_frame.loc[
    new_frame['Rule status'] == 'disabled'
    ]

reject_rules = new_frame.loc[
    (new_frame['Action'] == 'reject') &
    (new_frame['Rule status'] == 'enabled')
    ]

no_log_rules = new_frame.loc[
    (new_frame['Track'] == 'none') &
    (new_frame['Rule status'] == 'enabled')
    ]

# Un-Done!!! Needs a FIX
crossed_list = list()
for src_cross, dst_cross in zip(new_frame['Source'], new_frame['Destination']):
    crossed_conditions = new_frame.loc[
        (new_frame['Destination'] == src_cross) &
        (new_frame['Source'] == dst_cross) &
        (new_frame['Rule status'] == 'enabled')
        # ADD Service comparison!
        ]
    if not crossed_conditions.empty:
        for index, row in crossed_conditions.iterrows():
            crossed_list.append(row.str.lower())

crossed_frame = pd.DataFrame(crossed_list)
crossed_frame.to_excel(writer, sheet_name='Crossed Rules')


# print(sources)
# print(destinations)
# crossed_rules = new_frame.loc[
#     (new_frame['Source'].isin(new_frame['Destination'])) &
#     (new_frame['Destination'].isin(new_frame['Source'])) &
#     (new_frame['Rule status'] == 'enabled')
#     ]
# There may appear rules with the same Source & Destination but different Protocols,
# So, this one keeps rules with the same protocols
# crossed_rules = crossed_rules[crossed_rules.duplicated(['Service'], keep=False)]
# # Sorting rules by Service - for easier view
# crossed_rules = crossed_rules.sort_values('Service')


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

# Any Check at Source, Destination & Service Fields
check(data_frame=any_srv, sheet_name="Any Service", column=["Service"], pass_msg="PASS - Any Service", fail_msg="FAIL - Any Service")
check(data_frame=any_src, sheet_name="Any Source", column=["Source"], pass_msg="PASS - Any Source", fail_msg="FAIL - Any Source")
check(data_frame=any_dst, sheet_name="Any Destination", column=["Destination"], pass_msg="PASS - Any Destination", fail_msg="FAIL - Any Destination")
# Disabled Rules check
check(data_frame=disabled_rules, sheet_name="Disabled rules", column=["Rule status"], pass_msg="PASS - Disabled rules", fail_msg="FAIL - Disabled rules")
# Reject rules check
check(data_frame=reject_rules, sheet_name="Reject rules", column=["Action"], pass_msg="PASS - Reject rules", fail_msg="FAIL - Reject rules")
# No Log rules
check(data_frame=no_log_rules, sheet_name="No Log rules", column=["Track"], pass_msg="PASS - No Log rules", fail_msg="FAIL - No Log rules")
# Crossed Rules check
# check(data_frame=crossed_rules, sheet_name="Crossed Rules", column=["Source", "Destination", "Service"], pass_msg="PASS - Crossed Rules", fail_msg="FAIL - Crossed Rules")
# Un-Safe Protocols rules
check(data_frame=unsafe_srv, sheet_name="Un-Safe Protocols", column=["Service"], pass_msg="PASS - Un-Safe Protocols", fail_msg="FAIL - Un-Safe Protocols")


# Worst Rules - Presence of combination of multiple checks on one rule
# Example:
# From "Any" Zone & "Any" Source traffic may proceed to "Any" Zone & "Any" Destination on Any Service | Application
"""Needs Scripting"""

# Printing-out summary to console
print("=" * len(max(checks_summary)) * 2)
print("Audit Checks")
print("=" * len(max(checks_summary)) * 2)

for check in sorted(checks_summary, reverse=True):
    if check.startswith("PASS"):
        print(colors.get('PASS') + check + colors.get('RESET'))
        print("-" * len(max(checks_summary)) * 2)
    elif check.startswith("FAIL"):
        print(colors.get('FAIL') + check + colors.get('RESET'))
        print("-" * len(max(checks_summary)) * 2)
    else:
        pass

writer.save()
