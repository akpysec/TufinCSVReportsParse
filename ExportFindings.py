"""
Run with at least Python 3.8, pip install xlsxwriter if needed.
If "pandas.errors.ParserError: Error tokenizing data. C error: Expected 2 fields in line 10, saw 6 Occurs"
Open each file and save as .csv again, after that this error will disappear
"""

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
new_frame = new_frame.drop_duplicates(subset='SecureTrack Rule UID', keep='first')

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
def check(data_frame: pd.DataFrame, sheet_name: str, column: str, pass_msg: str, fail_msg: str):
    if not data_frame.empty:
        data_frame.dropna(how='all', axis=1, inplace=True)
        data_frame.to_excel(writer, sheet_name=sheet_name, startrow=0)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        finding_position = list(data_frame).index(column) + 1
        cell_format = workbook.add_format({'bold': True, 'font_color': 'red'})
        worksheet.set_column(first_col=finding_position, last_col=finding_position, cell_format=cell_format)

        checks_summary.append(fail_msg)
    else:
        checks_summary.append(pass_msg)


# Crossed Rules check function
def check_crossed(data_frame: pd.DataFrame, sheet_name: str, pass_msg: str, fail_msg: str):

    crossed_list = list()

    for src_cross, dst_cross, srv_cross in zip(data_frame['Source'], data_frame['Destination'], data_frame['Service']):
        try:
            crossed_conditions = new_frame.loc[
                (data_frame['Destination'] == src_cross) &
                (data_frame['Source'] == dst_cross) &
                (data_frame['Service'] == srv_cross) &
                (data_frame['Rule status'] == 'enabled')
                ]
            # !!! FIX ISSUE !!!
            if not crossed_conditions.empty:
                for index, row_enum in crossed_conditions.iterrows():
                    crossed_list.append(row_enum.str.lower())
            else:
                pass
        except KeyError:
            print(f"{KeyError} Occurred")
            pass

    if len(crossed_list) > 1:
        # Appending iterated data to a DataFrame
        crossed_frame = pd.DataFrame(crossed_list)

        # Dropping empty columns
        crossed_frame.dropna(how='all', axis=1, inplace=True)
        # Dropping duplicate values based upon rule ID
        crossed_frame = crossed_frame.drop_duplicates(subset='SecureTrack Rule UID', keep="first")

        unique = list()

        # Sorting based on Service
        crossed_frame = crossed_frame.sort_values('Service')

        # Writing to a 'Crossed rules' sheet

        for sr, dst in zip(crossed_frame['Source'], crossed_frame['Destination']):
            unique.append(sr)
            unique.append(dst)

        # DataFrame columns to list
        crossed_columns = crossed_frame.columns.tolist()
        # DataFrame to list
        crossed_frame = crossed_frame.values.tolist()
        # Combining frame lists
        crossed_frame = [crossed_columns] + crossed_frame

        cross_workbook = writer.book
        cross_worksheet = cross_workbook.add_worksheet(sheet_name)
        # # position = list(crossed_frame).index('Source') + 1

        colorize = ['green', 'blue']
        total_rows = len(crossed_frame) - 1  # Minus the header / column row

        # Creating a list for times to loop - times to loop is total rows without the header
        rows_range = list(range(0, total_rows))

        # Creating a switch like check later basing on if a value divisible by 2 - aka True / False
        # This is done only for coloring scheme on Crossed rules
        rows_range_true_false = list()

        for tf in rows_range:
            if (tf % 2) == 0:
                rows_range_true_false.append(True)
            else:
                rows_range_true_false.append(False)

        for row_enum, row_data in enumerate(crossed_frame):
            cross_worksheet.write_row(row_enum, 0, row_data)
            any_srv_format = cross_workbook.add_format({'bold': True, 'font_color': 'red'})
            cross_worksheet.set_column(first_col=13, last_col=13, cell_format=any_srv_format)
            row_fmt = cross_workbook.add_format({'bold': True, 'font_color': 'black'})
            cross_worksheet.set_row(0, cell_format=row_fmt)

            for n, u in zip(rows_range_true_false, unique):
                if n is True:
                    fmt = cross_workbook.add_format({'bold': True, 'font_color': colorize[0]})
                    cross_worksheet.conditional_format(
                        f"I2:K{len(crossed_frame)}",
                        {
                            'type': 'cell',
                            'criteria': '==',
                            'value': f'"{u}"',
                            'format': fmt
                        }
                    )
                elif n is False:
                    fmt = cross_workbook.add_format({'bold': True, 'font_color': colorize[1]})
                    cross_worksheet.conditional_format(
                        f"I2:K{len(crossed_frame)}",
                        {
                            'type': 'cell',
                            'criteria': '==',
                            'value': f'"{u}"',
                            'format': fmt
                        }
                    )

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
check(
    data_frame=any_srv,
    sheet_name="Any Service",
    column="Service",
    pass_msg="PASS - Any Service",
    fail_msg="FAIL - Any Service"
)
check(
    data_frame=any_src,
    sheet_name="Any Source",
    column="Source",
    pass_msg="PASS - Any Source",
    fail_msg="FAIL - Any Source"
)
check(
    data_frame=any_dst,
    sheet_name="Any Destination",
    column="Destination",
    pass_msg="PASS - Any Destination",
    fail_msg="FAIL - Any Destination"
)

# Disabled Rules check
check(
    data_frame=disabled_rules,
    sheet_name="Disabled rules",
    column="Rule status",
    pass_msg="PASS - Disabled rules",
    fail_msg="FAIL - Disabled rules"
)

# Reject rules check
check(
    data_frame=reject_rules,
    sheet_name="Reject rules",
    column="Action",
    pass_msg="PASS - Reject rules",
    fail_msg="FAIL - Reject rules"
)

# No Log rules
check(
    data_frame=no_log_rules,
    sheet_name="No Log rules",
    column="Track",
    pass_msg="PASS - No Log rules",
    fail_msg="FAIL - No Log rules"
)

# Un-Safe Protocols rules
check(
    data_frame=unsafe_srv,
    sheet_name="Un-Safe Protocols",
    column="Service",
    pass_msg="PASS - Un-Safe Protocols",
    fail_msg="FAIL - Un-Safe Protocols"
)

# Crossed Rules check
check_crossed(
    data_frame=new_frame,
    sheet_name="Crossed Rules",
    pass_msg="PASS - Crossed Rules",
    fail_msg="FAIL - Crossed Rules"
)

# Worst Rules - Presence of combination of multiple checks on one rule
# Example:
# From "Any" Zone & "Any" Source traffic may proceed to "Any" Zone & "Any" Destination on Any Service | Application
"""Needs Scripting"""

# Printing-out summary to console
print(colors.get('WARNING') + "=" * len(max(checks_summary)) * 2 + colors.get('RESET'))
print(colors.get('WARNING') + "Audit Checks" + colors.get('RESET'))
print(colors.get('WARNING') + "=" * len(max(checks_summary)) * 2 + colors.get('RESET'))
print("-" * len(max(checks_summary)) * 2)

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
