"""
Run with at least Python 3.8, pip install xlsxwriter if needed.
If "pandas.errors.ParserError: Error tokenizing data. C error: Expected 2 fields in line 10, saw 6 Occurs"
Open each file and save as .csv again, after that this error will disappear
"""
import collections
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

while (number := number + 1) <= len(files):
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
new_frame.to_excel(writer, sheet_name="All Rules", startrow=0, index=False)

# Color Definition
# For Console use
colors = {
    'PASS': '\033[92m',  # GREEN
    'WARNING': '\033[93m',  # YELLOW
    'FAIL': '\033[91m',  # RED
    'RESET': '\033[0m'  # RESET COLOR
}

# For in Excel use
colorize = ["black",  # 0
            "blue",  # 1
            "brown",  # 2
            "cyan",  # 3
            "gray",  # 4
            "green",  # 5
            "lime",  # 6
            "magenta",  # 7
            "navy",  # 8
            "orange",  # 9
            "pink",  # 10
            "purple",  # 11
            "red",  # 12
            "silver",  # 13
            "white",  # 14
            "yellow"  # 15
            ]

# Checks summary list
checks_summary = list()


def console_print(summary: list):
    print(colors.get('WARNING') + "=" * len(max(summary)) * 2 + colors.get('RESET'))
    print(colors.get('WARNING') + "Audit Checks" + colors.get('RESET'))
    print(colors.get('WARNING') + "=" * len(max(summary)) * 2 + colors.get('RESET'))
    print("-" * len(max(summary)) * 2)

    for c in sorted(summary, reverse=True):
        if c.startswith("PASS"):
            print(colors.get('PASS') + c + colors.get('RESET'))
            print("-" * len(max(summary)) * 2)
        elif c.startswith("FAIL"):
            print(colors.get('FAIL') + c + colors.get('RESET'))
            print("-" * len(max(summary)) * 2)
        else:
            pass


# Main Checks function
def check(data_frame: pd.DataFrame, sheet_name: str, column: list, pass_msg: str, fail_msg: str):
    if not data_frame.empty:
        data_frame.dropna(how='all', axis=1, inplace=True)
        data_frame.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)

        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        for c in column:
            finding_position = list(data_frame).index(c)
            cell_format = workbook.add_format({'bold': True, 'font_color': colorize[12]})
            worksheet.set_column(first_col=finding_position, last_col=finding_position, cell_format=cell_format)

        checks_summary.append(fail_msg + f" | Total Rules found: {data_frame.shape[0]}")
    else:
        checks_summary.append(pass_msg)


# Crossed Rules check function
def check_crossed(data_frame: pd.DataFrame, sheet_name: str, pass_msg: str, fail_msg: str):
    crossed_list = list()

    for src_zone, src_cross, dst_zone, dst_cross, srv_cross, app_srv, act_cross in zip(
            data_frame['From zone'],
            data_frame['Source'],
            data_frame['To zone'],
            data_frame['Destination'],
            data_frame['Service'],
            data_frame['Application Identity'],
            data_frame['Action']):

        # Check if source zone, destination zone, source, destination are crossed,
        # Source user is not specified & services / app identity are equal,
        # That rule in enabled state & not negated.
        crossed_conditions = new_frame.loc[
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'].isnull()) &
            (data_frame['Source user'].isnull()) &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'allow') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'] == 'any') &
            (data_frame['Source user'] == 'any') &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'allow') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'].isnull()) &
            (data_frame['Source user'] == 'any') &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'allow') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'] == 'any') &
            (data_frame['Source user'].isnull()) &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'allow') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'].isnull()) &
            (data_frame['Source user'].isnull()) &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'accept') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'] == 'any') &
            (data_frame['Source user'] == 'any') &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'accept') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'].isnull()) &
            (data_frame['Source user'] == 'any') &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'accept') &
            (data_frame['Rule status'] == 'enabled')
            |
            (data_frame['From zone'] == dst_zone) &
            (data_frame['Source'] == dst_cross) &
            (data_frame['To zone'] == src_zone) &
            (data_frame['Destination'] == src_cross) &
            (data_frame['Service'] == srv_cross) &
            (data_frame['Application Identity'] == 'any') &
            (data_frame['Source user'].isnull()) &
            (data_frame['Source negated'] == 'false') &
            (data_frame['Destination negated'] == 'false') &
            (data_frame['Service negated'] == 'false') &
            (data_frame['Action'] == 'accept') &
            (data_frame['Rule status'] == 'enabled')
            ]

        if not crossed_conditions.empty:
            for idx, row_enum in crossed_conditions.iterrows():
                crossed_list.append(row_enum.str.lower())
            else:
                pass

    if crossed_list:
        # Appending iterated data to a DataFrame
        crossed_frame = pd.DataFrame(crossed_list)
        # Dropping empty columns
        crossed_frame.dropna(how='all', axis=1, inplace=True)
        # Dropping duplicate values based upon rule ID
        crossed_frame = crossed_frame.drop_duplicates(subset='SecureTrack Rule UID', keep="first")

        src_dst_cross = list()
        src_z_dst_z = list()
        # Sorting based on Service
        crossed_frame = crossed_frame.sort_values('Service')

        # Writing to a 'Crossed rules' sheet

        for sr, dst in zip(crossed_frame['Source'], crossed_frame['Destination']):
            src_dst_cross.append(sr)
            src_dst_cross.append(dst)

        for src_z, dst_z in zip(crossed_frame['From zone'], crossed_frame['To zone']):
            src_z_dst_z.append(src_z)
            src_z_dst_z.append(dst_z)

        # DataFrame columns to list
        crossed_columns = crossed_frame.columns.tolist()
        # DataFrame to list
        crossed_frame = crossed_frame.values.tolist()
        # Combining frame lists
        crossed_frame = [crossed_columns] + crossed_frame

        cross_workbook = writer.book
        cross_worksheet = cross_workbook.add_worksheet(sheet_name)
        # # position = list(crossed_frame).index('Source') + 1

        total_rows = len(crossed_frame) - 1  # Minus the header / column row

        # Creating a list for times to loop - times to loop is total rows without the header
        # Multiplied because of format application to 4 columns instead of 2
        rows_range = list(range(0, total_rows * 2))

        # Creating a switch like check later basing on if a value divisible by 2 - aka True / False
        # This is done only for coloring scheme on Crossed rules
        rows_range_true_false = list()

        for tf in rows_range:
            if (tf % 2) == 0:
                rows_range_true_false.append(True)
            else:
                rows_range_true_false.append(False)

        for cross_row, row_data in enumerate(crossed_frame):
            try:
                cross_worksheet.write_row(cross_row, 0, row_data)
            except TypeError:
                pass

        positions = [
            list(crossed_frame)[0].index('From zone'),
            list(crossed_frame)[0].index('Source'),
            list(crossed_frame)[0].index('To zone'),
            list(crossed_frame)[0].index('Destination'),
            list(crossed_frame)[0].index('Service')
        ]

        any_srv_format = cross_workbook.add_format(

            {
                'bold': True,
                'font_color': colorize[12]
            }
        )

        row_fmt = cross_workbook.add_format(

            {
                'bold': True,
                'font_color': colorize[0]
            }
        )

        cross_worksheet.set_column(first_col=positions[4] - 1, last_col=positions[4], cell_format=any_srv_format)
        cross_worksheet.set_row(0, cell_format=row_fmt)

        for n, r, z in zip(rows_range_true_false, src_dst_cross, src_z_dst_z):
            if n is True:
                fmt = cross_workbook.add_format(
                    {
                        'bold': True,
                        'font_color': colorize[0],
                        'border': 2,
                        'border_color': colorize[0],
                        'bg_color': colorize[9]
                    }
                )

                # To zone / From Zone column formatting
                cross_worksheet.conditional_format(
                    first_row=1,
                    first_col=positions[0],
                    last_row=total_rows,
                    last_col=positions[2],
                    options={
                        'type': 'cell',
                        'criteria': '==',
                        'value': f'"{z}"',
                        'format': fmt}
                )

                # Source / Destination column formatting
                cross_worksheet.conditional_format(
                    first_row=1,
                    first_col=positions[1],
                    last_row=total_rows,
                    last_col=positions[3],
                    options={
                        'type': 'cell',
                        'criteria': '==',
                        'value': f'"{r}"',
                        'format': fmt}
                )
            elif n is False:
                fmt = cross_workbook.add_format(
                    {
                        'bold': True,
                        'font_color': colorize[0],
                        'border': 2,
                        'border_color': colorize[0],
                        'bg_color': colorize[15]
                    }
                )

                # To zone / From Zone column formatting
                cross_worksheet.conditional_format(
                    first_row=1,
                    first_col=positions[0],
                    last_row=total_rows,
                    last_col=positions[2],
                    options={
                        'type': 'cell',
                        'criteria': '==',
                        'value': f'"{z}"',
                        'format': fmt}
                )

                # Source / Destination column formatting
                cross_worksheet.conditional_format(
                    first_row=1,
                    first_col=positions[1],
                    last_row=total_rows,
                    last_col=positions[3],
                    options={
                        'type': 'cell',
                        'criteria': '==',
                        'value': f'"{r}"',
                        'format': fmt}
                )

        checks_summary.append(fail_msg + f" | Total Rules found: {total_rows}")
    else:
        checks_summary.append(pass_msg)


# Add as you wish to the list
unsafe_protocols = [
    'smb',
    'smbv1',
    'smb_v1',
    'microsoft-ds',
    'telnet',
    'ftp',
    'http',
    'tcp_80',
    'remote_desktop_protocol',
    'rdp',
    'sshv1',
    'ssh_v1'
]

unsafe_dict = dict()
new_unsafe_dict = collections.defaultdict(list)
tmp = list()

# Creating dictionary UID: [Service_1, Service_2, etc]
for srv, uid in zip(new_frame['Service'], new_frame['SecureTrack Rule UID']):
    unsafe_dict[uid] = srv.split("\n")

# Checking if more than 1 item in Value list, if it's only http or http\nssh\smb\n etc..
# And comparing to items in a list, on both fronts (if multiple values in a list or a single)
for k, v in unsafe_dict.items():
    for un in unsafe_protocols:

        # If multiple items in a list, iterate over them and compare
        if len(v) > 1:
            for value in v:
                founded_values = list()
                if value == un:
                    founded_values.append(value)
                    # Using collection lib for adding list to a dictionary value, list of protocols / ports found
                    new_unsafe_dict[k].append(*founded_values)

        # If single item in a list, select[0] and compare
        elif len(v) < 2:
            founded_values = list()
            if v[0] == un:
                founded_values.append(un)
                new_unsafe_dict[k].append(*founded_values)

# Switching type from collections to standard dict
new_unsafe_dict = dict(new_unsafe_dict)

for key, values in new_unsafe_dict.items():
    values = str(values).strip("[]")  # Can do better
    unsafe_srv = new_frame.loc[new_frame['SecureTrack Rule UID'] == key]
    unsafe_srv['Service'] = unsafe_srv['Service'] = [x.replace(x, values) for x in unsafe_srv['Service']]

    for index, row in unsafe_srv.iterrows():
        tmp.append(row.str.lower())

unsafe = pd.DataFrame(tmp)

# Checks must be "lowercase"
# Basically filtering column values & then using these filtered DF in the check(dataframe=DF) function
any_srv = new_frame.loc[
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['Application Identity'].isnull())
    |
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['Application Identity'] == 'any')
    |
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
    (new_frame['Application Identity'].isnull())
    |
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
    (new_frame['Application Identity'] == 'any')
    |
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['Application Identity'] == 'application-default')
    |
    (new_frame['Service'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Service negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
    (new_frame['Application Identity'] == 'application-default')
    ]

any_src = new_frame.loc[
    (new_frame['Source'] == 'any') &
    (new_frame['Source user'].isnull()) &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Source negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['From zone'].isnull())
    |
    (new_frame['Source'] == 'any') &
    (new_frame['Source user'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Source negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['From zone'] == 'any')
    |
    (new_frame['Source'] == 'any') &
    (new_frame['Source user'].isnull()) &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Source negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
    (new_frame['From zone'].isnull())
    |
    (new_frame['Source'] == 'any') &
    (new_frame['Source user'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Source negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
    (new_frame['From zone'] == 'any')
    ]

any_dst = new_frame.loc[
    (new_frame['Destination'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Destination negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['To zone'].isnull())
    |
    (new_frame['Destination'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Destination negated'] == 'false') &
    (new_frame['Action'] == 'allow') &
    (new_frame['To zone'] == 'any')
    |
    (new_frame['Destination'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Destination negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
    (new_frame['To zone'].isnull())
    |
    (new_frame['Destination'] == 'any') &
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Destination negated'] == 'false') &
    (new_frame['Action'] == 'accept') &
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
    (new_frame['Action'] == 'allow') &
    (new_frame['Rule status'] == 'enabled')
    |
    (new_frame['Track'] == 'none') &
    (new_frame['Action'] == 'accept') &
    (new_frame['Rule status'] == 'enabled')
    ]

unsafe_srv = unsafe.loc[
    (unsafe['Rule status'] == 'enabled') &
    (unsafe['Action'] == 'allow')
    |
    (unsafe['Rule status'] == 'enabled') &
    (unsafe['Action'] == 'accept')
    ]

# Sorting rules by Service - for easier view
unsafe_srv = unsafe_srv.sort_values('Service')

worst_rules = new_frame.loc[
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Action'] == 'accept') &
    (new_frame['Source'] == 'any') &
    (new_frame['Destination'] == 'any')
    |
    (new_frame['Rule status'] == 'enabled') &
    (new_frame['Action'] == 'allow') &
    (new_frame['Source'] == 'any') &
    (new_frame['Destination'] == 'any')
    ]

# Any Check at Source, Destination & Service Fields
check(
    data_frame=any_srv,
    sheet_name="Any Service",
    column=["Service"],
    pass_msg="PASS - Any Service",
    fail_msg="FAIL - Any Service"
)
check(
    data_frame=any_src,
    sheet_name="Any Source",
    column=["Source"],
    pass_msg="PASS - Any Source",
    fail_msg="FAIL - Any Source"
)

check(
    data_frame=any_dst,
    sheet_name="Any Destination",
    column=["Destination"],
    pass_msg="PASS - Any Destination",
    fail_msg="FAIL - Any Destination"
)

# Disabled Rules
check(
    data_frame=disabled_rules,
    sheet_name="Disabled rules",
    column=["Rule status"],
    pass_msg="PASS - Disabled rules",
    fail_msg="FAIL - Disabled rules"
)

# Reject rules check
check(
    data_frame=reject_rules,
    sheet_name="Reject rules",
    column=["Action"],
    pass_msg="PASS - Reject rules",
    fail_msg="FAIL - Reject rules"
)

# No Log rules
check(
    data_frame=no_log_rules,
    sheet_name="No Log rules",
    column=["Track"],
    pass_msg="PASS - No Log rules",
    fail_msg="FAIL - No Log rules"
)

# Un-Safe Protocols rules
check(
    data_frame=unsafe_srv,
    sheet_name="Un-Safe Protocols",
    column=["Service"],
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
check(
    data_frame=worst_rules,
    sheet_name="Worst Rules",
    column=[
        'Source',
        'Destination',
        'Service'
    ],
    pass_msg="PASS - Worst Rules",
    fail_msg="FAIL - Worst Rules"
)

# Printing-out summary to console
console_print(summary=checks_summary)

writer.save()
