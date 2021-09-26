"""
Run with at least Python 3.8, pip install xlsxwriter if needed.
If "pandas.errors.ParserError: Error tokenizing data. C error: Expected 2 fields in line 10, saw 6 Occurs"
Open each file and save as .csv again, after that this error will disappear
Good to know - Microsoft Excel doesn't approve more than 150 chars in a cell when conditional formatting is applied,
so the script will not be able to fill color to those fields, but the check still be performed.
"""

import collections
import os
import pandas as pd

try:

    FIELDS = [
        'from zone',
        'to zone',
        'source',
        'destination',
        'service',
        'application identity',
        'rule status',
        'action',
        'track',
        'securetrack rule uid',
        'service negated'
    ]

    # Specify path to .csv Reports
    path_to_files = str(input("Enter a path to .CSV reports folder:\n")) + "\\"
    # path = "C:\\path\\to\\folder\\containing\\tufin_reports\\"
    
    _path_to_files = []
    if len(path_to_files) >= 2:
        _path_to_files.append(path_to_files)

        # Iterate over .csv files in a path
        files = [x for x in os.listdir(path=_path_to_files[0]) if x.endswith(".csv")]

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
                # Reading files
                main_frame = pd.read_csv(_path_to_files[0] + f, encoding=encoding_files)

                # Lowering case of all cells in the dataframe for unity
                main_frame = main_frame.applymap(lambda x: x.lower() if pd.notnull(x) else x)

                # Getting a row number of column set for rules
                for settings in main_frame.itertuples():
                    if FIELDS[2] and FIELDS[3] and FIELDS[4] in settings:

                        # Changing main columns to the identified column for rules
                        main_frame.columns = main_frame.iloc[settings[0]]

                        # Dropping each row until founded columns + 1 - it's self,
                        # what leaves me with a new assigned columns & rules only
                        main_frame = main_frame.drop(main_frame.index[0:settings[0] + 1])

                        for index, row in main_frame.iterrows():
                            # Appending to list for further creation of a DataFrame
                            dataframe_list.append(row)

        # Creating a new DataFrame from a list
        new_frame = pd.DataFrame(dataframe_list)

        # Removing duplicate Rules from a DataFrame
        new_frame = new_frame.drop_duplicates(subset=FIELDS[9], keep='first')

        # Writing to all rules to "All Rules" sheet
        writer = pd.ExcelWriter(_path_to_files[0] + "Parsed_Rules.xlsx", engine='xlsxwriter')
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
            if not isinstance(data_frame, pd.DataFrame):
                checks_summary.append(pass_msg)
                pass
            elif isinstance(data_frame, pd.DataFrame):
                if not data_frame.empty:
                    data_frame.dropna(how='all', axis=1, inplace=True)
                    data_frame.to_excel(writer, sheet_name=sheet_name, startrow=0, index=False)

                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    for c in column:
                        finding_position = list(data_frame).index(c)
                        cell_format = workbook.add_format({'bold': True, 'font_color': colorize[12]})
                        worksheet.set_column(
                            first_col=finding_position,
                            last_col=finding_position,
                            cell_format=cell_format)

                    checks_summary.append(fail_msg + f" | Total Rules found: {data_frame.shape[0]}")
                else:
                    checks_summary.append(pass_msg)
            else:
                print('Something else happened')

        # Crossed Rules check function
        def check_crossed(data_frame: pd.DataFrame, sheet_name: str, pass_msg: str, fail_msg: str):
            crossed_list = list()

            for src_zone, src, dst_zone, dst, srv_cross, app_srv, act_cross, status in zip(
                    data_frame[FIELDS[0]],
                    data_frame[FIELDS[2]],
                    data_frame[FIELDS[1]],
                    data_frame[FIELDS[3]],
                    data_frame[FIELDS[4]],
                    data_frame[FIELDS[5]],
                    data_frame[FIELDS[7]],
                    data_frame[FIELDS[6]]):

                # Check if source zone, destination zone, source, destination are crossed,
                # source user is not specified & services / app identity are equal,
                # That rule in enabled state & not negated.
                crossed_conditions = new_frame.loc[
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'application-default') &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'application-default') &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]].isnull()) &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'accept') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == dst) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[3]] == src) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == data_frame[FIELDS[1]]) &
                    (data_frame[FIELDS[1]] == data_frame[FIELDS[0]]) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
                    |
                    (data_frame[FIELDS[0]] == dst_zone) &
                    (data_frame[FIELDS[1]] == src_zone) &
                    (data_frame[FIELDS[2]] == data_frame[FIELDS[3]]) &
                    (data_frame[FIELDS[3]] == data_frame[FIELDS[2]]) &
                    (data_frame[FIELDS[4]] == srv_cross) &
                    (data_frame[FIELDS[5]] == 'any') &
                    (data_frame[FIELDS[7]] == 'allow') &
                    (data_frame[FIELDS[6]] == 'enabled')
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
                crossed_frame = crossed_frame.drop_duplicates(subset=FIELDS[9], keep="first")

                src_dst_cross = list()
                src_z_dst_z = list()
                # Sorting based on service
                crossed_frame = crossed_frame.sort_values(FIELDS[4])

                # Writing to a 'Crossed rules' sheet

                for sr, dst in zip(crossed_frame[FIELDS[2]], crossed_frame[FIELDS[3]]):
                    src_dst_cross.append(sr)
                    src_dst_cross.append(dst)

                for src_z, dst_z in zip(crossed_frame[FIELDS[0]], crossed_frame[FIELDS[1]]):
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
                    list(crossed_frame)[0].index(FIELDS[0]),
                    list(crossed_frame)[0].index(FIELDS[2]),
                    list(crossed_frame)[0].index(FIELDS[1]),
                    list(crossed_frame)[0].index(FIELDS[3]),
                    list(crossed_frame)[0].index(FIELDS[4])
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

                cross_worksheet.set_column(first_col=positions[4] - 1,
                                           last_col=positions[4],
                                           cell_format=any_srv_format)

                cross_worksheet.set_row(0, cell_format=row_fmt)

                crossed_frame = pd.DataFrame(crossed_frame)
                crossed_frame.columns = crossed_columns
                crossed_frame.drop(axis=0, index=0)

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

                        # to zone / from zone column formatting
                        cross_worksheet.conditional_format(
                            first_row=1,
                            first_col=positions[0],
                            last_row=total_rows,
                            last_col=positions[0],
                            options={
                                'type': 'cell',
                                'criteria': '==',
                                'value': f'"{z}"',
                                'format': fmt}
                        )
                        # to zone / from zone column formatting
                        cross_worksheet.conditional_format(
                            first_row=1,
                            first_col=positions[2],
                            last_row=total_rows,
                            last_col=positions[2],
                            options={
                                'type': 'cell',
                                'criteria': '==',
                                'value': f'"{z}"',
                                'format': fmt}
                        )
                        # source / destination column formatting
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

                        # to zone / from zone column formatting
                        cross_worksheet.conditional_format(
                            first_row=1,
                            first_col=positions[0],
                            last_row=total_rows,
                            last_col=positions[0],
                            options={
                                'type': 'cell',
                                'criteria': '==',
                                'value': f'"{z}"',
                                'format': fmt}
                        )
                        # to zone / from zone column formatting
                        cross_worksheet.conditional_format(
                            first_row=1,
                            first_col=positions[2],
                            last_row=total_rows,
                            last_col=positions[2],
                            options={
                                'type': 'cell',
                                'criteria': '==',
                                'value': f'"{z}"',
                                'format': fmt}
                        )

                        # source / destination column formatting
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

                    else:
                        pass

                checks_summary.append(fail_msg + f" | Total Rules found: {total_rows}")
            else:
                checks_summary.append(pass_msg)


        def unsafe_protocols(data_frame: pd.DataFrame, protocols: list):
            # Add as you wish to the list

            unsafe_dict = dict()
            new_unsafe_dict = collections.defaultdict(list)
            tmp = list()

            # Creating dictionary UID: [service_1, service_2, etc]
            for srv, uid in zip(data_frame[FIELDS[4]], data_frame[FIELDS[9]]):
                unsafe_dict[uid] = srv.split("\n")

            # Checking if more than 1 item in Value list, if it's only http or http\ssh\smb\n etc..
            # And comparing to items in a list, on both fronts (if multiple values in a list or a single)
            for k, v in unsafe_dict.items():
                for un in protocols:

                    # If multiple items in a list, iterate over them and compare
                    if len(v) > 1:
                        for value in v:
                            founded_values = list()
                            if value == un:
                                founded_values.append(value)
                                # Using collection lib for adding list to a dictionary value,
                                # list of protocols / ports found
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
                unsafe_srv = data_frame.loc[data_frame[FIELDS[9]] == key]
                unsafe_srv[FIELDS[4]] = unsafe_srv[FIELDS[4]] = [x.replace(x, values) for x in unsafe_srv[FIELDS[4]]]

                for i, r in unsafe_srv.iterrows():
                    tmp.append(r.str.lower())

            unsafe = pd.DataFrame(tmp)

            if not unsafe.empty:
                # return unsafe
                unsafe_srv = unsafe.loc[
                    (unsafe[FIELDS[6]] == 'enabled') &
                    (unsafe[FIELDS[7]] == 'allow')
                    |
                    (unsafe[FIELDS[6]] == 'enabled') &
                    (unsafe[FIELDS[7]] == 'accept')
                    ]

                # Sorting rules by service - for easier view
                unsafe_srv = unsafe_srv.sort_values(FIELDS[4])

                return unsafe_srv
            else:
                # print('Dataframe unsafe is empty')
                pass


        # Checks must be "lowercase"
        # Basically filtering column values & then using these filtered DF in the check(dataframe=DF) function
        any_srv = new_frame.loc[
            (new_frame[FIELDS[4]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[10]] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[5]].isnull())
            |
            (new_frame[FIELDS[4]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[10]] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[5]] == 'any')
            |
            (new_frame[FIELDS[4]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[10]] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[5]].isnull())
            |
            (new_frame[FIELDS[4]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[10]] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[5]] == 'any')
            |
            (new_frame[FIELDS[4]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[10]] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[5]] == 'application-default')
            |
            (new_frame[FIELDS[4]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[10]] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[5]] == 'application-default')
            ]

        any_src = new_frame.loc[
            (new_frame[FIELDS[2]] == 'any') &
            (new_frame['source user'].isnull()) &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['source negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[0]].isnull())
            |
            (new_frame[FIELDS[2]] == 'any') &
            (new_frame['source user'] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['source negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[0]] == 'any')
            |
            (new_frame[FIELDS[2]] == 'any') &
            (new_frame['source user'].isnull()) &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['source negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[0]].isnull())
            |
            (new_frame[FIELDS[2]] == 'any') &
            (new_frame['source user'] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['source negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[0]] == 'any')
            ]

        any_dst = new_frame.loc[
            (new_frame[FIELDS[3]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['destination negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[1]].isnull())
            |
            (new_frame[FIELDS[3]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['destination negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[1]] == 'any')
            |
            (new_frame[FIELDS[3]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['destination negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[1]].isnull())
            |
            (new_frame[FIELDS[3]] == 'any') &
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame['destination negated'] == 'false') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[1]] == 'any')
            ]

        disabled_rules = new_frame.loc[
            new_frame[FIELDS[6]] == 'disabled'
            ]

        reject_rules = new_frame.loc[
            (new_frame[FIELDS[7]] == 'reject') &
            (new_frame[FIELDS[6]] == 'enabled')
            ]

        no_log_rules = new_frame.loc[
            (new_frame[FIELDS[8]] == 'none') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[6]] == 'enabled')
            |
            (new_frame[FIELDS[8]] == 'none') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[6]] == 'enabled')
            ]

        worst_rules = new_frame.loc[
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[7]] == 'accept') &
            (new_frame[FIELDS[2]] == 'any') &
            (new_frame[FIELDS[3]] == 'any')
            |
            (new_frame[FIELDS[6]] == 'enabled') &
            (new_frame[FIELDS[7]] == 'allow') &
            (new_frame[FIELDS[2]] == 'any') &
            (new_frame[FIELDS[3]] == 'any')
            ]

        # Any Check at source, destination & service Fields
        check(
            data_frame=any_srv,
            sheet_name="Any service",
            column=["service"],
            pass_msg="PASS - Any service",
            fail_msg="FAIL - Any service"
        )
        check(
            data_frame=any_src,
            sheet_name="Any source",
            column=["source"],
            pass_msg="PASS - Any source",
            fail_msg="FAIL - Any source"
        )

        check(
            data_frame=any_dst,
            sheet_name="Any destination",
            column=["destination"],
            pass_msg="PASS - Any destination",
            fail_msg="FAIL - Any destination"
        )

        # Disabled Rules
        check(
            data_frame=disabled_rules,
            sheet_name="Disabled rules",
            column=["rule status"],
            pass_msg="PASS - Disabled rules",
            fail_msg="FAIL - Disabled rules"
        )

        # Reject rules check
        check(
            data_frame=reject_rules,
            sheet_name="Reject rules",
            column=["action"],
            pass_msg="PASS - Reject rules",
            fail_msg="FAIL - Reject rules"
        )

        # No Log rules
        check(
            data_frame=no_log_rules,
            sheet_name="No Log rules",
            column=["track"],
            pass_msg="PASS - No Log rules",
            fail_msg="FAIL - No Log rules"
        )

        # Un-Safe Protocols rules
        check(
            data_frame=unsafe_protocols(
                data_frame=new_frame,
                protocols=[
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
                ],
            ),
            sheet_name="Un-Safe Protocols",
            column=["service"],
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
        # From "Any" Zone & "Any" source traffic may proceed to:
        # "Any" Zone & "Any" destination on Any service | Application
        check(
            data_frame=worst_rules,
            sheet_name="Worst Rules",
            column=[
                FIELDS[2],
                FIELDS[3],
                FIELDS[4]
            ],
            pass_msg="PASS - Worst Rules",
            fail_msg="FAIL - Worst Rules"
        )

        # Printing-out summary to console
        console_print(summary=checks_summary)

        writer.save()

    elif len(path_to_files) < 2:
        print("Nothing entered, \nPlease enter a legitimate path & re-run the program.")

except FileNotFoundError:
    print(f'Wrong path - "{_path_to_files[0]}" (or files are missing), \nCheck yourself & re-run the program')
