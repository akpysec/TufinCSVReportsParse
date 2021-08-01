import os
import pandas as pd

path = "C:\\path\\to\\folder\\containing\\tufin_reports\\"
files = [x for x in os.listdir(path=path) if x.endswith(".csv")]

number = 0

dataframe_list = []

while (number := number + 1) < len(files):
    for f in files:
        file = path + f

        # Reading files
        main_frame = pd.read_csv(file, encoding="windows-1252")

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

# Removing duplicate Rules
new_frame = new_frame.drop_duplicates(keep='first')

# Writing to all rules to "Rules" sheet
writer = pd.ExcelWriter(path + "Parsed_Rules.xlsx", engine='xlsxwriter')
new_frame.to_excel(writer, sheet_name="All Rules", startrow=0)


# Main function used across most of the checks - unpacks dataframe column & check if value is present
def check(column="Rule status", value="enabled".lower(), dataframe=new_frame):   # Just in Case .lower() is used
    """A default values used to save some code"""
    parse = dataframe[dataframe[column].str.contains(value)]
    return parse


# Checks must be "lowercase"
# Any Check at Service Field
if not check(column="Service", value="any").empty:
    if not check().empty:
        for src_neg in check(column="Service", value="any")['Service negated']:
            if src_neg == 'false':
                check(column="Service", value="any").to_excel(writer, sheet_name="Any Service", startrow=0)
            else:
                pass
    else:
        pass
else:
    pass

# Any Check at Source Field
if not check(column="Source", value="any").empty:
    if not check().empty:
        for src_neg in check(column="Source", value="any")['Source negated']:
            if src_neg == 'false':
                check(column="Source", value="any").to_excel(writer, sheet_name="Any Source", startrow=0)
            else:
                pass
    else:
        pass
else:
    pass

# Any Check at Destination Field
if not check(column="Destination", value="any").empty:
    if not check().empty:
        for src_neg in check(column="Destination", value="any")['Destination negated']:
            if src_neg == 'false':
                check(column="Destination", value="any").to_excel(writer, sheet_name="Any Destination", startrow=0)
            else:
                pass
    else:
        pass
else:
    pass

# Disabled Rules check
if not check(value="disabled").empty:
    check(value="disabled").to_excel(writer, sheet_name="Disabled rules", startrow=0)
else:
    pass

# Reject rules check
if not check(column="Action", value="reject").empty:
    if not check().empty:
        check(column="Action", value="reject").to_excel(writer, sheet_name="Reject rules", startrow=0)
    else:
        pass
else:
    pass

# No Log rules
if not check(column="Track", value="none").empty:
    if not check().empty:
        check(column="Track", value="none").to_excel(writer, sheet_name="No Log rules", startrow=0)
    else:
        pass
else:
    pass

# Crossed Rules check
for src_rules, src_srv_rules, src_secure_uid in zip(new_frame['Source'], new_frame['Service'],
                                                    new_frame['SecureTrack Rule UID']):
    for dst_rules, dst_srv_rules, dst_secure_uid in zip(new_frame['Destination'], new_frame['Service'],
                                                        new_frame['SecureTrack Rule UID']):
        if (src_rules, src_srv_rules) == (dst_rules, dst_srv_rules):
            crossed_rules = new_frame[new_frame['SecureTrack Rule UID'] == src_secure_uid].append(
                new_frame[new_frame['SecureTrack Rule UID'] == dst_secure_uid])
            crossed_rules = crossed_rules.drop_duplicates(keep='first')
            # Checking if found crossed rules in "Enabled" status
            if not check(dataframe=crossed_rules).empty:
                check(dataframe=crossed_rules).to_excel(writer, sheet_name="Crossed Rules", startrow=0)
            else:
                pass

# Un-Safe Protocols rules
unsafe_protocols = ['smb', 'microsoft-ds', 'telnet', 'ftp', 'http']
# Note searching for http will result in finding https also,
# consider removing if unnecessary, also add as you wish to the list
for protocol in unsafe_protocols:
    if not check(column="Service", value=protocol).empty:
        if not check().empty:
            check(column="Service", value=protocol).to_excel(writer, sheet_name="Un-Safe Protocols")
        else:
            pass
    else:
        pass

# RDP (Remote Desktop Protocol) rules
"""Needs Scripting"""

writer.save()
