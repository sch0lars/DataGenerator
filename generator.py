import openpyxl
import random
import re
import string


# TODO: Remove the test portion of the code.
# Assume we have something like AX5[A-Z]{3}[0-9]{4}[A-Z]{2}[0-9]{1}
##### Test #####
# pattern = 'VX[A-Z]{3}[1-5]{5}'
# print(pattern)

# Load the first and last names.
with open('first-names.txt') as file:
    first_names = [name.strip() for name in file.read().split('\n')]
with open('last-names.txt') as file:
    last_names = [name.strip() for name in file.read().split('\n')]

# Substitute the patterns.
# Uppercase letters.
def replace_lowercase(pattern: str) -> str:
    for sub_pattern in re.findall('\[a\-z\]\{\d+\}', pattern):
        length = int(re.findall('(?<=\{)\d+(?=\})', sub_pattern)[0])
        replacement = ''.join(random.choice(string.ascii_lowercase) for _ in range(length))
        pattern = pattern.replace(sub_pattern, replacement, 1)
    return pattern
# Lowercase letters.
def replace_uppercase(pattern: str) -> str:
    for sub_pattern in re.findall('\[A\-Z\]\{\d+\}', pattern):
        length = int(re.findall('(?<=\{)\d+(?=\})', sub_pattern)[0])
        replacement = ''.join(random.choice(string.ascii_uppercase) for _ in range(length))
        pattern = pattern.replace(sub_pattern, replacement, 1)
    return pattern
# Digits
def replace_digits(pattern: str) -> str:
    for sub_pattern in re.findall('\[\d+\-\d+\]\{\d+\}', pattern):
        length = int(re.findall('(?<=\{)\d+(?=\})', sub_pattern)[0])
        start, end = [int(d) for d in re.findall('\d+\-\d+', sub_pattern)[0].split('-')]
        replacement = ''.join(str(random.randint(start, end)) for _ in range(length))
        pattern = pattern.replace(sub_pattern, replacement, 1)
    return pattern
# Decimals.
def replace_decimals(pattern: str) -> str:
    for sub_pattern in re.findall('\[\d+\.\d+\-\d+\.\d+\]\{\d+\}', pattern):
        length = int(re.findall('(?<=\{)\d+(?=\})', sub_pattern)[0])
        start, end = [int(float(d)*100) for d in re.findall('\d+\.\d+\-\d+\.\d+', sub_pattern)[0].split('-')]
        replacement = ''.join(str(random.randint(start, end)/100))
        pattern = pattern.replace(sub_pattern, replacement, 1)
    return pattern
# Lists of values
def replace_lists(pattern: str) -> str:
    for sub_pattern in re.findall('\[[\w\d\s\,]+\]\{\d+\}', pattern):
        length = int(re.findall('(?<=\{)\d+(?=\})', sub_pattern)[0])
        value_list = [value.strip() for value in re.findall('(?<=\[)[\w\d\s\,]+(?=\])', sub_pattern)[0].split(',')]
        replacement = ''.join(random.choice(value_list) for _ in range(length))
        pattern = pattern.replace(sub_pattern, replacement)
    return pattern
# Replace names.
def replace_names(pattern: str) -> str:
    first_name = random.choice(first_names)
    last_name = random.choice(last_names)
    email = first_name + random.choice(['.', '_', '']) + last_name + '@' + random.choice(['example.com', 'example.net', 'test.com', 'test.net', 'mail.com', 'mail.net', 'gmail.com', 'yahoo.com', 'outlook.com'])
    primary_address = ''.join(random.choice(string.digits) for _ in range(random.randint(2, 4))) + ' ' + random.choice(last_names).upper() + ' ' + random.choice(['RD', 'LN', 'AVE', 'BLVD', 'ST', 'PL', 'WAY'])
    secondary_address = random.choice(['APT ', ' STE' '#']) + ''.join(random.choice(string.digits) for _ in range(random.randint(1, 4)))
    pattern = pattern.replace('{{first_name}}', first_name)
    pattern = pattern.replace('{{last_name}}', last_name)
    pattern = pattern.replace('{{primary_address}}', primary_address)
    pattern = pattern.replace('{{primary_address}}', secondary_address)
    return pattern

##### Loading the Excel file #####
excel_file = 'tables.xlsx'
wb = openpyxl.load_workbook(excel_file)
# Create a dictionary of tables.
tables = {sheetname: {} for sheetname in wb.sheetnames}
# Create a dictionary of aliases.
aliases = {}
# Create a SQL string.
sql = ''
n = 10
for i in range(n):
    sql += f'-- Insert {i+1} of {n}\n'
    # Iterate through each worksheet and populate the table dictionary.
    for table in list(tables):
        # Get the worksheet.
        ws = wb[table]
        # Create a table dictionary in the tables dictionary.
        tables[table] = {}
        # Iterate through the worksheet.
        for row in ws.iter_rows():
            attribute = row[0].value
            pattern = row[1].value
            alias = row[2].value
            # Make pattern substitutions.
            pattern = replace_lowercase(pattern)
            pattern = replace_uppercase(pattern)
            pattern = replace_digits(pattern)
            pattern = replace_decimals(pattern)
            pattern = replace_lists(pattern)
            pattern = replace_names(pattern)
            # Add the alias if one is set.
            if alias and alias not in aliases:
                aliases[alias] = pattern
            # Create an attribute dictionary in the table dictinoary.
            tables[table][attribute] = {}
            # Add the values to the the dictionary.
            tables[table][attribute]['pattern'] = pattern
            tables[table][attribute]['alias'] = alias
        # Generate the SQL.
        fields = list(tables[table])
        values = [tables[table][field]['pattern'] if not tables[table][field]['alias'] else aliases[tables[table][field]['alias']] for field in fields]
        sql += f'INSERT INTO {table} '
        sql += '(' + ', '.join(fields) + ') '
        sql += 'VALUES ( ' + ', '.join([f"'{value}'" for value in values]) + ')'
        sql += '\n'
    sql += '\n'
# Close the workbook.
# wb.close()

