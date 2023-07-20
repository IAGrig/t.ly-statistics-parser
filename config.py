import string, itertools

TOKEN = ""
FILEPATH = ""

COLUMNS = []
for length in range(1, 4):
    for column_name in itertools.product(string.ascii_uppercase, repeat=length):
        COLUMNS.append(''.join(column_name))
