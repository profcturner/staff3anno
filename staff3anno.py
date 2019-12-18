"""Produces a spreadsheet from STAFF III annotations with one line per file"""

import openpyxl


def process_row_br(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Baseline room (BR) section

    see process_row for argument data
    """

    print(f"  BR,  source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["D"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "BaselineRoom"

    return last_dest_row


def process_row_cr1(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Cathlab (CR) section

    see process_row for argument data
    """

    print(f"  CR1, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["E"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "BaselineCathlab1"

    return last_dest_row


def process_row_cr2(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Cathlab (CR) section

    see process_row for argument data
    """

    print(f"  CR2, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["F"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "BaselineCathlab2"

    return last_dest_row


def process_row_bi1(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Balloon Inflation (bi) section

    see process_row for argument data
    """

    print(f"  BI1, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["G"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = source_sheet["H"+srow].value
    # Copy the inflation times.
    inflations = source_sheet["I"+srow].value.split(";")
    dest_sheet["G"+drow] = inflations[0]
    dest_sheet["H"+drow] = inflations[1]
    dest_sheet["I"+drow] = inflations[2]

    return last_dest_row


def process_row_bi2(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Balloon Inflation (bi) section

    see process_row for argument data
    """

    print(f"  BI2, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["K"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = source_sheet["L"+srow].value
    # Copy the inflation times.
    inflations = source_sheet["M"+srow].value.split(";")
    dest_sheet["G"+drow] = inflations[0]
    dest_sheet["H"+drow] = inflations[1]
    dest_sheet["I"+drow] = inflations[2]

    return last_dest_row


def process_row_bi3(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Balloon Inflation (bi) section

    see process_row for argument data
    """

    print(f"  BI3, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["O"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = source_sheet["P"+srow].value
    # Copy the inflation times.
    inflations = source_sheet["Q"+srow].value.split(";")
    dest_sheet["G"+drow] = inflations[0]
    dest_sheet["H"+drow] = inflations[1]
    dest_sheet["I"+drow] = inflations[2]

    return last_dest_row


def process_row_bi4(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Balloon Inflation (bi) section

    see process_row for argument data
    """

    print(f"  BI4, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["S"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = source_sheet["T"+srow].value
    # Copy the inflation times.
    inflations = source_sheet["U"+srow].value.split(";")
    dest_sheet["G"+drow] = inflations[0]
    dest_sheet["H"+drow] = inflations[1]
    dest_sheet["I"+drow] = inflations[2]

    return last_dest_row


def process_row_bi5(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Balloon Inflation (bi) section

    see process_row for argument data
    """

    print(f"  BI5, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["V"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = source_sheet["W"+srow].value
    # Copy the inflation times.
    inflations = source_sheet["X"+srow].value.split(";")
    dest_sheet["G"+drow] = inflations[0]
    dest_sheet["H"+drow] = inflations[1]
    dest_sheet["I"+drow] = inflations[2]

    return last_dest_row


def process_row_pc1(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Postinflation Cathlab (CR) section

    see process_row for argument data
    """

    print(f"  PC1, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["Y"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "PostCathlab1"

    return last_dest_row


def process_row_pc2(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Postinflation Cathlab (CR) section

    see process_row for argument data
    """

    print(f"  PC2, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["Z"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "PostCathlab2"

    return last_dest_row


def process_row_pr1(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Postinflation Room (PR) section

    see process_row for argument data
    """

    print(f"  PR1, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["AA"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "PostRoom1"

    return last_dest_row


def process_row_pr2(row, source_sheet, dest_sheet, last_dest_row):
    """
    Process the Postinflation Room (PR) section

    see process_row for argument data
    """

    print(f"  PR2, source row {row}")
    # Take a new row
    last_dest_row += 1
    # Convert that row to a string and the source row
    drow = str(last_dest_row)
    srow = str(row)

    # Copy the filename, and pad it with zeroes
    dest_sheet["A"+drow] = f'{source_sheet["AB"+srow].value:0>4}'
    # Copy the patient
    dest_sheet["B"+drow] = source_sheet["A"+srow].value
    # Copy the age
    dest_sheet["C"+drow] = source_sheet["B"+srow].value
    # Copy the sex
    dest_sheet["D"+drow] = source_sheet["C"+srow].value
    # Copy prior MI location
    dest_sheet["E"+drow] = source_sheet["AC"+srow].value
    # Copy the "artery"
    dest_sheet["F"+drow] = "PostRoom2"

    return last_dest_row


def process_row(row, source_sheet, dest_sheet, last_dest_row):
    """

    :param row:             the row number in the source sheet to process
    :param source_sheet:    the source spreadsheet
    :param dest_sheet:      the destination spreadsheet
    :param last_dest_row:   the last used row in the destination
    :return:                the new last used row in the destination
    """

    # Get the string version of the row
    srow = str(row)

    # Look for Baseline Room Data (in column D)
    if source_sheet["D"+srow].value:
        last_dest_row = process_row_br(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Baseline Cathlab Data (in column E)
    if source_sheet["E"+srow].value:
        last_dest_row = process_row_cr1(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Baseline Cathlab Data (in column F)
    if source_sheet["F"+srow].value:
        last_dest_row = process_row_cr2(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Balloon inflation 1 data (in column G)
    if source_sheet["G"+srow].value:
        last_dest_row = process_row_bi1(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Balloon inflation 2 data (in column K)
    if source_sheet["K"+srow].value:
        last_dest_row = process_row_bi2(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Balloon inflation 3 data (in column O)
    if source_sheet["O"+srow].value:
        last_dest_row = process_row_bi3(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Balloon inflation 4 data (in column S)
    if source_sheet["S"+srow].value:
        last_dest_row = process_row_bi4(row, source_sheet, dest_sheet, last_dest_row)

    # Look for Balloon inflation 5 data (in column V)
    if source_sheet["V"+srow].value:
        last_dest_row = process_row_bi5(row, source_sheet, dest_sheet, last_dest_row)

    # Look for post inflation Cathlab data (in column Y)
    if source_sheet["Y"+srow].value:
        last_dest_row = process_row_pc1(row, source_sheet, dest_sheet, last_dest_row)

    # Look for post inflation Cathlab data (in column Z)
    if source_sheet["Z"+srow].value:
        last_dest_row = process_row_pc2(row, source_sheet, dest_sheet, last_dest_row)

    # Look for post inflation room data (in column AA)
    if source_sheet["AA"+srow].value:
        last_dest_row = process_row_pr1(row, source_sheet, dest_sheet, last_dest_row)

    # Look for post inflation room data (in column AB)
    if source_sheet["AB"+srow].value:
        last_dest_row = process_row_pr2(row, source_sheet, dest_sheet, last_dest_row)

    return last_dest_row


def main():
    """Starts the process"""

    print("Staff III helper - produces a line per file spreadsheet from annotation data.")

    # Here's the name of the standard spreadsheet supplied with the data
    input_filename = "STAFF-III-Database-Annotations.xlsx"
    # And the row where the data starts and ends
    data_row_start = 11
    data_row_end =  118

    # Specific the name for the output spreadsheet here
    output_filename = "output.xlsx"

    # Get the default sheet from the source
    source_workbook = openpyxl.load_workbook(input_filename, data_only=True, read_only=True)
    source_sheet = source_workbook.active

    # And select the default sheet in the output
    dest_workbook = openpyxl.Workbook()
    dest_sheet = dest_workbook.active

    dest_sheet["A1"]="filename"
    dest_sheet["B1"]="patient"
    dest_sheet["C1"]="age"
    dest_sheet["D1"]="gender"
    dest_sheet["E1"]="prior_mi"
    dest_sheet["F1"]="artery"
    dest_sheet["G1"]="inflation_start"
    dest_sheet["H1"]="inflation_duration"
    dest_sheet["I1"]="inflation_after"

    # Keep track of the last row in the destination
    last_dest_row = 1

    for row in range(data_row_start, data_row_end+1):
        last_dest_row = process_row(row, source_sheet, dest_sheet, last_dest_row)

    dest_workbook.save(output_filename)


if __name__ == '__main__':
    main()