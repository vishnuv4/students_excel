import numpy as np
import pandas as pd
from color_print import color, print_blue, print_red
import click

EXCEL_FILE = "students_list.xlsx"
STUDENT_NAMES_SHEET = "Names"
RAW_DATA_SHEET = "Raw"

def write_to_sheet(data:any, filename:str, sheet_name:str, columns:list[str]):
    """
    Writes data to a sheet

    Args:
        data:
            List or np.ndarray (as of now) to be written
        filename:
            Name of the excel workbook
        sheet_name:
            Sheet to be written into (either creates a new one or replaces an existing one)
        columns:
            Column headers for the data
    """

    if isinstance(data, list):
        df = pd.DataFrame(pd.Series(data))
    elif isinstance(data, np.ndarray):
        df = pd.DataFrame(data)

    df.columns=columns
    # Write to excel
    with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    print_blue(f"Successfully wrote [{', '.join(columns)}] column(s) to the {sheet_name} sheet in {filename}")

def reformat_names(filename:str, sheet_name:str, num_students:int) -> list[str]:
    """
    When downloading the student list from canvas, the first column is
    the list of students in a <lastname, firstname> format.

    This function reads the data, takes the first column, and converts it
    to <firstname lastname> format in a new sheet called "Names".

    Args:
        filename:
            Name of the excel file.
        sheet_name:
            Name of the sheet that contains the full list.
        num_students:
            Expected number of students for validation reasons.

    """
    print_blue("Reading from the excel file...")
    df = pd.read_excel(filename, sheet_name=None)

    # The first row is a header and the last row is "Test Student", so excluding those in iloc
    names = list(map(str.strip, list(df[sheet_name].iloc[1:-1, 0])))

    # Reformat the names
    for i, name in enumerate(names):
        split_name = list(map(str.strip, name.split(",")))
        newname = split_name[1] + ' ' + split_name[0]
        names[i] = newname
    
    # Validation
    print_blue("Checking expected number of students...")
    if len(names) != num_students:
        raise ValueError("Number of students does not match")
    else:
        print_blue("Number of students is correct")

    return names

def get_names(filename:str, sheet_name:str) -> np.array:
    """
    Get the list of students as a numpy array

    Args:
        filename:
            Name of excel file
        sheet_name:
            Sheet containing a single column (it should be the first one)
            of the student names
    """
    print_blue(f"Reading the names of students from the {STUDENT_NAMES_SHEET} sheet...")
    df = pd.read_excel(filename, sheet_name=sheet_name)
    name_array = df.values
    return name_array

def random_pairings(names:np.array) -> np.ndarray:
    """
    Given an array of students, create a 2xN array with random pairings

    Args:
        names:
            Names of students
    """
    if len(names) % 2 != 0:
        print_blue("Odd number of students, adding an empty string to the list")
        names = np.append(names, "")
    else:
        print_blue("Even number of students, no need to add a placeholder")

    print_blue("Shuffling names in a random order...")
    np.random.shuffle(names)

    print_blue("Splitting the list of shuffled names into pairs...")
    pairings = names.reshape(-1, 2)

    return pairings

def check_duplicates(filename:str, sheet_list:list[str]):
    """
    Given an excel file and a list of sheet names to check, it checks for duplicate teams.
    The sheets must have names in the first and second column.

    Args:
        filename:
            Name of the excel file
        sheet_list:
            List of sheets to check duplicates between
    """
    data = pd.read_excel(filename, sheet_name=None)
    for i in data:
        data[i] = data[i].fillna("")


    if len(sheet_list) < 2:
        raise ValueError("There must be at least two sheets to compare")

    for sheet in sheet_list:
        if sheet not in list(data.keys()):
            raise NameError(f"Invalid sheet name! {sheet} is not a sheet in the excel file", "RED")

    for i in range(len(sheet_list)):
        for j in range(i+1, len(sheet_list)):
            print()

            ref_A = data[sheet_list[i]].iloc[:, 0].apply(str.strip)
            ref_B = data[sheet_list[i]].iloc[:, 1].apply(str.strip)
            compare_A = data[sheet_list[j]].iloc[:, 0].apply(str.strip)
            compare_B = data[sheet_list[j]].iloc[:, 1].apply(str.strip)

            lab1_pairs = list(zip(list(ref_A), list(ref_B)))
            lab2_pairs = list(zip(list(compare_A), list(compare_B)))

            lab1_pairs_set = [set(t) for t in lab1_pairs]
            lab2_pairs_set = [set(t) for t in lab2_pairs]

            common_pairs = [tuple(s) for s in lab1_pairs_set if s in lab2_pairs_set]

            if len(common_pairs) == 0:
                print(color(f"[{sheet_list[i]}, {sheet_list[j]}]\t\tNo common pairs found", "GREEN"))
            else:
                print(color(f"[{sheet_list[i]}, {sheet_list[j]}]\t\tCommon pairs found \t\t", "RED"))
                for c in common_pairs:
                    print(c)
    
    print()

### CLI RELATED STUFF ###

@click.group()
def cli():
    """
    Utility to manage student names and grouping into random pairs for labs.

    Meant to take in raw data from canvas and do everything else through this script.
    """

@cli.command(help="Format and save names in a new sheet")
@click.option(
    "--file", "-f",
    prompt="(--file, -f)\tEnter the excel file name",
    type=str,
    help=f"""\b
    Name of the file containing student info. 
    Default: {EXCEL_FILE}
    """,
    default=f"{EXCEL_FILE}",
    required=True
)
@click.option(
    "--raw", "-r",
    prompt="(--raw, -r)\tEnter the sheet containing raw data",
    type=str,
    help=f"""\b
    Name of sheet containing raw data. 
    Default: {RAW_DATA_SHEET}
    """,
    default=f"{RAW_DATA_SHEET}"
)
@click.option(
    "--output", "-o",
    prompt="(--output, -o)\tEnter the name of the sheet to write the names to",
    type=str,
    help="Name of the sheet to write the formatted names to",
    default=f"{STUDENT_NAMES_SHEET}"
)
@click.option(
    "--num", "-n",
    prompt="(--num, -n)\tEnter the number of students",
    type=int,
    help="Expected number of students, for validation purposes",
    required=True
)
def names(file, raw, output, num):
    formatted_names = reformat_names(filename=file, sheet_name=raw, num_students=num)
    write_to_sheet(data=formatted_names, filename=file, sheet_name=output, columns=[f"{output}"])


@cli.command(help="Pair students up randomly and save in a new sheet")
@click.option(
    "--file", "-f",
    prompt="(--file, -f)\tEnter the excel file name",
    type=str,
    help=f"""\b
    Name of the file containing student info. 
    Default: {EXCEL_FILE}
    """,
    default=f"{EXCEL_FILE}",
    required=True
)
@click.option(
    "--names", "-n",
    prompt="(--names, -n)\tEnter the name of the sheet with formatted names",
    type=str,
    help=f"""\b
    Name of the sheet containing formatted student names. 
    Default: {STUDENT_NAMES_SHEET}
    """,
    default=f"{STUDENT_NAMES_SHEET}"
)
@click.option(
    "--output", "-o",
    prompt="(--output, -o)\tEnter the sheet to save the pairing in",
    help="The sheet to save the pairing in",
    required=True,
)
def pair(file, names, output):
    names = get_names(file, names)
    pairings = random_pairings(names)
    write_to_sheet(data=pairings, filename=file, sheet_name=output, columns=["A", "B"])

@cli.command(help="Prints the list of sheets in the excel file")
@click.option(
    "--file", "-f",
    prompt="(--file, -f)\tEnter the excel file name",
    type=str,
    help=f"""\b
    Name of the file containing student info. 
    Default: {EXCEL_FILE}
    """,
    default=f"{EXCEL_FILE}",
    required=True
)
def sheets(file):
    print_blue("Name of sheets in excel workbook:")
    df = pd.read_excel(file, sheet_name=None)
    print(list(df.keys()))

@cli.command(help="Check for duplicates in a given list of sheets")
@click.option(
    "--file", "-f",
    prompt="(--file, -f)\tEnter the excel file name",
    type=str,
    help=f"""\b
    Name of the file containing student info. 
    Default: {EXCEL_FILE}
    """,
    default=f"{EXCEL_FILE}",
    required=True
)
@click.option(
    "--list", "-l",
    prompt="(--list, -l)\tEnter the list of sheets to check: ",
    type = str,
    help = """\b
    List of sheets that need to be checked.
    Enter the values as a single comma-separated string
    """,
    required=True,
)
def duplicates(file, list):
    split_list = [item.strip() for item in list.split(",")]
    check_duplicates(file, split_list)


### END OF CLI RELATED STUFF

if __name__ == "__main__":
    try:
        cli()
    except Exception as exp:
        print_red(str(exp))
