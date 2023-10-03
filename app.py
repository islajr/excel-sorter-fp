# Steps
"""
- request for input
- search for occurrences in source file
- copy entire row of occurrence to new file
- when done, ask for input again
- repeat process on new sheet.
"""
import openpyxl
import sys


def main():
    count = 0

    # load a workbook
    source = openpyxl.load_workbook("./src/S4 BALANCES FOR THE MONTH ENDED 30.09.2023.xlsx")
    sheet = source.active

    # main loop
    while True:
        while count >= 1:
            re_prompt = input("Do you want to continue? (y or n) ")

            re_prompt.strip().lower()

            if re_prompt == "y" or re_prompt == "yes":
                break

            elif re_prompt == "n" or re_prompt == "no":
                sys.exit()

            else:
                print("Provide a valid answer! ")
                continue

        pattern = input("What do you wish to sort? ")

        # input validation
        pattern.strip().capitalize()

        # search for pattern in file
        for rows in range(1, 50001):
            # condition for EOF: Three empty lines in a row.
            if sheet[f"C{rows}"].value is None and sheet[f"C{rows + 1}"].value is None:  # if the iterator comes on
                # two empty lines
                if sheet[f"C{rows + 2}"].value is None:  # check if the following line is also empty.
                    break  # break out of loop

            # logic for crossing domains
            if sheet[f"C{rows}"].value is None and sheet[f"C{rows + 1}"].value is not None:
                print(sheet[f"C{rows + 1}"].value)

            # accounting for a possible two-line break
            elif sheet[f"C{rows}"].value is None and sheet[f"C{rows + 1}"].value is None and sheet[f"C{rows + 2}"].value is not None:
                print(sheet[f"C{rows + 1}"].value)
            #   logic for grepping pattern
            else:
                if pattern in sheet[f"C{rows}"].value:
                    print(rows, sheet[f"C{rows}"].value)

        # saving to a file
        # source.save("test.csv")

        # accessing one cell
        # source.cell(row='', column='', value='')
        # # or
        # cell = source['A4'] = 4

        # looping through cells

        count = count + 1


if __name__ == "__main__":
    main()
