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
        pattern.strip()

        # load a workbook
        source = openpyxl.load_workbook("./NLPC PFA.xlsx")
        sheet = source.active

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
