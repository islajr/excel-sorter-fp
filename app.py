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

    connector = []
    current_list = []

    # main loop
    while True:
        while count >= 1:
            re_prompt = input("Do you want to sort for another company? (y/n) ")

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
        new_book = openpyxl.Workbook()
        new_sheet = new_book.active

        # search for pattern in file and store them
        for rows in range(1, 10001):
            # condition for EOF: Three empty lines in a row.
            if sheet[f"C{rows}"].value is None and sheet[f"C{rows + 1}"].value is None:  # if the iterator comes on
                # two empty lines
                if sheet[f"C{rows + 2}"].value is None:  # check if the following line is also empty.
                    break  # break out of loop

            # logic for crossing domains
            # if sheet[f"C{rows}"].value is None and sheet[f"C{rows + 2}"].value is not None:
            #     # connector.append([])
            #     # connector.append(list(sheet[f"C{rows + 1}"].value))
            #     # connector.append([])
            #     ...

            # finding and copying the heading
            if sheet[f"C{rows}"].value is None:
                ...
            elif "Participant" in sheet[f"C{rows}"].value and sheet[f"C{rows}"].value is not None:
                current = list(sheet[rows])
                for i in current:
                    current_list.append(i.value)

                # looping through the current list inorder to append its values independently to the connector
                connector.append([])    # create an empty list to house the heading
                for i in range(len(current_list)):
                    connector[0].append(current_list[i])
                current_list.clear()

            # logic for grepping pattern
            if sheet[f"C{rows}"].value is None:
                continue
            elif pattern in sheet[f"C{rows}"].value and sheet[f"C{rows}"].value is not None:
                current = list(sheet[rows])
                for i in current:
                    current_list.append(i.value)

                # looping through the current list inorder to append its values independently to the connector
                connector.append([])  # create an empty list to house the heading
                for i in current_list:
                    connector[len(connector) - 1].append(i)
                current_list.clear()

        # error checking for final step
        if len(connector) < 2:
            print("{} not found!".format(pattern))
        else:
            # writing to new file from storage
            for i in connector:
                new_sheet.append(i)  # new worksheet uses just as many columns as required.

            new_book.save(f"./output/{pattern}.xlsx")
            print("Successfully sorted {}".format(pattern))

        count += 1


if __name__ == "__main__":
    main()
