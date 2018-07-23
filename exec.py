import main, SheetUtil, sys
from os import system
from time import sleep

while True:
    system('cls')
    module = int(input("To run a function, enter the corresponding number\n"
                      "1. Sort turns and mainline counts\n"
                      "2. Remove duplicate pdfs\n"
                      "3. Format ADT counts\n"
                      "4. Format Counts Unlimited Counts"
                      "5. Exit Program\n"))

    if module == 1:
        main.excelSort()
    elif module == 2:
        main.pdfDuplicates()
    elif module == 3:
        main.ADTFormat()
    elif module == 4:
        main.CountsUnlimitedFormat()
    elif module == 5:
        sys.exit()
    else:
        print "Invalid input"

    raw_input("Press enter to continue")
    print("The scipt will now restart...")
    sleep(3)

