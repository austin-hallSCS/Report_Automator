"""
Austin's Report Automator!!! No more slogging through your school device reports!
Takes an SCCM report as a .xlsx file, removes all the unneccessary columns, adds the Dell warranty dates to the correct devices, then outputs it all as a beautiful spreadsheet.

CURRENTLY WORKING ON:

TO DO:
-Taking out playwright/ report automatic downloading

"""
# You know the drill, import the necessary modules
import os
import traceback
import keyboard
import reportAutomatorUtils as Utils
from report import Report

try:
    while True:
        #Clear the console
        os.system('cls')

        # Have user select report file, turn it into a dataframe
        report = Report()

        # Add warranty dates to the original report dataframe
        print("Getting warranty info...")
        report.add_Warranty_Dates()
        print("Dell warranty info added.")

        # Write the dataframe to an excel file, make it looks pretty and save it
        Utils.wait_For_Enter("Press enter to choose where to save the file...")
        report.format_And_Export()
        print(f"File has been saved to {report.outFile}")

        # Ask the user if they want to run another report, loop back to the beginning if they do, close the program if they don't
        print("Would you like to do another one?(Y/N)")
        userInput = keyboard.read_key()
        if userInput == "y":
            pass
        elif userInput == "n":
            print("kthxbye")
            break
except Exception as e:
    traceback.print_exception(e)
    Utils.wait_For_Enter("Press Enter to exit...")
finally:
    # Delete all files in "temp" and "dellWarrantyOutput" folders
    Utils.cleanup_Files()