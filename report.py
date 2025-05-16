import pandas as pd
import os
import time
from datetime import datetime as dt, timedelta
from filepath import FilePath
import reportAutomatorUtils as Utils
class Report:
    def __init__(self):
        self.inFile = FilePath(fileIsOut=False, )
        self.dataFrame = self.prepare_Spreadsheet(self.inFile.filePath)
        self.key = pd.read_csv('Dell_Key.csv', index_col=0, header=None, usecols=[0, 1])[1]

    # Loads excel file as a dataframe and formats the data.
    @staticmethod
    def prepare_Spreadsheet(file):
        # Load the excel file as a dataframe
        df = pd.read_excel(file, skiprows=4)
        
        # Drop the empty columns
        df.dropna(how='all', axis=1, inplace=True)

        # Convert the first row to headers
        headers = df.iloc[0]
        df = pd.DataFrame(df.values[1:], columns=headers)

        # Remove all the columns we don't use
        headersToDelete = ['Domain/Workgroup', 'ConfigMgr Site Name', 'Service Pack Level', 'Memory (KBytes)', 'Processor (GHz)', 'Total Disk Space (MB)', 'Free Disk Space (MB)']
        for header in headersToDelete:
            try:
                df.drop(header, axis=1, inplace=True)
            except KeyError as e:
                print(f"There was an error removing one or more unused columns:\n{e}")

        # Add new column at the end for warranty dates
        df['Warranty End Date'] = ''

        return df
    
    # Adds warranty dates from the Dell info dataframe to the report dataframe
    def add_Warranty_Dates(self):
        # Get the warranty info from the local Dell Key file and add to the dataframe
        for i in self.key.index:
            for j in self.dataFrame.index:
                if isinstance(self.dataFrame.loc[j, 'Serial Number'], (float)):
                    continue
                if i in self.dataFrame.loc[j, 'Serial Number']:
                    self.dataFrame.loc[j, 'Warranty End Date'] = self.key.loc[i]

        # Convert the 'Warranty End Date' column to datetime type
        self.dataFrame['Warranty End Date'] = pd.to_datetime(self.dataFrame['Warranty End Date'])        

    # Makes the excel sheet look pretty
    def format_And_Export(self):
        while True:
            # Ask to name/where to save the file
            self.outFile = Utils.gui_File_Path_Out(".xlsx")

            # Write the dataframe to excel
            try:
                writer = pd.ExcelWriter(self.outFile, engine='xlsxwriter', datetime_format='mm/dd/yyyy')
            except FileNotFoundError:
                print("Uh oh! Looks like you didn't select a valid file path. Please try again.")
            else:
                break
        self.dataFrame.to_excel(writer, index=False, sheet_name='Sheet1')

        # Get the XlsxWriter workbook and worksheet objects
        wb = writer.book
        ws = writer.sheets['Sheet1']

        # Get cell notation for last column of dataframe
        lastColIndex = self.dataFrame.shape[1] - 1
        lastColRange = (1, lastColIndex, len(self.dataFrame), lastColIndex)
        
        # Variable to store dates
        today = dt.now()
        outOfWarranty = (today - timedelta(days=(3*365.24)))
        outOfService = (today - timedelta(days=(5*365.24)))
        
        # Variables to store the conditional formats
        format_In_Warranty = wb.add_format({'bg_color': 'green', 'font_color': 'white'})
        format_In_Service = wb.add_format({'bg_color': 'yellow', 'font_color': 'black'})
        format_Out_Of_Service = wb.add_format({'bg_color': 'red', 'font_color': 'white'})

        # Apply conditional format to the last column
        ws.conditional_format(*lastColRange,
                      {'type': 'date',
                      'criteria': '>',
                      'value': today,
                      'format': format_In_Warranty})
        ws.conditional_format(*lastColRange,
                      {'type': 'date',
                       'criteria': 'between',
                       'minimum': outOfService,
                       'maximum': today,
                       'format': format_In_Service})
        ws.conditional_format(*lastColRange,
                        {'type': 'date',
                         'criteria': '<',
                         'value': outOfService,
                         'format': format_Out_Of_Service})
    
        # Adjust column widths based on content
        for i, col in enumerate(self.dataFrame.columns):
            try:
                width = max(self.dataFrame[col].apply(lambda x: len(str(x))).max(), len(col))
                ws.set_column(i, i, width)
            except Exception:
                print("There was an error adjusting the width of one or more columns.")


        # Save and quit
        writer.close()