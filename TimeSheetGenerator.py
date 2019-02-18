from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import pandas as pd

class TimeSheetGenerator(object):
    '''Need a sample file location for generating the actual timesheet document - xslx'''
    def __init__(self,config_json):
        self.config_json = config_json
        self.wb = load_workbook(filename=config_json["sample_file_location"])
        # Only the first sheet would be used - so hard codeing this
        self.sheet = self.wb[self.wb.sheetnames[0]] 
        self.dates_to_cell = self.map_dates_to_cells()

    def map_dates_to_cells(self):
        '''Dates 1 - 15 is range B13 to P13 and 16 - 31 is range B14 to Q14. Generate a dict with that info.
        Sheet object is the input'''
        dates_to_cell = {}
        dates_to_cell = dates_to_cell.fromkeys(range(1,32)) #Keys are dates from 1,31
        #Cell ranges for work hour cells are hard-coded. This never changes
        cell_range_B13_P13 = self.sheet['B13:P13'][0]
        cell_range_B14_Q14 = self.sheet['B14:Q14'][0] 

        for key,value in dates_to_cell.items():
            if (key > 0) and (key < 16):
                dates_to_cell[key] = cell_range_B13_P13[key-1]
            else:
                dates_to_cell[key] = cell_range_B14_Q14[key - 16]
    
        return dates_to_cell

    def reset_all_cells(self):
        '''Resets all cells in range to zero'''
        for key in self.dates_to_cell.keys():
            self.dates_to_cell[key].value = 0
        
    def fill_cell(self,map_of_dates_work_hrs):
        '''fill_cells based on map_of_dates_work_hrs - for eg {1:8,2:4} means 8 hours of work on 1 and 4 hours on 2'''
        for key,value in map_of_dates_work_hrs.items():
            self.dates_to_cell[key].value = value

    def fill_fillers(self):
        '''This is to fill meta-data in various cells. Al the cells here are harcoded'''
        #O4 is the month for which time sheet is generated

        #P4 is the year for which time sheet is generated
        #O5 x is if dates between 1 to 15
        #O6 x if dates between 16 to 31
        #R36 is todays date
        #Insert image
        img = Image("pridevel_image.png")
        self.sheet.add_image(img,'A1')
        

    def save_file(self):
        self.fill_fillers()
        self.wb.save(filename=self.config_json['file_to_generate'])


if __name__ == "__main__":
    test_config = {}
    test_config['sample_file_location'] = 'sample_pridevel.xlsx'
    test_config['file_to_generate'] = 'test.xlsx'

    #Generate a Timesheetgenerator object
    TG = TimeSheetGenerator(test_config)
    TG.reset_all_cells()
    TG.save_file()
