from openpyxl import load_workbook
import numpy as np
import matplotlib.pyplot as plt
import pylab
workbook = "C:\\Users\\user\\OneDrive - University of Edinburgh\\excelSheet\\age_at_death.xlsx"

class Sheet(object):
    def __init__(self, excel_workbook_name, sheet_name):
                #clear previous workbook
        self.wb = None
        self.excel_workbook_name = excel_workbook_name
        self.sheet_name = sheet_name
        ##set workbook as file object
        try:
            self.wb = load_workbook(self.excel_workbook_name, data_only = True)
        except IOError:
            self.wb = None
            print("Could not open file"+ self.excel_workbook_name)
        ##get sheet as file object
        self.sheet = self.wb[self.sheet_name]
        
class Column(Sheet):
    def __init__(self, excel_workbook_name, sheet_name):
        Sheet.__init__(self, excel_workbook_name, sheet_name)
        self.column_label = ""
        self.column_list = []
        self.mult_column_lists = []
    def set_column_label(self, row, column):
        try:
            val = self.sheet.cell(row, column).value
        except:   
            raise ColumnLabelError("A problem occurred setting the row label.")
        self.column_label = val     
    def get_column_label(self):
        return self.column_label
    def set_column_list(self, start_row, column):
        self.column_list = []
        for row in range(start_row, 1000):        
            try:
                val = self.sheet.cell(row, column).value
            except AttributeError:   
                print("A problem occurred reading data from a row")            
            if val == None:
                break
            self.column_list = self.column_list + [val]    
    def get_column_list(self):
        return self.column_list
    def set_mult_column_lists(self):
        self.mult_column_lists = []
        i = 1
        while True:
            self.set_column_list(2, i)
            if self.column_list == []:
                break;
            self.mult_column_lists.append(self.column_list)
            i += 1
    def get_mult_column_lists(self):
        return self.mult_column_lists
    
    def display_box_plot(self, variable):
        fig, box_plot = plt.subplots() #remember plt = matplotlib.pyplot
        box_plot.set_title(variable)
        box_plot.boxplot(self.column_list)
        pylab.show()

    def display_mult_box_plots(self, variable):
        fig, box_plot = plt.subplots() #remember plt = matplotlib.pyplot
        ##box_plot.set_axisbelow(True)
        box_plot.set_title(variable)
        box_plot.set_ylabel("Age")
        box_plot.boxplot(self.mult_column_lists)
        box_plot.set_xticklabels(["Age","Age minus 5"], rotation=45, fontsize=8)
        pylab.show()

c = Column(workbook,"Sheet1")
c.set_mult_column_lists()
##print(c.get_mult_column_lists())
##c.set_column_list(2,1)
##print(c.get_column_list())
c.display_mult_box_plots("AgeAtDeath")
##c.display_box_plot("AgeAtDeath")

    


### method for creating work book object and worksheetobject
### method for extracting data from a column and putting it in a list
###set_column_list
###get_column_list
###method for setting up plot
