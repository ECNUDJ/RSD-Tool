#! python3

import tkinter as tk
import tkinter.messagebox
from decimal import *
from time import localtime, strftime
import openpyxl

# The file which holds information about calibration gas serial numbers/ makeup.
# Change this to the absolute path of your gas dictionary text file.
# Keep it in the format """ GasSerialNumber:GasMakeup \n """
gasDictFile = 'dictFile.txt'

# The Excel file which your calibration results will be stored in.
# Change this to the absolute path of your Excel file.
excelFile = 'sampleExcelFile.xlsx'

# Set 10 places after decimal for using Decimal module
getcontext().prec = 10

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.grid()
        self.create_widgets()
        self.calculate_RSD         
        self.add_To_Spreadsheet
        self.open_Settings_Window
        self.open_About_Window
        self.fetch_Serial_Numbers
        self.fetch_Gas_Dict

    def create_widgets(self):
        # Create the 'Calculate RSD' button
        self.calc_RSD_Button = tk.Button(self)
        self.calc_RSD_Button["text"] = "Calculate RSD:"
        self.calc_RSD_Button["command"] = self.calculate_RSD
        self.calc_RSD_Button.config(pady=3, padx=3)
        self.calc_RSD_Button.grid(row = 6, column = 0)

        # Spacer for the bottom
        separator = tk.Frame(height=15, bd=1, relief='sunken')
        separator.config(padx=5, pady=5)
        separator.grid(row = 7, column = 0)

        # Create a label for peak area entry boxes
        self.label = tk.Label(self, text="Peak Areas").grid(row = 0, column = 0)

        # Create the peak area entry boxes (x5)
        self.x0 = tk.Entry(self, width="7")
        self.x0.grid(row = 1, column = 0)

        self.x1 = tk.Entry(self, width="7")
        self.x1.grid(row = 2, column = 0)

        self.x2 = tk.Entry(self, width="7")
        self.x2.grid(row = 3, column = 0)

        self.x3 = tk.Entry(self, width="7")
        self.x3.grid(row = 4, column = 0)

        self.x4 = tk.Entry(self, width="7")
        self.x4.grid(row = 5, column = 0)

        # Create a label for RSD output
        self.rsdOutput = tk.Label(self, fg="red", width=7, font=("Helvetica", 14))
        self.rsdOutput.grid(row = 6, column = 1)

        # Create dropdown menu for total number of GC runs, with a label one row above it
        self.runsLabel = tk.Label(self, text="Total # of GC Runs:").grid(row = 0, column = 3)
        self.variable = tk.StringVar()
        self.runs = tk.OptionMenu(self, self.variable, "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12",
        "13", "14", "15").grid(row = 1, column = 3)

        # Create dropdown menu for choosing calibration standard S/N, with a label one row above
        gases = self.fetch_Serial_Numbers(gasDictFile)
        
        self.calLabel = tk.Label(self, text="Calibration Standard").grid(row = 0, column = 4)
        self.variable2 = tk.StringVar()
        self.calStandard = tk.OptionMenu(self, self.variable2, *gases)
        self.calStandard.grid(row = 1, column = 4)

        # Button for entering RSD into excel sheet
        self.enterRSD = tk.Button(self, text="Enter into Excel Sheet", command = self.add_To_Spreadsheet)
        self.enterRSD.grid(row = 3, column = 4)
        self.enterRSD.config(pady=3, padx=3)

    
        ############################################################
        ##################       MENU          #####################
        ############################################################

        # create a toplevel menu
        self.menubar = tk.Menu(self)

        # create a pulldown menu, and add it to the menu bar
        filemenu = tk.Menu(self.menubar, tearoff=0)
        filemenu.add_command(label="Settings", command=self.open_Settings_Window)
        filemenu.add_command(label="About", command=self.open_About_Window)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=root.destroy)
        self.menubar.add_cascade(label="File", menu=filemenu)

        # display the menu
        root.config(menu=self.menubar)

    # Define function for 'Calculate RSD' button command
    def calculate_RSD(self):
        runs = 0
        peakAreas = []
        
        if self.x0.get() != "":
            runs += 1
            peakAreas.append(float(self.x0.get()))
        if self.x1.get() != "":
            runs += 1
            peakAreas.append(float(self.x1.get()))
        if self.x2.get() != "":
            runs += 1
            peakAreas.append(float(self.x2.get()))
        if self.x3.get() != "":
            runs += 1
            peakAreas.append(float(self.x3.get()))
        if self.x4.get() != "":
            runs += 1
            peakAreas.append(float(self.x4.get()))

        # Calculate average peak area (xAvg)
        num = 0.0
        for items in peakAreas:
            num += items
        xAvg = num/runs

        # Copy runs into a tkinter variable
        self.runs = tk.StringVar()
        self.runs = runs

        # Calculate RSD values in cases of 2,3,4 or 5 entries
        if runs == 2:
            sd = float(Decimal(((peakAreas[0] - xAvg)**2 + (peakAreas[1] - xAvg)**2)/1).sqrt())
            rsd = sd/ xAvg
            rsd = round(rsd * 100, 2)
            self.rsdOutput["text"] = str(rsd) + " %"
            self.rsd = tk.StringVar()
            self.rsd = rsd

        elif runs == 3:
            sd = float(Decimal(((peakAreas[0] - xAvg)**2 + (peakAreas[1] - xAvg)**2 + (peakAreas[2] - xAvg)**2)/2).sqrt())
            rsd = sd/ xAvg
            rsd = round(rsd * 100, 2)
            self.rsdOutput["text"] = str(rsd) + " %"
            self.rsd = tk.StringVar()
            self.rsd = rsd

        elif runs == 4:
            sd = float(Decimal(((peakAreas[0] - xAvg)**2 + (peakAreas[1] - xAvg)**2 + (peakAreas[2] - xAvg)**2 \
            + (peakAreas[3] - xAvg)**2)/3).sqrt())
            rsd = sd/ xAvg
            rsd = round(rsd * 100, 2)
            self.rsdOutput["text"] = str(rsd) + " %"
            self.rsd = tk.StringVar()
            self.rsd = rsd

        elif runs == 5:
            sd = float(Decimal(((peakAreas[0] - xAvg)**2 + (peakAreas[1] - xAvg)**2 + (peakAreas[2] - xAvg)**2 \
            + (peakAreas[3] - xAvg)**2 + (peakAreas[4] - xAvg)**2)/4).sqrt())
            rsd = sd/ xAvg
            rsd = round(rsd * 100, 2)
            self.rsdOutput["text"] = str(rsd) + " %"
            self.rsd = tk.StringVar()
            self.rsd = rsd

    def add_To_Spreadsheet(self):

        time = strftime("%Y-%m-%d %H:%M:%S", localtime())

        calGases = self.fetch_Gas_Dict(gasDictFile)

        # Show error if user forgot to enter Total # of GC Trials
        if self.variable.get() == "":

            tk.messagebox.showerror(
            "Missing Info",
            "Enter Total # of GC Trials"
        )
            return False

        # Data for entering a single row into Excel spreadsheet
        try:
            data =  {
                "Date" : time,
                "Calibration S/N" : self.variable2.get(),
                "Calibration Gas" : calGases[self.variable2.get()],
                "Trials" : self.runs,
                "Total Trials" : self.variable.get(),
                "RSD" : self.rsd
            }
        except KeyError:
            tk.messagebox.showerror(
            "Missing Info",
            "Enter a Calibration Gas S/N"
        )
        
        # Open Excel file, navigate to appropriate sheet
        wb = openpyxl.load_workbook(excelFile)
        ws = wb['RSD Log']

        try:
            ws.append((data["Date"], data["Calibration Gas"], data["Calibration S/N"], data["Trials"], \
                       data["Total Trials"], data["RSD"]))
            wb.save(excelFile)
            # Show info box that data was entered successfully
            tk.messagebox.showinfo(
            "Success",
            "Data entered into %s" % excelFile
        )

        except PermissionError or UnboundLocalError:
            tk.messagebox.showerror(
            "Close Excel File",
            "Cannot edit file while it is open.\nClose %s" % filename
        )

    def open_Settings_Window(self):

        # Finds the length of a file, in number of lines
        def file_len(fileObj):
            for i, l in enumerate(fileObj):
                pass
            return i + 1

        # Define function for adding new gas to calGases dictionary
        def add_To_Cal_Gases():
            # Open file in append mode
            fileObj = open(gasDictFile, 'a')
            # Retrieve new entry from entry widgets
            newGas = settings.SNEntry.get() + ":" + settings.gasEntry.get()
            # Add a newline and the new gas to the text file
            fileObj.write("\n")
            fileObj.write(newGas)
            # Close the file
            fileObj.close()
            # Show info box that new gas was entered successfully
            tk.messagebox.showinfo(
            "Success",
            "   Gas was added. \n  Restart app to see it \n in the dropdown menu."
            )

        # Define function for removing gas from dictionary
        def remove_Gas():
            # Open file
            f = open(gasDictFile, 'r+')

            # Get serial number: gas makeup entry that we want to remove from our dictionary
            SN = str(settings.variable3.get())
            gasDict = self.fetch_Gas_Dict(gasDictFile)
            gasMakeup = gasDict[settings.variable3.get()].strip()
            entryToRemove = SN + ":" + gasMakeup

            # Get the file contents in an array called d and then seek back to beginning of file
            d = f.readlines()
            f.seek(0)

            # Get length of file and seek back to beginning
            length = file_len(f)
            f.seek(0)

            # Get the line # of the line we want to remove from text file 
            j = 0
            for i in d:
                j += 1
                if i.strip() == entryToRemove:
                    lineNumToRemove = j
            f.seek(0)

            # Write new file, minus entryToRemove. If entryToRemove is currently the last entry in our dict,
            # then remove the newline at the end of the new final entry in the dict. This is to prevent our
            # dictionary from having white lines at the end of it, which can mess up some other widgets.
            k = 0
            for i in d:
                k += 1
                if i.strip() != entryToRemove:
                    if lineNumToRemove == length:
                        if k == length - 1:
                            f.write(i.strip())
                        else:
                            f.write(i)
                    else:
                        f.write(i)

            # Show info box that gas was removed successfully
            tk.messagebox.showinfo(
            "Success",
            "%s removed from list " % settings.variable3.get()
            )

            # Truncate and close the file        
            f.truncate()
            f.close()
                
        # Create the new window
        settings = tk.Toplevel(self, height=120, width=330)

        ### Provide functionality for adding a new calibration standard cylinder
        settings.addCalGas = tk.Label(settings, text="Add New Calibration Standard").place(x=0, y=0)
        # Enter S/N
        settings.SNEntry = tk.Entry(settings, width="10")
        settings.SNEntry.insert(0, "Serial #")
        settings.SNEntry.place(x=3, y=20)
        # Enter gas composition
        settings.gasEntry = tk.Entry(settings, width="26")
        settings.gasEntry.insert(0, "Gas Type (1% O2 Bal N2, etc.)")
        settings.gasEntry.place(x=90, y=20)
        # Add button
        settings.addCalGasButton = tk.Button(settings, text="ADD", command=add_To_Cal_Gases)
        settings.addCalGasButton.place(x=285, y=9)

        ### Provide functionality for removing a calibration standard cylinder
        settings.removeCalGas = tk.Label(settings, text="Remove Calibration Standard").place(x=0, y=50)

        gases = self.fetch_Serial_Numbers(gasDictFile)
        
        settings.variable3 = tk.StringVar()
        settings.calStandard = tk.OptionMenu(settings, settings.variable3, *gases)
        settings.calStandard.place(x=0, y=75)
        # Remove gas button
        settings.removeCalGasButton = tk.Button(settings, text="REMOVE", command=remove_Gas).place(x=110, y=75)

    def open_About_Window(self):
        about = tk.Toplevel(self)

        ### Provide functionality for removing a calibration standard cylinder
        about.info = tk.Text(about, height=12, width=80)
        about.info.pack()
        about.info.insert(tk.END, "The RSD Tool is designed to assist in tasks related to quality control \nat Med Tech Gases. Current functionality includes the following: \n\n-Calculating RSD values\
 given peak areas, taking that data and inserting it into an Excel spreadsheet.\n-Settings allows the user to manage the list of calibration standards. \n\nFuture versions will add additional functionality. \
\n\nVersion 1.0 \nApril 2018 \nCreated by James Gibson")
        about.info.config(fg="white", bg="black")
        

    def fetch_Serial_Numbers(self, textFile):
       
        # Empty list for holding serial numbers
        sNs = []

        # Open file
        fileObj = open(textFile, mode = "r+")

        # Get each row of the file in an array, calGases
        calGases = fileObj.readlines()

        # For each item in the array, retrive the serial number
        for items in calGases:
            items.strip()
            gasSN, gasMakeup = items.split(':')
            sNs.append(gasSN)

        fileObj.close()

        return sNs

    def fetch_Gas_Dict(self, textFile):
   
        sNs = []
        compositions = []

        fileObj = open(textFile, mode = "r+")

        calGases = fileObj.readlines()

        for items in calGases:
            items.strip()
            gasSN, gasMakeup = items.split(':')
            gasMakeup = gasMakeup.strip()
            sNs.append(gasSN)
            compositions.append(gasMakeup)

        dictionary = dict(zip(sNs, compositions))

        fileObj.close()
        
        return dictionary

root = tk.Tk()
app = Application(master=root)
app.mainloop()
