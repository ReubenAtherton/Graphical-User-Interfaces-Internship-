#####################################################################################
#####                   GUI FOR GENERIC TESTING RESULTS SHEET                   #####
#####               REUBEN ATHERTON - SOFTWARE ENGINEERING INTERN               #####
#####                                15/08/2022                                 #####
#####################################################################################

from tkinter import *
from pathlib import Path 
from PIL import ImageTk, Image
from tkinter import messagebox

import pyodbc 
import warnings
warnings.filterwarnings('ignore')

from tkinter import ttk
import tkinter.font as tkFont

# Enter DataBase connection information
conn = pyodbc.connect('Driver={SQL SERVER};'
                      'Server=SQLEXPRESS;' 
                      'Database=ProjectDATABASE;'
                      'Trusted_Connection=yes;')

cursor = conn.cursor()

root = Tk()
root.state("zoomed")
root.resizable(False, True)
root.title("Database")

# Create minimum window size 
root.minsize(1585,1020)

def header_config(root):

    global Header
    Header=Frame(root)

    #Display Header
    Header.pack()
    for i in range (0, 3):
        Header.grid_rowconfigure(i, weight=1)

    # Set each column to equal width weight of 1
    for w in range(4):
        Grid.columnconfigure(Header, w, weight=1)

# --PAGE CONFIGURATION--
def page_1_config(root):

    global Page1
    Page1=Frame(root)

    # Page 1 rows
    for Page1_rows in range (0, 40):
        Grid.rowconfigure(Page1, Page1_rows, weight=1)

    # Page 1 columns
    for Page1_cols in range(4):
        Grid.columnconfigure(Page1, Page1_cols, weight=1)

    #Display Page1
    Page1.pack(side="bottom")

def page_2_config(root):

    global Page2
    Page2=Frame(root)

    # Page 2 rows
    for Page2_rows in range (0, 40):
        Grid.rowconfigure(Page2, Page2_rows, weight=1)
    # Page 2 rows
    for Page2_cols in range(4):
        Grid.columnconfigure(Page2, Page2_cols, weight=1)

# Displays error message if not all fields are full
def show_error_message():
    messagebox.showinfo('Information', 'All fields marked * are required before submitting ')

# Displays success message if minimum fields are full
def show_successful_message():
    conn.commit() 
    messagebox.showinfo('Information', 'Submission Successful')

# Submit button functionality
def submit():

    # Checks that the required fields are full before submission
    if len(Header4.get()) > 0 and len(EquipmentID_info_label.get()) > 0 and len(Calibration_Due_Date_info_label.get()) > 0 and len(Technician_Name_info_label.get()) > 0 and len(Date_info_label1.get()) > 0:

        # Writes a query from Python to insert user entries into Header
        cursor.execute("""INSERT INTO [Header]
                            VALUES (?,?,?,?,?,?)""", Header1.get(), Header2.get(), Header3.get(), Header4.get(), Header5.get(), Header6.get()).rowcount
        
        # Test Information Table Entries
        cursor.execute("""INSERT INTO [Test_Information_Table] (Equipment_ID, Calibration_Date, Header4)
                            VALUES (?,?,?)""", EquipmentID_info_label.get(), Calibration_Due_Date_info_label.get(), Header4.get()).rowcount

        # Retrieves auto generated Test_ID value from SQL table and inserts into other tables for reference
        cursor.execute("SELECT Test_ID FROM Test_Information_Table")
        Test_ID = cursor.fetchall()[-1][0] # finds the lastest ID number by getting the last element in the Table_ID column 
        print("Current Test ID: ", Test_ID) # Useful to find whether has been any unpassed queries that have incremented the user ID

        # Inserting each element from the 2D Results Table Array
        for j in range(13):
            Result_ID = j+1

            # Assigns each column to Leg_Pin, Test_Voltage or IR -- included to improve organisation 
            LP = Results_Table_Input[j][0]
            TV = Results_Table_Input[j][1]
            IR = Results_Table_Input[j][2]

            # When j==12 the LP value must be NULL because this is auto filled on the sheet and so the entry is missed to fill element as NULL
            # **  data type is int for this column therefore printing the text here was not possible **
            if j == 12:
                cursor.execute("""INSERT INTO [Results_Table] (Test_ID, Result_ID, Test_Voltage, IR) 
                                    VALUES (?,?,?,?)""", Test_ID, Result_ID, TV.get(), IR.get()).rowcount
            else:
                cursor.execute("INSERT INTO [Results_Table] VALUES (?,?,?,?,?)", Test_ID, Result_ID, LP.get(), TV.get(), IR.get()).rowcount

        cursor.execute("""INSERT INTO [Test_Environment]
                            VALUES (?,?,?,?)""", Test_ID, Temperature_Info_label.get(), Humidity_Info_label.get(), Barometric_Pressure_Info.get()).rowcount
        
        cursor.execute("""INSERT INTO [Operator_Table] 
                            VALUES (?,?,?,?,?,?)""", Test_ID, Technician_Name_info_label.get(), Date_info_label1.get(), Third_Party_Witness_info_label.get(), Date_info_label2.get(), AppendixB_info_label.get()).rowcount
        
        # Signals when entries have been queried correctly
        print("Record inserted successfully")

        # Clearing the user entry boxes
        Header1.delete(0, END)
        Header2.delete(0, END)
        Header3.delete(0, END)
        Header4.delete(0, END)
        Header5.delete(0, END)
        Header6.delete(0, END)

        EquipmentID_info_label.delete(0, END)
        Calibration_Due_Date_info_label.delete(0, END)
        Temperature_Info_label.delete(0, END)
        Humidity_Info_label.delete(0, END)
        Barometric_Pressure_Info.delete(0, END)

        # Parses Array and deletes contents
        for j in range(13): # rows 
            for i in range(3): # columns 
                Results_Table_Input[j][i].delete(0, END)

        Technician_Name_info_label.delete(0, END)
        Date_info_label1.delete(0, END)
        Third_Party_Witness_info_label.delete(0, END)
        Date_info_label2.delete(0, END)
        AppendixB_info_label.delete(0, END)

        # Display 'Successful Submission' message
        show_successful_message()

        # There will still be errors for example when user enters wrong data type 
        # and there wil be no error message displayed. Needs fixing 
        
    else:
        # Display Submission Error
        show_error_message()

# CONFIG HEADER AND PAGES 

header_config(root)
page_1_config(root)
page_2_config(root)

#########################################
############  PAGE HEADER  ##############
#########################################

# --Initialising Labels--

# Company Logo - enter logo path
img = ImageTk.PhotoImage(Image.open("C:___PATH____final3_logo.png"))
logo_label = Label(Header,image=img)

#Title
title_label = Label(Header, text="TEST RESULTS SHEET \n FOR COMPANY PRODUCT", font=('Helvetica', 18, 'bold'))

# --- CATEGORY LABELS ---

# Category 1 Label and Text Info
CategoryLabel1 = Label(Header, text="Category Label 1", relief=GROOVE, font=('bold', 14))
CategoryLabel1_info = Label(Header, text="TEXT", relief=GROOVE, font=('bold', 14))

# Category 2 Label and Text nfo
CategoryLabel2 = Label(Header, text="Category Label 2", relief=GROOVE, font=('bold', 14))
CategoryLabel2_info = Label(Header, text="TEXT", relief=GROOVE, font=('bold', 14))

# Category 3 Label and Text info
CategoryLabel3 = Label(Header, text="Category Label 3", relief=GROOVE, font=('bold', 14))
CategoryLabel3_info = Label(Header, text="TEXT", relief=GROOVE, font=('bold', 14))

# Category 4 Label and Text Info
CategoryLabel4 = Label(Header, text="Category Label 4", relief=GROOVE, font=('bold', 14))
CategoryLabel4_info = Label(Header, text="TEXT", relief=GROOVE, font=('bold', 14))

# --- HEADER LABELS ---

# Header 1 Label and Entry Label
HeaderLabel1 = Label(Header, text="Header Label 1", anchor="w", bg = "light blue", font=('bold', 14), relief=GROOVE)
Header1 = Entry(Header, justify=CENTER, font=('Helvetica', 13))

# Header 2 Label and Entry Label
HeaderLabel2 = Label(Header, text="Header Label 2", anchor="w", bg = "light blue", font=('bold', 14), relief=GROOVE)
Header2 = Entry(Header, justify=CENTER, font=('Helvetica', 13))

# Header 3 Label and Entry Label
HeaderLabel3 = Label(Header, text="Header Label 3", anchor="w", bg = "light blue", font=('bold', 14), relief=GROOVE)
Header3 = Entry(Header, justify=CENTER, font=('Helvetica', 13))

# Header 4 Label and Entry Label
HeaderLabel4 = Label(Header, text="Header Label 4*", anchor="w", bg = "light blue", font=('bold', 14), relief=GROOVE)
Header4 = Entry(Header, justify=CENTER, font=('Helvetica', 13)) 

# Header 5 Label and Entry Label
HeaderLabel5 = Label(Header, text="Header Label 5", anchor="w", bg = "light blue", font=('bold', 14), relief=GROOVE)
Header5 = Entry(Header, justify=CENTER, font=('Helvetica', 13))

# Header 6 Label and Entry Label
HeaderLabel6 = Label(Header, text="Header Label 6", anchor="w", bg = "light blue", font=('bold', 14), relief=GROOVE)
Header6 = Entry(Header, justify=CENTER, font=('Helvetica', 13))

# Initialising Pages to allow for page switching - attempted to use OOP with different pages contained in different .py files but still needs work
pages = ["Test Page 1",
          "Test Page 2", 
          "Test Page 3",
          "Test Page 4",
          "Test Page 5",
          "Test Page 6"
          ]
          
# Change page after drop down menu event
# Must be a simpler, cleaner way of doing this -- look into methods such as getcurrentpage() so only have to forget current page?
#                                              -- this can be scaled up for when more pages are built and added

def PageChange(event):

    #Page 1
    if variable.get() == pages[0]:
        Page2.pack_forget()
        Page1.pack()

    #Page 2
    if variable.get() == pages[1]:
        Page1.pack_forget()
        Page2.pack()

# Attach the drop down to Header
variable = StringVar(Header)
variable.set(pages[0]) # default value

bigfont = tkFont.Font(family="Helvetica",size=14)
Header.option_add("*TCombobox*Listbox*Font", bigfont)

drop_down_menu = ttk.Combobox(Header, textvariable=variable, values=pages, font=("bold", 14), state="readonly")
drop_down_menu.bind("<<ComboboxSelected>>", PageChange)

# PRINT LABELS TO PAGE WITH .GRID

logo_label.grid(row=0, column= 0, rowspan=4, sticky="nsew")
title_label.grid(row= 0, column=1, rowspan=4, sticky="nsew")

CategoryLabel1.grid(row=0, column=2, sticky="nsew")
CategoryLabel1_info.grid(row=0, column=3, sticky="nsew")
CategoryLabel2.grid(row=1, column=2, sticky="nsew")
CategoryLabel2_info.grid(row=1, column=3, sticky="nsew")
CategoryLabel3.grid(row=2, column=2, sticky="nsew")
CategoryLabel3_info.grid(row=2, column=3, sticky="nsew")
CategoryLabel4.grid(row=3, column=2, sticky="nsew")
CategoryLabel4_info.grid(row=3, column=3, sticky="nsew")

HeaderLabel1.grid(row=4, column=0, sticky="nsew")
Header1.grid(row=4, column=1, sticky="nsew")
HeaderLabel2.grid(row=4, column=2, sticky="nsew")
Header2.grid(row=4, column=3, sticky="nsew")
HeaderLabel3.grid(row=5, column=0, sticky="nsew")
Header3.grid(row=5, column=1, sticky="nsew")
HeaderLabel4.grid(row=5, column=2, sticky="nsew")
Header4.grid(row=5, column=3, sticky="nsew")
HeaderLabel5.grid(row=6, column=0, sticky="nsew")
Header5.grid(row=6, column=1, sticky="nsew")
HeaderLabel6.grid(row=6, column=2, sticky="nsew")
Header6.grid(row=6, column=3, sticky="nsew")

# Issue with page dimensions - struggled to set each column to an equal size because they always conform to smallest width
# Solution - create 4 seemingly blank columns (also acts as a spacer) with tabs /t that fills entire screen width  
#          - combining this with the equally spaced column parameter set (lines 59 - 61), this forces each column to fit within the screen 
#            limits and maintain equal column widths

for i in range(4):
    spacer2 = Label(Header, text='\t\t\t\t\t\t\t\t\t\t\t\t\t\t').grid(row=7, column=i, sticky="nsew")

drop_down_menu.grid(row=8, column=1, columnspan=2, sticky="nsew")

#########################################
##############   PAGE 1   ###############
#########################################

# --SECTION 1--

#Calibrated Equipment Used Label
Calibrated_equipment_used_label = Label(Page1, text="Calibrated Equipment Used", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)

#Insulation Meter Label
EquipmentUsedLabel = Label(Page1, text="TEXT", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)

#Equipment ID Label and Entry
Equipment_ID_label = Label(Page1, text="Equipment ID *", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)
EquipmentID_info_label = Entry(Page1, justify=CENTER, font=('Helvetica', 13))

#Calibration Due Date Label and Entry
Calibration_Due_Date_label = Label(Page1, text="Calibration Due Date (YYYY-MM-DD) *", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)
Calibration_Due_Date_info_label = Entry(Page1, justify=CENTER, font=('Helvetica', 13))

#Temperature Label and Entry
Temperature_label = Label(Page1, text="Temperature (°C)", bg = "light grey", font=('bold', 14), relief=GROOVE)
Temperature_Info_label = Entry(Page1, justify=CENTER, font=('bold', 13))

#Humidity Label and Entry
Humidity_label = Label(Page1, text="Humidity (%)", bg = "light grey", font=('bold', 14), relief=GROOVE)
Humidity_Info_label = Entry(Page1, justify=CENTER, font=('bold', 13))

#Barometric Pressure Label and Entry
Barometric_Pressure_label = Label(Page1, text="Barometric Pressure (mb)", bg = "light grey", font=('bold', 14), relief=GROOVE)
Barometric_Pressure_Info = Entry(Page1, justify=CENTER, font=('bold', 13))

#Test Voltage Label and Entry
Test_Voltage_label = Label(Page1, text="Test Voltage", bg = "light grey", font=('bold', 14), relief=GROOVE)
Test_Voltage_info_label = Label(Page1, text="TEXT", font=('bold', 14), relief=GROOVE)

#Test Voltage for Instrument Label and Entry
Test_Voltage_Intsrument_label = Label(Page1, text="Test Voltage (for Instrument assemblies)", bg = "light grey", font=('bold', 14), relief=GROOVE)
Test_Voltage_Intrument_info_label = Label(Page1, text="TEXT", font=('bold', 14), relief=GROOVE)

#Test Duration Label and Entry
Test_Duration_label = Label(Page1, text="Test Duration", bg = "light grey",  font=('bold', 14), relief=GROOVE)
Test_Duration_info_label = Label(Page1, text="TEXT", font=('bold', 14), relief=GROOVE)

# PRINTING LABELS TO PAGE - printing them all in same place makes it easy to visualise and fix any errors

Calibrated_equipment_used_label.grid(row=9, column=0, columnspan=2, sticky="nsew")
EquipmentUsedLabel.grid(row=10, column=0, columnspan=2, sticky="nsew")
Equipment_ID_label.grid(row=9, column=2, sticky="nsew")
EquipmentID_info_label.grid(row=10, column=2, sticky="nsew")
Calibration_Due_Date_label.grid(row=9, column=3, sticky="nsew")
Calibration_Due_Date_info_label.grid(row=10, column=3, sticky="nsew")

spacer3 = Label(Page1, text='').grid(row=11, column=0)

Temperature_label.grid(row=12, column=0, columnspan=2, sticky="nsew")
Temperature_Info_label.grid(row=12, column=2, columnspan=4, sticky="nsew")
Humidity_label.grid(row=13, column=0, columnspan=2, sticky="nsew")
Humidity_Info_label.grid(row=13, column=2, columnspan=4, sticky="nsew")
Barometric_Pressure_label.grid(row=14, column=0, columnspan=2, sticky="nsew")
Barometric_Pressure_Info.grid(row=14, column=2, columnspan=4, sticky="nsew")

spacer4 = Label(Page1, text='').grid(row=15, column=0, sticky="nsew")

Test_Voltage_label.grid(row=16, column=0, columnspan=2, sticky="nsew")
Test_Voltage_info_label.grid(row=16, column=2, columnspan=4, sticky="nsew")
Test_Voltage_Intsrument_label.grid(row=17, column=0, columnspan=2, sticky="nsew")
Test_Voltage_Intrument_info_label.grid(row=17, column=2, columnspan=4, sticky="nsew")
Test_Duration_label.grid(row=18, column=0, columnspan=2, sticky="nsew")
Test_Duration_info_label.grid(row=18, column=2, columnspan=4, sticky="nsew")

spacer5 = Label(Page1, text='\t\t\t\t\t\t\t\t\t\t\t\t\t\t').grid(row=19, column=0, sticky="nsew")

for i in range(4):
    col_spacer1 = Label(Page1, text='\t\t\t\t\t\t\t\t\t\t\t\t\t\t').grid(row=19, column=i, sticky="nsew")

# --RESULTS TABLE SECTION 2--

#Leg, Pin  Header
Leg_Pin_label = Label(Page1, text="Leg, Pin", bg = "light grey",  font=('bold', 14), relief=GROOVE)
Leg_Pin_label.grid(row=20, column=0, sticky="nsew")

# Test Voltage Header
Test_Voltage_label2 = Label(Page1, text="Test Voltage", bg = "light grey", font=('bold', 14), relief=GROOVE)
Test_Voltage_label2.grid(row=20, column=1, sticky="nsew")

#IR Value Header
IR_label = Label(Page1, text="IR", bg = "light grey", font=('bold', 14), relief=GROOVE)
IR_label.grid(row=20, column=2, sticky="nsew")

#Pass Criteria Header
Pass_Criteria_label = Label(Page1, text="Pass Criteria", bg = "light grey", font=('bold', 14), relief=GROOVE) # Tabs included to achieve central page divide
Pass_Criteria_label.grid(row=20, column=3, sticky="nsew")

# Results Table User Entries (3x12 array)
rows, cols = 13, 3
#[['' for i in range(cols)] for j in range(rows)]
Results_Table_Input = [['' for i in range(cols)] for j in range(rows)]

# Array created for the table. Each entry is assigned an element within the array which can then 
# be later accessed for sending to database. Array structure is as follows:
# Results_Table_Input = [[Leg_Pin Values (0-12).....], [Test_Voltage Values (0-12).....], [IR (0-12).....]]

#       | LEG_PIN | TEST_VOLTAGE | IR |
# 0     |_________|______________|____|
# 1     |_________|______________|____|
# 2     |_________|______________|____|
# 3     |_________|______________|____|
# 4     |_________|______________|____|
# ...   |_________|______________|____|
# 12    |_________|______________|____|

# --GENERATING THE RESULTS TABLE USING A 'FOR LOOP'--
for row in range(12):
    for col in range(3):
        Results_Table_Input[row][col] = Entry(Page1, justify=CENTER, font=('bold', 12))
        Results_Table_Input[row][col].grid(row=row+21, column=col, sticky="nsew")

# Row 12 requires more detail due to NULL requirement in element [0][12]
All_to_Earth = Label(Page1, text="--N/A--", font=('bold', 12), bg="white", relief=GROOVE)
All_to_Earth.grid(row=33, column=0, sticky="nsew")

# Not shown on grid and must always return a NULL value because isn't actually an entry 
Results_Table_Input[12][0] = Entry(Page1, justify=CENTER, font=('bold', 12)) 

Results_Table_Input[12][1] = Entry(Page1, justify=CENTER, font=('bold', 12))
Results_Table_Input[12][1].grid(row=33, column=1, sticky="nsew")

Results_Table_Input[12][2] = Entry(Page1, justify=CENTER, font=('bold', 12))
Results_Table_Input[12][2].grid(row=33, column=2, sticky="nsew")

# Pass Criteria Labels
PC_1 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_1.grid(row=21, column=3, sticky="nsew")

PC_2 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_2.grid(row=22, column=3, sticky="nsew")

PC_3 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_3.grid(row=23, column=3, sticky="nsew")

PC_5 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_5.grid(row=25, column=3, sticky="nsew")

PC_7 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_7.grid(row=27, column=3, sticky="nsew")

PC_9 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_9.grid(row=29, column=3, sticky="nsew")

PC_11 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_11.grid(row=31, column=3, sticky="nsew")

PC_13 = Label(Page1, text="TEXT", anchor="w", font=('bold', 12), relief=GROOVE)
PC_13.grid(row=33, column=3, sticky="nsew")

spacer6 = Label(Page1, text='')
spacer6.grid(row=34, column=0)

########## SECTION 3 ##########

#Technician Name Label and Entry
Technician_Name_label = Label(Page1, text="Technician Name *", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Technician_Name_info_label = Entry(Page1, justify=CENTER, font=('bold', 13))

# 3rd Party Witness Label and Entry
Third_Party_Witness_label = Label(Page1, text="3rd Party Witness", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Third_Party_Witness_info_label = Entry(Page1, justify=CENTER, font=('bold', 13))

# Date (Technician Name) Label and Entry
Date_label1 = Label(Page1, text="Date (YYYY-MM-DD) *", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Date_info_label1 = Entry(Page1, justify=CENTER, font=('bold', 13))

# Date (3rd Party Witness) Label and Entry
Date_label2 = Label(Page1, text="Date (YYYY-MM-DD)", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Date_info_label2 = Entry(Page1, justify=CENTER, font=('bold', 13))

spacer7 = Label(Page1, text='')

#Apppendix Note Label and Entry
Appendix_Note_label = Label(Page1, text="• Please use Appendix B if more results require to be recorded.", font=('Helvetica', 13))

spacer8 = Label(Page1, text='')

#Appendix B Used? Label and Entry
AppendixB_label = Label(Page1, text="Appendix B Used? (Y/N)", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
AppendixB_info_label = Entry(Page1, justify=CENTER, font=('bold', 13))

# Submit Button 
submit_button = Button(Page1, text="Submit", command=submit, bg = "light blue", font=('Helvetica', 11, 'bold'))

# Displaying Section 3
Technician_Name_label.grid(row=35, column=0, sticky="nsew")
Technician_Name_info_label.grid(row=35, column=1, sticky="nsew")
Third_Party_Witness_label.grid(row=35, column=2, sticky="nsew")
Third_Party_Witness_info_label.grid(row=35, column=3, sticky="nsew")
Date_label1.grid(row=36, column=0, sticky="nsew")
Date_info_label1.grid(row=36, column=1, sticky="nsew")
Date_label2.grid(row=36, column=2, sticky="nsew")
Date_info_label2.grid(row=36, column=3, sticky="nsew")

spacer7.grid(row=37, column=0, columnspan=4, sticky="nsew")
Appendix_Note_label.grid(row=38, column=0, columnspan=2, sticky="nsew")
spacer8.grid(row=39, column=0)

AppendixB_label.grid(row=38, column=2, sticky="nsew")
AppendixB_info_label.grid(row=38, column=3, sticky="nsew")
submit_button.grid(row=39, column=3, sticky="nsew")

################################################################################################################################################################################ PAGE 2

## Page 2

# Submit button functionality 
def Page2_submit():
    
    # Checks that the required fields are full before submission
    if len(Header4.get()) > 0 and len(P2_Calibrated_equipment_used_info_label.get()) > 0 and len(P2_EquipmentID_info_label.get()) > 0 and len(P2_Calibration_Due_Date_info_label.get()) > 0 and len(P2_Technician_Name_info_label.get()) > 0 and len(P2_Date_info_label1.get()) > 0:

        # Enters SQL commands that will enter these .get() varibles into the table Header Table 
        cursor.execute("""INSERT INTO [Header] 
                            VALUES (?,?,?,?,?,?)""", Header1.get(), Header2.get(), Header3.get(), Header4.get(), Header5.get(), Header6.get()).rowcount

        # Test Information Table Entries
        cursor.execute("""INSERT INTO [Test_Information_Table] (Equipment_ID, Calibration_Date, Header4)
                            VALUES (?,?,?)""", P2_EquipmentID_info_label.get(), P2_Calibration_Due_Date_info_label.get(), Header4.get()).rowcount

        cursor.execute("SELECT Test_ID FROM Test_Information_Table")

        Test_ID = cursor.fetchall()[-1][0]
        print("2_Current Test ID: ", Test_ID) 

        cursor.execute("""INSERT INTO [Calibrated_Equipment]
                            VALUES (?,?)""", Test_ID, P2_Calibrated_equipment_used_info_label.get()).rowcount

        cursor.execute("""INSERT INTO [Test_Environment]
                            VALUES (?,?,?,?)""", Test_ID, P2_Temperature_Info_label.get(), P2_Humidity_Info_label.get(), P2_Barometric_Pressure_Info.get()).rowcount
        for j in range(8):
        
            P2_Result_ID = j+1

            # Assigns each column to Leg_Pin, Test_Voltage or IR -- included to improve organisation 
            B2B = Continuity_Results_Table[j][0]
            Con = Continuity_Results_Table[j][1]

            cursor.execute("""INSERT INTO [Page_2_Results_Table] (Test_ID, Result_ID, Body_to_Body, Continuity) 
                                VALUES (?,?,?,?)""", Test_ID, P2_Result_ID, B2B.get(), Con.get()).rowcount
        
        #Enters these .get() varibles into the table Test Info Table 
        cursor.execute("""INSERT INTO [Operator_Table] (Test_ID, Technician_Name, Date_1, Third_Party_Witness, Date_2)
                            VALUES (?,?,?,?,?)""", Test_ID, P2_Technician_Name_info_label.get(), P2_Date_info_label1.get(), P2_Third_Party_Witness_info_label.get(), P2_Date_info_label2.get()).rowcount

        Header1.delete(0, END)
        Header2.delete(0, END)
        Header3.delete(0, END)
        Header4.delete(0, END)
        Header5.delete(0, END)
        Header6.delete(0, END)

        P2_Calibrated_equipment_used_info_label.delete(0, END)
        P2_EquipmentID_info_label.delete(0, END)
        P2_Calibration_Due_Date_info_label.delete(0, END)
        P2_Temperature_Info_label.delete(0, END)
        P2_Humidity_Info_label.delete(0, END)
        P2_Barometric_Pressure_Info.delete(0, END)

        for row in range(8):
            for col in range(2):
                # Assigns each column to Leg_Pin, Test_Voltage or IR -- included to improve organisation 
                B2B = Continuity_Results_Table[row][col].delete(0, END)
                Con = Continuity_Results_Table[col][col].delete(0, END)

        P2_Technician_Name_info_label.delete(0, END)
        P2_Date_info_label1.delete(0, END)
        P2_Third_Party_Witness_info_label.delete(0, END)
        P2_Date_info_label2.delete(0, END)

        # Display 'Successful Submission' message
        show_successful_message()
    else:
        # Display Submission Error
        show_error_message()
        
Calibrated_equipment_used_label = Label(Page2, text="Calibrated Equipment Used *", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)

P2_Calibrated_equipment_used_info_label = Entry(Page2, justify =CENTER, font=('Helvetica', 13), relief="ridge")

Equipment_ID_label = Label(Page2, text="Equipment ID *", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)

P2_EquipmentID_info_label = Entry(Page2, justify=CENTER, font=('Helvetica', 13), relief="ridge")

Calibration_Due_Date_label = Label(Page2, text="Calibration Due Date (YYYY-MM-DD) *", bg = "light grey", font=('Helvetica', 14), relief=GROOVE)

P2_Calibration_Due_Date_info_label = Entry(Page2, justify=CENTER, font=('Helvetica', 13), relief="ridge")

#### Grid Layout
Calibrated_equipment_used_label.grid(row=0, column=0, columnspan=2, sticky="nsew")
P2_Calibrated_equipment_used_info_label.grid(row=1, column=0, columnspan=2, sticky="nsew")
Equipment_ID_label.grid(row=0, column=2, sticky="nsew")
P2_EquipmentID_info_label.grid(row=1, column=2, sticky="nsew")
Calibration_Due_Date_label.grid(row=0, column=3, sticky="nsew")
P2_Calibration_Due_Date_info_label.grid(row=1, column=3, sticky="nsew", ipady=2)
###

Page2_spacer1 = Label(Page2, text='').grid(row=2, column=0, sticky="nsew")
Page2_spacer3 = Label(Page2, text='').grid(row=3, column=0, sticky="nsew")

Temperature_label = Label(Page2, text="Temperature (°C)", bg = "light grey", font=('bold', 14), relief=GROOVE)
P2_Temperature_Info_label = Entry(Page2, justify=CENTER, font=('bold', 13))

#Humidity Label and Entry
Humidity_label = Label(Page2, text="Humidity (%)", bg = "light grey", font=('bold', 14), relief=GROOVE)
P2_Humidity_Info_label = Entry(Page2, justify=CENTER, font=('bold', 13))

#Barometric Pressure Label and Entry
Barometric_Pressure_label = Label(Page2, text="Barometric Pressure (mb)", bg = "light grey", font=('bold', 14), relief=GROOVE)
P2_Barometric_Pressure_Info = Entry(Page2, justify=CENTER, font=('bold', 13))

## ## ##
Temperature_label.grid(row=4, column=0, columnspan=2, sticky="nsew")
P2_Temperature_Info_label.grid(row=4, column=2, columnspan=4, sticky="nsew")
Humidity_label.grid(row=5, column=0, columnspan=2, sticky="nsew")
P2_Humidity_Info_label.grid(row=5, column=2, columnspan=4, sticky="nsew")
Barometric_Pressure_label.grid(row=6, column=0, columnspan=2, sticky="nsew")
P2_Barometric_Pressure_Info.grid(row=6, column=2, columnspan=4, sticky="nsew")

# set each column equal width
for i in range (4):
    Page2_col_spacer1 =  Label(Page2, text='\t\t\t\t\t\t\t\t\t\t\t\t\t\t').grid(row=7, column=i, sticky="nsew")

Page2_spacer4 = Label(Page2, text='').grid(row=8, column=0)

# Test Voltage Header
Body_to_body_label = Label(Page2, text="Body to Body / Earth Strap", bg = "light grey", font=('bold', 14), relief=GROOVE)
Body_to_body_label.grid(row=9, column=0, columnspan=2, sticky="nsew")

# IR Value Header
CON_label = Label(Page2, text="Continuity", bg = "light grey", font=('bold', 14), relief=GROOVE)
CON_label.grid(row=9, column=2, sticky="nsew")

# Pass Criteria Header
Pass_Criteria_label = Label(Page2, text="Pass Criteria", bg = "light grey", font=('bold', 14), relief=GROOVE) 
Pass_Criteria_label.grid(row=9, column=3, sticky="nsew")

# USER ENTRIES 

Continuity_rows, Continuity_cols = 8, 2
Continuity_Results_Table =  [['' for i in range(Continuity_cols)] for j in range(Continuity_rows)]

#       |        BODY_TO_BODY     | CONTINUITY | 
# 0     |____________.____________|____________|
# 1     |____________.____________|____________|
# 2     |____________.____________|____________|
# 3     |____________.____________|____________|
# 4     |____________.____________|____________|
# ...   |____________.____________|____________|
# 12    |____________.____________|____________|

# With columnspan this table was harder although the logic is the same within the array initialisation, 
# the Entry(row=..., column=...) are different as shown below

# --GENERATING THE RESULTS TABLE USING A 'FOR LOOP'--

for row in range(Continuity_rows):

    for col in range(Continuity_cols):
        if col == 0:
            Continuity_Results_Table[row][col] = Entry(Page2, justify=CENTER, font=('bold', 12))
            Continuity_Results_Table[row][col].grid(row=row+10, column=col, columnspan=2, sticky="nsew")
        else:
            Continuity_Results_Table[row][col] = Entry(Page2, justify=CENTER, font=('bold', 12))
            Continuity_Results_Table[row][col].grid(row=row+10, column=col+1, sticky="nsew")

        
PC_1 = Label(Page2, text="TEXT", bg="white", font=('bold', 12), relief=GROOVE)
PC_1.grid(row=10, column=3, rowspan=8, sticky="nsew")

Page2_spacer2 = Label(Page2, text='').grid(row=18, column=0)
Page2_spacer5 = Label(Page2, text='').grid(row=19, column=0)

#Technician Name Label and Entry
Technician_Name_label = Label(Page2, text="Technician Name *", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Technician_Name_label.grid(row=20, column=0, sticky="nsew")

P2_Technician_Name_info_label = Entry(Page2, justify=CENTER, font=('bold', 13))
P2_Technician_Name_info_label.grid(row=20, column=1, sticky="nsew")

# 3rd Party Witness Label and Entry
Third_Party_Witness_label = Label(Page2, text="3rd Party Witness (if required - refer to Routing)", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Third_Party_Witness_label.grid(row=20, column=2, sticky="nsew")

P2_Third_Party_Witness_info_label = Entry(Page2, justify=CENTER, font=('bold', 13))
P2_Third_Party_Witness_info_label.grid(row=20, column=3, sticky="nsew")

# Date (Technician Name) Label and Entry
Date_label1 = Label(Page2, text="Date (YYYY-MM-DD) *", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Date_label1.grid(row=21, column=0, sticky="nsew")

P2_Date_info_label1 = Entry(Page2, justify=CENTER, font=('bold', 13))
P2_Date_info_label1.grid(row=21, column=1, sticky="nsew")

# Date (3rd Party Witness) Label and Entry
Date_label2 = Label(Page2, text="Date (YYYY-MM-DD)", anchor="w", font=('bold', 14), bg="light grey", relief=GROOVE)
Date_label2.grid(row=21, column=2, sticky="nsew")

P2_Date_info_label2 = Entry(Page2, justify=CENTER, font=('bold', 13))
P2_Date_info_label2.grid(row=21, column=3, sticky="nsew")

Page2_spacer5= Label(Page2, text='').grid(row=22, column=0)
#Page2_spacer6= Label(Page2, text='').grid(row=23, column=0)

# Submit Button 
P2_submit_button = Button(Page2, text="Submit", command=Page2_submit, bg = "light blue", font=('Helvetica', 11, 'bold'))
P2_submit_button.grid(row=23, column=3, sticky="nsew")

#Run Loop...
root.mainloop()
