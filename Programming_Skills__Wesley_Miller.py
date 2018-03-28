#!/usr/bin/python

#Coded by Wesley Miller

from Tkinter import *                                                                                                                             #Import TKinter,ttk, sqlite and the dependencies used
from ttk import Treeview, Combobox, Scrollbar
from sqlite3 import Error

import Tkinter, tkFileDialog, tkMessageBox
import ttk
import sqlite3
import os
import csv
import datetime


class wine_DB_manager:                                                                                                                             #Class for wrapping TKinter widgets
    
    def __init__(self, master):                                                                                                                    #Initilize the attributes of the class

        master.minsize(width = 600, height = 700)                                                                                                  #Set window size and title
        master.title("Member and Store Manager")

        master.grid_columnconfigure(0, weight = 1)                                                                                                 #Configure first and last columns for centering columns 1-5
        master.grid_columnconfigure(6, weight = 1)
        master.grid_rowconfigure(12, weight = 1)                                                                                                   #Configure last row for filling left over space at the bottom

        self.imp_store_label = Label(master, text = "Import Store Information (.CSV)")                                                             #Initilize lable, entry, and buttons for importing stores
        self.imp_store_entry = Entry(master, width = 70)
        self.imp_store_browse_button = Button(master, text = "Browse", command = lambda: self.browse_for_CSV("store"))                             
        self.imp_store_import_button = Button(master, text = "Import", command = lambda: self.import_to_DB(self.imp_store_entry.get(),"store"))    

        self.imp_member_label = Label(master, text = "Import Member Information (.CSV)")                                                           #Initilize lable, entry, and buttons for importing members
        self.imp_member_entry = Entry(master, width= 70)
        self.imp_member_browse_button = Button(master, text = "Browse", command = lambda: self.browse_for_CSV("member"))                           
        self.imp_member_import_button = Button(master, text = "Import", command = lambda: self.import_to_DB(self.imp_member_entry.get(),"member")) 

        self.horz_divider = Frame(master, height = 1, width = 500, bg = "black")                                                                   #Initilize a frame shaped as a horizonal line as a divider

        self.view_all_members_button = Button(master, text = "View All Members", width = 17, command = lambda: self.select_all_members())          #Initilize button for selecting all members and displaying them in tree

        self.zip_label = Label(master, text = "Zip Code")                                                                                          #Initilize lable, entry, combobox, and buttons for 
        self.zip_entry = Entry(master, width = 10)                                                                                                 #finding who paid dues by zip code and month paid
        self.month_label = Label(master, text = "Month")
        self.month_combobox = Combobox(master, width = 10)
        self.month_combobox['values'] = ('January', 'February', 'March', 'April', 'May',
                                        'June', 'July', 'August', 'September', 'October', 'November', 'December')
        self.find_users_that_paid = Button(master, text = "Current Paid Users", width = 20, 
                                           command = lambda: self.select_paid_dues_month(self.zip_entry.get(), self.month_combobox.current()))

        self.state_label = Label(master, text = "State")                                                                                           #Initilize lable, entry, combobox, and buttons for
        self.state_combobox = Combobox(master, width = 3)                                                                                          #finding users who have joined the club since a date
        self.state_combobox['values'] = ('MD','NC','PA','VA','WV')                                                                                 #and belong to a specific state
        self.date_label = Label(master, text = "Date (YYYY-MM-DD)")
        self.date_entry = Entry(master, width = 10)
        self.find_users_that_joined = Button(master, text = "Users Joined Since", width = 20,
                                             command = lambda: self.select_users_joined_since(self.state_combobox.get(), self.date_entry.get()))

        self.users_that_love_total_wine = Button(master, text = "Users that Love Total Wine", width = 20,                                           
                                                 command = lambda: self.select_users_love_Tot_Wine())                                               #Initilize button for finding users that love Total Wines

        self.users_favorite_stores = Button(master, text = "User's Favorite Stores", width = 20,
                                            command = lambda: self.select_users_fav_stores())                                                       #Initilize button for finding users, their favorite stores, and store locations

        self.table_tree = Treeview(master, selectmode="extended")                                                                                   #Initilize tree for data viewing
        self.table_tree["columns"] = ("one", "two","three","four","five","six","seven","eight","nine","ten","eleven")                               #Provide max column count and identifiers
        for columns in self.table_tree["columns"]:                                                                                                  #For loop to add all columns
            self.table_tree.column(columns, width=70, anchor=W)
        self.table_tree['show'] = 'headings'                                                                                                        #Remove empty identity column

        self.vert_scroll_bar = Scrollbar(orient="vertical",command=self.table_tree.yview)                                                           #Initilize scroll bar
        self.table_tree.configure(yscrollcommand=self.vert_scroll_bar.set)                                                                          #Add scroll bar to table_tree 



        self.imp_store_label.grid(sticky = "W", row = 1, column = 1, padx = 10, pady = (20,0))                                                      #Grid positioning for all initialized attributes
        self.imp_store_entry.grid(sticky = "W", row = 2, column = 1, columnspan = 3, padx = 10)
        self.imp_store_browse_button.grid(row = 2, column = 4, padx = 10)
        self.imp_store_import_button.grid(row = 2, column = 5)

        self.imp_member_label.grid(sticky = "W", row = 3, column = 1, padx = 10)
        self.imp_member_entry.grid(sticky = "W", row = 4, column = 1, columnspan = 3, padx = 10)
        self.imp_member_browse_button.grid(row = 4, column = 4, padx = 10)
        self.imp_member_import_button.grid(row = 4, column = 5)

        self.horz_divider.grid(row = 5, column = 0, columnspan = 7, pady = 20)

        self.view_all_members_button.grid(row = 6, column = 1, columnspan = 5)

        self.zip_label.grid(sticky = "W", row = 7, column = 1, pady = (15,0))
        self.zip_entry.grid(sticky = "W", row = 8, column = 1, pady = 5)
        self.month_label.grid(sticky = "W", row = 7, column = 1, padx = (100,0), pady = (15,0))
        self.month_combobox.grid(sticky = "W", row = 8, column = 1, padx = (100,0), pady = 5)
        self.find_users_that_paid.grid(sticky = "W",row = 9, column = 1, columnspan = 3, pady = 5)

        self.state_label.grid(sticky = "W", row = 7, column = 4, pady = (15,0))
        self.state_combobox.grid(sticky = "W", row = 8, column = 4, pady = 5)
        self.date_label.grid(sticky = "W", row = 7, column = 4, columnspan = 3, padx = (85,0), pady = (15,0))
        self.date_entry.grid(sticky = "W", row = 8, column = 4, columnspan = 2, padx = (85,0), pady = 5)
        self.find_users_that_joined.grid(sticky = "W",row = 9, column = 4, columnspan = 2, pady = 5)

        self.users_that_love_total_wine.grid(sticky = "W",row = 10, column = 1, columnspan = 3, pady = (15,0))

        self.users_favorite_stores.grid(sticky = "W",row = 10, column = 4, columnspan = 2, pady = (15,0))

        self.table_tree.grid(sticky = "NWES", row = 11, column = 0, columnspan = 7, rowspan = 2, pady = 10, padx = (10,0))
        self.vert_scroll_bar.grid(sticky = "NWES", row = 11, column = 7, rowspan = 2, pady = 10, padx = (0,10))


    def browse_for_CSV(self, CSV_type):                                                                                                            #Class method for browse buttons. Used for passing file path to TKinter entries

        file = tkFileDialog.askopenfile(parent = root, mode = 'rb', title = 'Choose a file')                                                       #Opens browse for file window 

        if (file != None):                                                                                                                         #If file exists read into data and close file

            data = file.read()
            file.close()

            if(CSV_type == "store"):                                                                                                               #In order to resuse code, this method works for both buttons 
                self.imp_store_entry.delete(0, END)                                                                                                #through using a passed button identity variable
                self.imp_store_entry.insert(0, os.path.abspath(file.name))

            else:
                self.imp_member_entry.delete(0, END)                                                                                               #Empties entry widget
                self.imp_member_entry.insert(0, os.path.abspath(file.name))                                                                        #Inserts file path into entry widget using os.path import
        else:                                                                                                                                      #Catches no file selected possibility
            tkMessageBox.showinfo("Error", "No File Selected")

        return None


    def create_DB_connection(self):                                                                                                                #Class method for opening a database connection to the pythonsqlite.db file

        try:
            DB_file = "SQLite_db\pythonsqlite.db"
            conn = sqlite3.connect(DB_file)                                                                                                         
            return conn

        except Error as e:                                                                                                                         #Catches non-connectivity errors
            print(e)
 
        return None


    def import_to_DB(self, file_path, CSV_type):                                                                                                   #Class method for import buttons. Used to open csv files from path string
                                                                                                                                                   #in entry widgets. Then opens db connection to import data to db
        try:                                                                                                                                       
            self.CSV_file = open(file_path, "r")                                                                                                   #Opens CSV file in read mode
        except IOError as e:                                                                                                                       #Catches file not found error
            tkMessageBox.showinfo("Error", "File Not Found")
            return

        CSV_reader = csv.reader(self.CSV_file)                                                                                                     #Reads CSV file into CSV_reader using csv.reader import 
        conn = self.create_DB_connection()                                                                                                         #Calls for DB connection to open
    
        cur = conn.cursor()

        if (CSV_type == "store"):                                                                                                                  #In order to resuse code, this method works for both buttons by passed type.
            if(next(CSV_reader)[0][0:8] != "Store id"):                                                                                            #Checks CSV for proper store headings
                tkMessageBox.showinfo("CSV Type Error", "Please Import a CSV file formated for store data.")
                self.CSV_file.close()                                                                                                              #If not proper headings close file and display message
                return
            else:                                                                                                                                  #Create table Stores
                cur.execute('''CREATE TABLE IF NOT EXISTS Stores (                                                                                 
                                store_id INTEGER PRIMARY KEY, 
                                store_name VARCHAR, 
                                location VARCHAR)''')

                for row in CSV_reader:                                                                                                             #Insert new values for each row in CSV
                    cur.execute('''INSERT OR IGNORE INTO Stores (                                                                                  
                                    store_id, 
                                    store_name, 
                                    location) 
                                    VALUES (?, ?, ?)''', row)
                tkMessageBox.showinfo("Success!", "Successfully Updated the Database")                                                             #Display confirmation
                self.imp_store_entry.delete(0,END)                                                                                                 #Clear import entry

        else:
            if(next(CSV_reader)[0][0:6] != "Member"):                                                                                              #Checks CSV for proper member headings
                tkMessageBox.showinfo("CSV Type Error", "Please Import a CSV file formated for member data.")
                self.CSV_file.close()                                                                                                              #If not proper headings close file and display message
                return
            else:                                                                                                                                  #Create table Members
                cur.execute('''CREATE TABLE IF NOT EXISTS Members (     
                                member_id INTEGER PRIMARY KEY, 
                                last_name VARCHAR, 
                                first_name VARCHAR, 
                                street VARCHAR, 
                                city VARCHAR, 
                                state VARCHAR, 
                                zip VARCHAR, 
                                phone VARCHAR, 
                                favorite_store INTEGER, 
                                date_joined DATE, 
                                dues_paid DATE, 
                                FOREIGN KEY(favorite_store) 
                                REFERENCES Stores(store_id))''')

                for row in CSV_reader:                                                                                                             #Insert new values for each row in CSV    
                    cur.execute('''INSERT OR IGNORE INTO Members (
                                        member_id, 
                                        last_name, 
                                        first_name, 
                                        street, 
                                        city , 
                                        state, 
                                        zip, 
                                        phone, 
                                        favorite_store, 
                                        date_joined, 
                                        dues_paid) 
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''', row)
                tkMessageBox.showinfo("Success!", "Successfully Updated the Database")
                self.imp_member_entry.delete(0,END)

        self.CSV_file.close()                                                                                                                      #Close CSV file
        conn.commit()                                                                                                                              #Commit DB changes
        conn.close()                                                                                                                               #Close DB Connection

        return None


    def generate_member_column_names(self):                                                                                                        #Class utility method for adding member column headers to tree

        all_member_column_names = ("Member ID", "Last Name", "First Name","Street","City","State","Zip Code","Phone","Favorite Store","Date Joined","Dues Paid")

        for index, columns in enumerate(self.table_tree["columns"]):                                                                               #For loop to add all column headers
            self.table_tree.heading(columns, text = all_member_column_names[index])

        return None


    def select_all_members(self):                                                                                                                  #Class method for View All Members Buttons and outputs new data to tree

        cur = self.create_DB_connection().cursor()                                                                                                 #Calls for DB connection to open 
        cur.execute("SELECT * FROM Members ORDER BY last_name ASC")                                                                                #SQL Statement to SELECT all members and their data.
 
        rows = cur.fetchall()                                                                                                                      #Fetches data into rows

        self.table_tree.delete(*self.table_tree.get_children())                                                                                    #Deletes old data in tree if present

        self.generate_member_column_names()                                                                                                        #Call utility to generate column member headers

        for index, row in enumerate(rows):                                                                                                         #Inserts sql data into tree rows for viewing
            self.table_tree.insert('', index, values = (row))

        self.create_DB_connection().close()                                                                                                        #Close DB Connection
        return None


    def select_paid_dues_month(self, zip, month):                                                                                                  #Class method for Current Paid User Button and outputs new data to tree

        month = month+ 1                                                                                                                           #Add 1 to month index count since it range is 0-11 and months are 1-12
        if (month < 10):
            self.month_string= "0" + str(month)

        cur = self.create_DB_connection().cursor()                                                                                                 #Calls for DB connection to open

        if (month not in range(1,12)):                                                                                                             #If no month selected display message
            tkMessageBox.showinfo("Error", "No Month Selected")
            return

        if (len(zip)!=5):                                                                                                                          #If zip not 5 digits display message
            tkMessageBox.showinfo("Error", "Enter a Zip Code")
            return

        try:
            cur.execute("SELECT * FROM Members WHERE zip = "+zip+" AND strftime('%m',dues_paid) = '"+self.month_string+"'")                        #If zip is numeric execute Select Statement
        except Error as e:                                                                                                                         #Catches error if zip not numeric
            tkMessageBox.showinfo("Error", "Enter a Zip Code")
            return

        rows = cur.fetchall()

        self.table_tree.delete(*self.table_tree.get_children())                                                                                    #Deletes old data in tree if present

        self.generate_member_column_names()                                                                                                        #Call utility to generate column member headers

        for index, row in enumerate(rows):                                                                                                         #Inserts sql data into tree rows for viewing
            self.table_tree.insert('', index, values = (row))

        self.create_DB_connection().close()                                                                                                        #Close DB Connection
        return None


    def select_users_joined_since(self, state, the_date):                                                                                          #Class method for Current Paid User Button and outputs new data to tree

        cur = self.create_DB_connection().cursor()                                                                                                 #Calls for DB connection to open
        
        if state == "":                                                                                                                            #if no state selected display message        
            tkMessageBox.showinfo("Error", "No State Selected")
            return

        try:                                                                                                                                       #Checks date format YYYY-MM-DD 
            datetime.datetime.strptime(the_date, '%Y-%m-%d')
        except ValueError:                                                                                                                         #Catches invalid date format
            tkMessageBox.showinfo("Error", "Incorrect data format, should be YYYY-MM-DD")
            return

        cur.execute("SELECT * FROM Members WHERE state = '"+state+"' AND date_joined > '"+the_date+"'")                                            #Execute Select members Where state and date are what user selected.

        rows = cur.fetchall()

        self.table_tree.delete(*self.table_tree.get_children())                                                                                    #Deletes old data in tree if present

        self.generate_member_column_names()                                                                                                        #Call utility to generate column member headers

        for index, row in enumerate(rows):                                                                                                         #Inserts sql data into tree rows for viewing
            self.table_tree.insert('', index, values = (row))
        
        self.create_DB_connection().close()                                                                                                        #Close DB Connection
        return None


    def select_users_love_Tot_Wine(self):                                                                                                          #Class method for Users That Love Tot Wine and outputs new data to tree

        cur = self.create_DB_connection().cursor()                                                                                                 #Calls for DB connection to open
                                                                                                                                                   #Execute Select Where, and Join on foreign key values.
        cur.execute('''SELECT Members.last_name,
                            Members.first_name, 
                            Stores.store_name 
                            FROM Members 
                            JOIN Stores 
                            ON Members.favorite_store = Stores.store_id 
                            WHERE favorite_store = '3' ''')
 
        rows = cur.fetchall()

        self.table_tree.delete(*self.table_tree.get_children())                                                                                    #Deletes old data in tree if present

        all_member_column_names = ("Last Name", "First Name","Favorite Store","","","","","","","","")

        for index, columns in enumerate(self.table_tree["columns"]):                                                                               #Generate *custom* column member headers
            self.table_tree.heading(columns, text = all_member_column_names[index])

        for index, row in enumerate(rows):                                                                                                         #Inserts sql data into tree rows for viewing
            self.table_tree.insert('', index, values = (row))

        self.create_DB_connection().close()                                                                                                        #Close DB Connection
        return None


    def select_users_fav_stores(self):                                                                                                             #Class method for Users Favorite Stores and outputs new data to tree

        cur = self.create_DB_connection().cursor()                                                                                                 #Calls for DB connection to open
                                                                                                                                                   #Execute Select Where, and Join on foreign key values.
        cur.execute('''SELECT Members.last_name,
                    Members.first_name, 
                    Stores.store_name, 
                    Stores.location
                    FROM Members 
                    JOIN Stores 
                    ON Members.favorite_store = Stores.store_id ''')
 
        rows = cur.fetchall()

        self.table_tree.delete(*self.table_tree.get_children())                                                                                    #Deletes old data in tree if present

        all_member_column_names = ("Last Name", "First Name", "Favorite Store", "Location","","","","","","","")

        for index, columns in enumerate(self.table_tree["columns"]):                                                                               #Generate *custom* column member headers
            self.table_tree.heading(columns, text = all_member_column_names[index])

        for index, row in enumerate(rows):                                                                                                         #Inserts sql data into tree rows for viewing
            self.table_tree.insert('', index, values = (row))
        
        self.create_DB_connection().close()                                                                                                        #Close DB Connection
        return None


root = Tkinter.Tk()                                                                                                                                #Create Main(master) Window                                                              

app = wine_DB_manager(root)                                                                                                                        #Create instance of wine_DB_manager class

root.mainloop()                                                                                                                                    #Infinite loop to run GUI