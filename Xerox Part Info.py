import gspread
import pandas as pd
from openpyxl import load_workbook
import re
from selenium import webdriver
import os
from selenium.webdriver.common.by import By
import datetime
import time
Zero = True
#print('--------------------------------------------------')
#print('Welcome')
#print('--------------------------------------------------')
#print("1.) Adds the Part Number and other part info ")
#print("2.) Scout Rifles")
#print("3.) Pulse Rifles")

while Zero:

    def Main():
        Menu_2 = input('\n\nPress [1] to add part info\nPress [2] to read the sheet\nPress [3] to edit the product quantity \n-->')



        def Part_Entry():

            run = 0
            print('--------------------------------------------------')
            print('Part Information entry')
            print('--------------------------------------------------')
            sa = gspread.service_account(filename="xerox-part-inventory-6c3999d8969d.json")
            sh = sa.open("Xerox Part Info")
            # Sheet_title = input('Enter Day\n')
            try:
                Part_Number = input('Part Number\n-->')
            except:
                pass
            try:
                Part_Description = input('Part Description\n-->')
            except:
                pass
            try:
                Model_N = input('Printer Model\n-->')
            except:
                pass
            try:
                Quantity = int(input('Quantity\n-->'))
            except:
                pass
            try:
                now = datetime.datetime.now()
                date = now.strftime("%Y-%m-%d %H:%M:%S")
            except:
                pass
            try:
                state = '1st Entry'
            except:
                pass

            print('--------------------------------------------------')

            sa = gspread.service_account(filename="xerox-part-inventory-6c3999d8969d.json")

            sh = sa.open("Xerox Part Info")
            ts = sh.worksheet('Part Catalog')
            values_list = ts.row_values(1)
            values_list2 = ts.col_values(1)
            # print(ts.get_all_records())
            show = ts.get_all_records()


            for up in show:
                # print(a)
                key = list(up.keys())
                value = list(up.values())
                # print(key)
                # print(value)
                ## Validating to see if username and password are in the value variable

                if Part_Number in value:
                    run = 1
                    break

            if run == 0:

                sa = gspread.service_account(filename="xerox-part-inventory-6c3999d8969d.json")
                sh = sa.open("Xerox Part Info")
                # Sheet_title = input('Enter Day\n')

                wks = sh.worksheet('Part Catalog')
                try:
                    df = pd.DataFrame(
                        {'Part Number': [Part_Number], 'Part Description': [Part_Description], 'Model': [Model_N],
                         'Quantity': [Quantity], 'Date of last entry': [date], 'Last Action': [state]})
                except:
                    pass
                try:
                    df_values = df.values.tolist()
                except:
                    pass
                try:
                    sh.values_append('Part Catalog', {'valueInputOption': 'RAW'}, {'values': df_values})
                except:
                    pass

                print('Your entry was entered successfully')

        def Sub_2():
            # read
            sa = gspread.service_account(filename="xerox-part-inventory-6c3999d8969d.json")
            sh = sa.open("Xerox Part Info")
            print('test')
            Sheet_title = str("Part Catalog")

            try:
                wks_2 = sh.worksheet(Sheet_title)
            except:
                pass
            # print(wks_2.get_all_records())
            try:
                for i in wks_2.get_all_records():
                    print(i)
            except:
                pass

        def Sub_3():
            log = True



            while log:
                Part_Number2 = input('Please enter the Part Number\n--> ')


                sa = gspread.service_account(filename="xerox-part-inventory-6c3999d8969d.json")
                sh = sa.open("Xerox Part Info")
                ts = sh.worksheet('Part Catalog')
                values_list = ts.row_values(1)
                values_list2 = ts.col_values(1)
                # print(ts.get_all_records())
                show = ts.get_all_records()
                #print(show)

                for up in show:
                    # print(a)
                    key = list(up.keys())
                    value = list(up.values())
                    #print(value)
                    # print(key)
                    # print(value)
                    ## Validating to see if username and password are in the value variable

                    ##Changing password



                    if Part_Number2 in value:
                        #print('Yes')
                        # print([up])
                        # print(up)
                        usr_ar = list(up.values())
                        #print(usr_ar)
                        usr_zero = usr_ar[0]
                        pas_one = usr_ar[1]
                        Key_e = usr_ar[2]
                        Quantity_Num = usr_ar[3]
                        usr_str = ''.join(usr_zero)
                        pas_str = ''.join(pas_one)
                        #print(Quantity_Num)

                        cell = ts.find(Part_Number2)
                        cell2 = "C%sR%s" % (cell.col + 3, cell.row)
                        cell3 = "C%sR%s" % (cell.col + 4, cell.row)
                        cell4 = "C%sR%s" % (cell.col + 5, cell.row)
                        find = re.findall(r'\d{1,}', cell2)
                        find2 = re.findall(r'\d{1,}', cell3)
                        find3 = re.findall(r'\d{1,}', cell4)


                        for f in find[0]:
                            invert = int(f)
                            #print(invert)

                        for g in find[1:]:
                            invert2 = int(g)
                            #print(invert2)

                        for f in find2[0]:
                            invert3 = int(f)
                            #print(invert)

                        for g in find2[1:]:
                            invert4 = int(g)
                            #print(invert2)

                        for f in find3[0]:
                            invert5 = int(f)
                            #print(invert)

                        for g in find3[1:]:
                            invert6 = int(g)
                            #print(invert2)


                        asr = input('Press 1 to add quantity\nPress 2 to subtract quantity\nPress 3 to replace quantity\nPress 4 to quit\n--> ')


                        if asr == '1':

                            try:
                                Quan = int (input('Please enter the Quantity amount\n--> '))

                            except:
                                pass
                            try:
                                Add_Quan = Quantity_Num + Quan
                            except:
                                pass
                            #print(Add_Quan)
                            try:
                                if Quan > 0:
                                    now2 = datetime.datetime.now()
                                    date2 = now2.strftime("%Y-%m-%d %H:%M:%S")
                                    L_Action = 'Add +'
                                    ts.update_cell(invert2, invert, Add_Quan)
                                    ts.update_cell(invert4, invert3, date2)
                                    ts.update_cell(invert6, invert5, L_Action)
                                    print('Updated Quantity is' + Add_Quan)
                                    print('--------------------------------------------------')
                                    log = False
                            except:
                                pass







                        if asr == '2':

                            try:
                                Quan = int (input('Please enter the Quantity amount\n--> '))

                            except:
                                pass
                            try:
                                Add_Quan = Quantity_Num - Quan
                            except:
                                pass
                            #print(Add_Quan)
                            try:
                                if Quan > 0:
                                    now2 = datetime.datetime.now()
                                    date2 = now2.strftime("%Y-%m-%d %H:%M:%S")
                                    L_Action = 'Sub -'
                                    ts.update_cell(invert2, invert, Add_Quan)
                                    ts.update_cell(invert4, invert3, date2)
                                    ts.update_cell(invert6, invert5, L_Action)
                                    print('Updated Quantity is' + Add_Quan)
                                    print('--------------------------------------------------')
                                    log = False
                            except:
                                pass


                        if asr == '3':

                            try:
                                Quan = int (input('Please enter the Quantity amount\n--> '))

                            except:
                                pass
                            try:
                                Add_Quan = Quantity_Num + Quan
                            except:
                                pass
                            #print(Add_Quan)
                            try:
                                if Quan > 0:
                                    now2 = datetime.datetime.now()
                                    date2 = now2.strftime("%Y-%m-%d %H:%M:%S")
                                    L_Action = 'Replace R'
                                    ts.update_cell(invert2, invert, Quan)
                                    ts.update_cell(invert4, invert3, date2)
                                    ts.update_cell(invert6, invert5, L_Action)
                                    print('Updated Quantity is' + Add_Quan)
                                    print('--------------------------------------------------')
                                    log = False
                            except:
                                pass






                        if asr == '4':
                            log = False

                print('Sorry but this Product number does not exist')
                log = False





        if Menu_2 == '1':
            Part_Entry()

        elif Menu_2 == '2':
            Sub_2()

        elif Menu_2 == '3':
            Sub_3()

    print('--------------------------------------------------')

    ext = input('Press y to continue or x to exit\n-->')

    if ext == 'y':
        Zero = True
        Main()
    elif ext == 'x':
        Zero = False
    else:
        Zero = False


# any input is string
#number = input("Please guess what number I'm thinking of. HINT: it's between 1 and 30: ")
#try:                      # if possible, try to convert the input into integer
 #   number = int(number)
#except:                   # if the input couldn't be converted into integer, then do nothing
 #   pass
#print(type(number))       # see the input type after processing
