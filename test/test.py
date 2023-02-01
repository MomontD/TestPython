import sqlite3
db = sqlite3.connect("testbase.db")
cursor = db.cursor()

import pymongo
from pymongo import MongoClient
cluster = MongoClient("mongodb+srv://momontd:momontd@clusterpy.0c9pxtr.mongodb.net/?retryWrites=true&w=majority")
dbmongo = cluster["test"]           # db = cluster.ovdp_data
collection = dbmongo["testcoll"]    #collection = db.ovdp_collections

class ovdp_parameters() :
    """Опис параметрів ОВДП"""
    def __init__(self,_id,start_date,end_date,cost,rate,repayments):
        self._id        = _id
        self.start_date = start_date
        self.end_date   = end_date
        self.cost       = cost
        self.rate       = rate
        self.repayments = repayments

import openpyxl
#зробити через try/expect і перенести нижче код в категорію Операції з витратами
book = openpyxl.open("statements_privat.xlsx", read_only = True)
sheet= book.active
active_cells= sheet.iter_rows(min_row = 3, max_row = sheet.max_row-1, min_col = 1, max_col = 6)

def insert_data_in_SQL (arg1, arg2, arg3, arg4, arg5, arg6, arg7):
    cursor.execute( f"INSERT INTO {arg1} VALUES(?, ?, ?, ?, ?, ?)",(arg2, arg3, arg4, arg5, arg6, arg7))

menu=0

while menu != 6 :
    menu = int(input("""What do you wont to do ? \n
                     1: Operation with expenses   \n    
                     2: Operation with investmens \n
                     3: Analitics \n
                     5: Base zvit \n
                     6: Exit      \n
                     Enter your choise : """))

    ######### Base menu ########
    if menu == 1 :
        menu1 = 0
        while menu1 != 3 :
            menu1 = int(input(""" 
                                   1: Show expenses from DataBase \n    
                                   2: Add expenses to DataBase    \n
                                   3: Exit                        \n
                                   Enter your choise : """))
            if menu1 == 1 :
                for data in cursor.execute("SELECT rowid,* FROM test_data") :
                    print(data)

            if menu == 2 :
                #------- read actual date and tume from SQL test_data)
                cursor.execute("SELECT MAX(date) FROM test_data")
                actual_date= cursor.fetchone()[0] # повертає кортеж (date,empty)- забираємо з нього 1 елемент
                cursor.execute(f"SELECT MAX(time) FROM test_data WHERE date = '{actual_date}'")
                actual_time= cursor.fetchone()[0] # повертає кортеж (time,empty)- забираємо з нього 1 елемент
                print("Актуальна дата в базі даних : ", actual_date, actual_time)

                #------- add new expenses to SQL table test_data
                for date, time, category, card, opertion, sum_operation in active_cells :
                    if date.value == actual_date :
                        if time.value > actual_time :
                            insert_data_in_SQL ("test_data", date.value, time.value, category.value, card.value, opertion.value, sum_operation.value)
                            print('Доповнені витрати за дату:', date.value , 'час :' ,time.value)
                        else :
                            print ("Такий запис вже є в базі даних",date.value, time.value)
                    elif date.value > actual_date :
                        insert_data_in_SQL ("test_data", date.value, time.value, category.value, card.value, opertion.value, sum_operation.value)
                        print("Додано дані за дату",date.value)
                    else : print ("Дані з файлу вже є в базі даних :",date.value, time.value)
                db.commit()

    ########################## MENU 2 : Operation with investmens ############################################
    if menu == 2 :
        menu2 = 0
        while menu2 != 5 :
            menu2 = int(input(""" 
                                   1: Show investments          \n    
                                   2: Add  investmens (OVDP)    \n
                                   3: Add  investmens (Deposit) \n
                                   4: Add  investmens (Debt)    \n
                                   5: Exit                      \n
                                   Enter your choise : """))

            if menu2 == 1 : # Show investmens
                print("Investments in OVDP")
                incomming_data = collection.find()
                for obj in incomming_data :
                    for el in obj :
                        print(el,": ", obj[el])
                    print("\n")

                print("Investments in Deposits \n")
                for deposit_data in cursor.execute("SELECT rowid,* FROM deposit_data") :
                    print(deposit_data,"\n")

                print("Investments in Debt \n")
                for debt_data in cursor.execute("SELECT rowid,* FROM debt_data") :
                    print(debt_data,"\n")

            if menu2 == 2 : #OVDP
                _id        = input ("Введіть номер ОВДП : ")
                start_date = input ("Введіть дату пчатку вкладу : ")
                end_date   = input ("Введіть дату завершення вкладу : ")
                cost       = input ("Введіть суму вкладу : ")
                rate       = input ("Вкажіть відсоткову стаку (%):")

                counter = 1
                repayments = []
                answer = "y"

                while answer != "n" :
                    repayment_date = input (f"Введіть дату {counter} погашення : ")
                    repayment_sum  = input (f"Введіть суму {counter} погашення ")

                    repayments.append({"date" : repayment_date , "sum" : float(repayment_sum)})

                    answer = input ("Продовжити (y/n)? : ")
                    counter+=1

                new_ovdp = ovdp_parameters(_id,start_date,end_date,float(cost),rate,repayments)
                collection.insert_one(new_ovdp.__dict__)

            if menu2 == 3 : #Deposit
                deposit_name   = input ("Введіть назву депозиту : ")
                deposit_sum    = input ("Введіть суму депозиту : ")
                deposit_rate   = input ("Введіть відсоткову ставку % депозиту : ")
                deposit_period = input ("Введіть термін депозиту в місяцях : ")
                start_date     = input ("Введіть дату початку вкладу : ")
                end_date       = input ("Введіть дату завершення вкладу : ")

                insert_data_in_SQL ("deposit_data", deposit_name, deposit_sum, deposit_rate, deposit_period, start_date, end_date)
                db.commit()
                print ("Data entered successfully! \n")

            if menu2 == 4 :  #Debt
                debt_name   = input ("Введіть назву позики : ")
                debt_sum    = input ("Введіть суму позики  : ")
                debt_rate   = input ("Введіть відсоткову ставку % позики : ")
                debt_period = input ("Введіть термін позики в місяцях    : ")
                start_date  = input ("Введіть дату початку    : ")
                end_date    = input ("Введіть дату завершення : ")

                insert_data_in_SQL ("debt_data", debt_name, debt_sum, debt_rate, debt_period, start_date, end_date)
                db.commit()
                print ("Data entered successfully! \n")

book.close()
db.close()