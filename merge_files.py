from dbfread import DBF
import xlwt
import psutil


for proc in psutil.process_iter():
    name = proc.name()
    if name == "PbxCollect.exe":
        print('__________________________________________________________________')
        print('ОТКЛЮЧИТЕ ВСЕ ПРОЦЕССЫ PBX COLLECTOR (запустите файл kill_process_PBX)')
        print('__________________________________________________________________')
        print('для выхода введите любой символ...')
        input()
        break
    else:
               
        #инициализация переменных основных файлов с данными
        try:     
            table1=DBF('./1/Calls/CALLS.DBF','cp866')
            table2=DBF('./2/Calls/CALLS.DBF','cp866')
            table3=DBF('./3/Calls/CALLS.DBF','cp866')
            table4=DBF('./4/Calls/CALLS.DBF','cp866')
            table5=DBF('./5/Calls/CALLS.DBF','cp866')
            table6=DBF('./6/Calls/CALLS.DBF','cp866')
            table8=DBF('./8/Calls/CALLS.DBF','cp866')
            table9=DBF('./9/Calls/CALLS.DBF','cp866')
            table10=DBF('./10/Calls/CALLS.DBF','cp866')
            table11=DBF('./11/Calls/CALLS.DBF','cp866')
            table12=DBF('./12/Calls/CALLS.DBF','cp866')
            table13=DBF('./13/Calls/CALLS.DBF','cp866')
            table14=DBF('./14/Calls/CALLS.DBF','cp866')
            table15=DBF('./15/Calls/CALLS.DBF','cp866')
            table16=DBF('./16/Calls/CALLS.DBF','cp866')
            table17=DBF('./17/Calls/CALLS.DBF','cp866')
            table18=DBF('./18/Calls/CALLS.DBF','cp866')
            table19=DBF('./19/Calls/CALLS.DBF','cp866')
            table20=DBF('./20/Calls/CALLS.DBF','cp866')
            table21=DBF('./21/Calls/CALLS.DBF','cp866')
            table22=DBF('./22/Calls/CALLS.DBF','cp866')
            table23=DBF('./23/Calls/CALLS.DBF','cp866')
            table24=DBF('./24/Calls/CALLS.DBF','cp866')
            table25=DBF('./25/Calls/CALLS.DBF','cp866')
            table26=DBF('./26/Calls/CALLS.DBF','cp866')
            table27=DBF('./27/Calls/CALLS.DBF','cp866')


            
            list1=[table1,table2,table3,table4,table5,table6,table8,table9,table10,table11,table12,table13,table14,table15,table16,table17,table18,table19,table20,table21,table22,table23,table24,table25,table26,table27]
            data=[]
            row1=()
            
            for i in list1:
                for rows in i:
                    row1 =tuple(rows.keys())
                    temp=tuple(rows.values())
                    if temp not in data:
                        data.append(temp)
            
            data.insert(0,row1)
            # print(data)
            # print(len(data))
            

            # Сводим все полученные списки с данными в один файл с помощью библиотеки xlwt

            WB = xlwt.Workbook () # Create a work
            f = WB.add_sheet ('class1') # Create a worksheet
            
            for i in range(len(data)):
                for j in range(len(row1)):
                    f.write(i,j,data[i][j])
            
            WB.save('./merge_Calls.DBF')
        
        except:  
            print('__________________________________________________________________')
            print('ФАЙЛЫ ДЛЯ СЛИЯНИЯ НЕ НАЙДЕНЫ')
            print('__________________________________________________________________')
            input('для выхода введите любой символ...')
            break
