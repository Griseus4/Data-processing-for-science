import pandas as pd
import os
import glob
import numpy as np
from datetime import datetime as dt
import time
import sys

# Launch:
# (base) C:\Users>python "C:\Users\capku\Korovnikov\Кафедральная прога\prog2.py" "C:\Users\capku\Korovnikov\Кафедральная прога"

class files_processing:
    
    def __init__(self, path):
        self.path = path
        self.files = glob.glob(os.path.join(self.path, "*.dat"))
        self.df = [None for x in range(4)]
        # print(self.files)
    
    
    def common(self):
    
        """
        Возвращает датафрейм на основе введённых пользователем данных
        """;

        self.df[0] = pd.DataFrame([["Дата", "Название образца", "Температура измерений", "Общее время травления", "Электроды"]])
        
        txt = input("\n1. Измерения проводились сегодня? [y/n]:\n>>> ")
        while txt not in ["y", "n"]:
            txt = input("\n1. Измерения проводились сегодня?\nВведите латинскую \"y\", если да, и латинскую \"n\" иначе:\n>>> ")
        if txt == "n":
            txt = input("\n1.1 Введите дату измерений в формате \"ДД.ММ.ГГ\"):\n>>> ")
        else:
            txt = dt.strftime(dt.now(), "%d.%m.%y")
        self.df[0].loc[1, 0] = txt

        self.df[0].loc[1, 1] = input("\n2. Введите название образца (например, \"ME_a1_H left\"):\n>>> ")

        txt = input("\n3. Измерения проводились при комнатной температуре? [y/n]:\n>>> ")
        while txt not in ["y", "n"]:
            txt = input("\n3. Измерения проводились при комнатной температуре?\nВведите латинскую \"y\", если да, и латинскую \"n\" иначе:\n>>> ")
        self.df[0].loc[1, 2] = "Комнатная" if (txt == "y") else "2.4 К"

        txt = input("\n4. Введите общее время травления в секундах(например, \"360\"):\n>>> ")
        while not txt.isdigit():
            txt = input("\n4. Введите общее время травления в секундах:\n>>> ")
        self.df[0].loc[1, 3] = txt

        txt = input("\n5. Используются только электроды 1-8? [y/n]:\n>>> ")
        while txt not in ["y", "n"]:
            txt = input("\n5. Используются только электроды 1-8?\nВведите латинскую \"y\", если да, и латинскую \"n\" иначе:\n>>> ")
        if txt == "n":
            txt = input("\n5.1 Введите номера используемых электродов, разделённые запятой (например, \"1,2,5,6,7,8,9,10\"):\n>>> ")
        else:
            txt = "1,2,3,4,5,6,7,8"
        self.df[0].loc[1, 4] = txt
        
        
    
    def spreadsheets(self):
        data_cols = ['U'] + [f"I{str(i)}" for i in range(1, 9)]
        # df_cols = list("12345678")
        self.df[1] = pd.DataFrame(columns=[i for i in range(1, 9)])
        self.df[2] = pd.DataFrame(columns=[i for i in range(1, 9)])
        
        for num in range(8):
            with open(self.files[num]) as f:
                i = 0
                for x in f:
                    i += 1
                    if i == 267:
                        elec_num = int(x[20:-1])
                # print(f"# {elec_num = }", end="\n\n")

                data = pd.read_csv(self.files[num], sep="\t", header=None, skiprows=268, encoding="latin-1")
                data.columns = data_cols
                
                # display(data.head())
                
                delta_U = (data['U'].max() - data['U'].min()) * 1000
                delta_I = (data.loc[:, 'I1':].max() - data.loc[:, 'I1':].min()) / delta_U * 10**9
                delta_I.reset_index(drop=True, inplace=True)
                
                # display(delta_I.head())
                
                # print(f"# delta_U\n{delta_U}", end="\n\n")
                # print(f"# delta_I\n{delta_I}", end="\n\n")

                self.df[1].loc[:, elec_num] = delta_I

                # print(f"# self.df[1][{elec_num}]\n{self.df[1].loc[:, elec_num]}", end='\n\n')

                self.df[2].loc[:, elec_num] = delta_U / delta_I *10**3

                # print(f"# self.df[2][{elec_num}]\n{self.df[2].loc[:, elec_num]}", end="\n\n")
        
        # print('# self.df[1]:'); display(self.df[1])
        # print('# self.df[2]:'); display(self.df[2])
        
        
    def concat_df(self):
        
        """for i in range(3):
            try:
                print(f"\tdf[{i}]:")
                display(self.df[i])
                # print(f"{self.df[i].columns = }")
                # print(f"{self.df[i].index = }\n-----------\n\n")
            except: pass"""
        
        df_for_IU = pd.DataFrame([[None], [int(i) for i in range(1, 9)]], columns=[i for i in range(1, 9)])
        df_left = pd.DataFrame([None] + ['I (нА при 1 мВ)'] + [i for i in range(1, 9)] + [None, 'R (кОм)'] + [i for i in range(1, 9)])
        
        df = pd.concat([df_for_IU, self.df[1], df_for_IU, self.df[2]], axis=0, ignore_index=True)
        df.reset_index(drop=True, inplace=True)
        df = pd.concat([df_left, df], axis=1)
        df = pd.concat([self.df[0], None, df])
        
        # print('\n\n\tdf_left:\n'); display(df_left)
        # print('\n\n\tdf:\n'); display(df)
        # print('\n\n\tdf_for_IU:\n'); display(df_for_IU)
        
        self.df = df
    
    
    def create_xlsx(self):
    
        """
        Создаёт xlsx-файл, названный по времени своего создания.
        df - датафрейм, на основе которого создаётся файл
        """;
        
        # display(self.df)
        # self.df.to_excel(dt.strftime(dt.now(), "%y%m%d_%H%M") + ".xlsx", header=False, index=False)
        self.df.to_excel(os.path.join(self.path, dt.strftime(dt.now(), "%y%m%d_%H%M") + ".xlsx"), header=False, index=False)       


def main():
    file_path = sys.argv[1]
    print("______________________________\n\tНачало работы программы\n\n")
    files = files_processing(file_path)
    files.common()
    files.spreadsheets()
    files.concat_df()
    files.create_xlsx()
    print("\tКонец работы программы.\n\tИщите готовый excel файл в папке с измерениями\n______________________________\n\n")

if __name__ == "__main__":
    main()