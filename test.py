import os
import glob
import pandas as pd
import pyxlsb
import openpyxl

all_files = {}
pl_folder = r"C:\Users\mikolaj.chrzan\Unilever\Poland Price Control - Team Documents"

for folder in os.listdir(pl_folder):
    i = 1
    try:
        if folder != "BALTICS":
            for file in os.listdir(r"{}\{}".format(pl_folder, folder)):
                if file == "99. Archive":
                    for x in os.listdir(r"{}\{}\{}".format(pl_folder, folder, file)):
                        if x == "2018":
                            new_list = []
                            all_data = pd.DataFrame()

                            for j in os.listdir(r"{}\{}\{}\{}".format(pl_folder, folder, file, x)):
                                if j == "2017":
                                    continue
                                else:
                                    if os.path.isdir(r"{}\{}\{}\{}\{}".format(pl_folder, folder, file, x, j)):
                                        for k in os.listdir(r"{}\{}\{}\{}\{}".format(pl_folder, folder, file, x, j)):
                                            try:
                                                df = pd.read_excel(r"{}\{}\{}\{}\{}\{}".format(pl_folder, folder,
                                                                                               file, x, j, k))
                                                all_data = all_data.append(df, ignore_index=True)
                                            except:
                                                try:
                                                    df = pd.read_excel(
                                                        r"{}\{}\{}\{}\{}\{}".format(pl_folder,
                                                            folder, file, x, j, k), engine='pyxlsb')
                                                    all_data = all_data.append(df, ignore_index=True)
                                                except Exception as e:
                                                    print(e)
                                    else:
                                        try:
                                            df = pd.read_excel(r"{}\{}\{}\{}\{}".format(pl_folder, folder, file, x, j))
                                            all_data = all_data.append(df, ignore_index=True)
                                        except:
                                            try:
                                                df = pd.read_excel(
                                                    r"{}\{}\{}\{}\{}".format(pl_folder,
                                                        folder, file, x, j), engine='pyxlsb')
                                                all_data = all_data.append(df, ignore_index=True)
                                            except Exception as e:
                                                print(e)

                            all_data.to_csv(r"C:\Temp\output\{}.csv".format(folder), sep=";", header=True)
        else:
            for bal_folder in os.listdir(r"{}\{}".format(pl_folder, folder)):
                for file in os.listdir(r"{}\{}\{}".format(pl_folder, folder, bal_folder)):
                    if file == "99. Archive":
                        for x in os.listdir(r"{}\{}\{}\{}".format(pl_folder, folder, bal_folder, file)):
                            if x == "2018":
                                new_list = []
                                all_data = pd.DataFrame()

                                for j in os.listdir(r"{}\{}\{}\{}\{}".format(pl_folder, folder, bal_folder, file, x)):
                                    if j == "2017":
                                        continue
                                    else:
                                        if os.path.isdir(r"{}\{}\{}\{}\{}\{}".format(pl_folder, folder, bal_folder,
                                                                                     file, x, j)):
                                            for k in os.listdir(r"{}\{}\{}\{}\{}\{}".format(pl_folder, folder,
                                                                                            bal_folder, file, x, j)):
                                                try:
                                                    df = pd.read_excel(r"{}\{}\{}\{}\{}\{}\{}".format(pl_folder, folder,
                                                                                                   bal_folder, file, x, j, k))
                                                    all_data = all_data.append(df, ignore_index=True)
                                                except:
                                                    try:
                                                        df = pd.read_excel(
                                                            r"{}\{}\{}\{}\{}\{}\{}".format(pl_folder,
                                                                                        folder, bal_folder, file, x, j, k),
                                                            engine='pyxlsb')
                                                        all_data = all_data.append(df, ignore_index=True)
                                                    except Exception as e:
                                                        print(e)
                                        else:
                                            try:
                                                df = pd.read_excel(r"{}\{}\{}\{}\{}\{}".format(pl_folder, folder,
                                                                                            bal_folder, file, x, j))
                                                all_data = all_data.append(df, ignore_index=True)
                                            except:
                                                try:
                                                    df = pd.read_excel(
                                                        r"{}\{}\{}\{}\{}\{}".format(pl_folder,
                                                                                 folder, bal_folder, file, x, j), engine='pyxlsb')
                                                    all_data = all_data.append(df, ignore_index=True)
                                                except Exception as e:
                                                    print(e)

                                all_data.to_csv(r"C:\Temp\output\{}.csv".format(bal_folder), sep=";", header=True)
    except Exception as e:
        print(e)
        continue

    all_files[folder] = new_list

