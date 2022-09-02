# -*- coding: utf-8 -*-

import tkinter as tk
import tkinter.filedialog
import tkinter.simpledialog
from tkinter import Menu

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from pandas import ExcelWriter

while True:
    def MIC_data_cleaning(path, num_plates, excel_name, units, png=False):
        count = 0
        while True:
            if count == 1:
                break
            else:
                try:
                    # 1.Opening the dataset and settings:
                    MIC_raw = pd.read_csv(path, encoding='unicode_escape')
                    pd.set_option('display.max_rows', 400)

                    # 2.Number of plates:
                    def number_of_plates_rows_drop(ncolumns, nrows=12):
                        MIC_drop_rows = MIC_raw.drop(np.arange(0, nrows * ncolumns))
                        return MIC_drop_rows

                    MIC = number_of_plates_rows_drop(
                        num_plates).copy()  # Change the value equal to the name of plates you did.

                    # 3. Cleaning the columns:
                    MIC = MIC.reset_index().drop('index', axis=1).copy()
                    MIC = MIC.rename(columns=MIC.iloc[1]).drop(1, axis=0).copy()
                    MIC = MIC.reset_index().drop('index', axis=1).copy()
                    MIC = MIC.loc[:, MIC.columns.notnull()].copy()

                    # 4.Regex:
                    pattern_1 = 'Group:'

                    Bacteria_split = MIC['Sample'].str.extract(r'(Group: )([a-zA-Z0-9_:. ]*)')

                    MIC['Bacteria'] = Bacteria_split[1]
                    MIC['Bacteria'] = MIC['Bacteria'].fillna(method='ffill')

                    Boolean_mask_cleaning = MIC['Sample'].str.contains(pattern_1, regex=True).fillna(False)
                    Boolean_mask_cleaning = MIC[Boolean_mask_cleaning]
                    MIC = MIC.drop(index=Boolean_mask_cleaning.index).copy()  # Dropping the rows with the name of Group: Bacteria

                    # 5.More boolean mask searching for NaN and 13 rows forward.
                    MIC = MIC.dropna().copy()
                    MIC = MIC.drop(index=MIC[MIC['Sample'] == 'Sample'].index).copy()

                    MIC['Sample'] = MIC['Sample'].str.extract(r'([a-zA-Z0-9/ ]*)(_[0-9]*)')[0].copy()
                    MIC = MIC.dropna().copy()

                    # 6. Clean Data:
                    MIC_1 = MIC.drop(['Well', 'Values'], axis=1).copy()

                    # 7. Single Antibiotic vs All bacteria:
                    MIC_2 = MIC_1.pivot_table(values='MeanValue', index=['Sample', 'Bacteria'], columns=units,
                                              aggfunc='first')
                    MIC_2.columns = MIC_2.columns.astype('float')
                    MIC_2 = MIC_2.sort_index(axis=1, ascending=False)

                    # 8. Single Bacteria vs All Antibiotics:
                    MIC_3 = MIC_1.pivot_table(values='MeanValue', index=['Bacteria', 'Sample'], columns=units,
                                              aggfunc='first')
                    MIC_3.columns = MIC_3.columns.astype('float')
                    MIC_3 = MIC_3.sort_index(axis=1, ascending=False)

                    # 9. Excel file_writer:
                    writer = ExcelWriter(excel_name)
                    MIC_raw.to_excel(writer, sheet_name="Raw Data")
                    MIC_1.to_excel(writer, sheet_name="Clean Data")
                    MIC_2.to_excel(writer, sheet_name="Single ABX vs All bacteria")
                    MIC_3.to_excel(writer, sheet_name="Single Bacteria vs All ABX")

                    # 10. Plots and Std.dev bars:
                    MIC_1[["MeanValue", "Std.Dev.", "CV%"]] = MIC_1[["MeanValue", "Std.Dev.", "CV%"]].astype('float')
                    bacteria = MIC_1['Bacteria'].unique()
                    antibiotic = MIC_1['Sample'].unique()

                    count = 0
                    MIC_final = []
                    for b in bacteria:
                        bac = MIC_1[MIC_1['Bacteria'] == b]
                        for p in antibiotic:
                            count += 1
                            abx = bac[bac['Sample'] == p]
                            dictionary = dict(zip(abx[units], abx['MeanValue']))
                            print(abx)

                            # MIC final results inside the loop for select the final values:
                            MIC_5 = abx[abx["MeanValue"].astype("float") <= 0.009].copy()
                            MIC_5[units] = MIC_5[units].astype("float")
                            MIC_5 = MIC_5[MIC_5[units] == min(MIC_5[units], default=0)]
                            MIC_final.append(MIC_5)

                            # Bar plots bacteria and antibiotics:
                            plt.bar(dictionary.keys(), height=dictionary.values(), width=0.4)
                            plt.errorbar(dictionary.keys(), dictionary.values(), yerr=abx["Std.Dev."], fmt='.k')
                            plt.title(b + ', ' + p)
                            plt.xlabel(units)
                            plt.ylabel('MeanValue')
                            plt.xticks(abx[units])
                            xlocs, xlabs = plt.xticks()
                            for i, v in enumerate(dictionary.values()):
                                plt.text(xlocs[i] - 0.25, v + 0.01, str(v))

                            #To Save the figures in png format uncomment this section:
                            if png == True:
                                plt.savefig(png + b + p + ".png")
                                plt.show()

                            # Dividing MIC_1 in an independent sheet, positioning the dataframes in the worksheet and plot the data.
                            position = 15 * (count - 1)
                            abx.to_excel(writer, sheet_name="bac_abx_plot", startrow=position)
                            workbook = writer.book
                            worksheet = writer.sheets['bac_abx_plot']
                            chart = workbook.add_chart({'type': 'column'})
                            chart.add_series({
                                'name': ['bac_abx_plot', 1 + position, 1],
                                'categories': ['bac_abx_plot', 1 + position, 2, 9 + position, 2],
                                'values': ['bac_abx_plot', 1 + position, 3, 9 + position, 3],
                                'data_labels': {'value': True},
                            })

                            chart.set_title({'name': b + ' ' + p})
                            chart.set_x_axis({'name': units})
                            chart.set_y_axis({'name': 'OD600', 'major_gridlines': {'visible': False}})

                            # Insert the chart into the worksheet.
                            worksheet.insert_chart(0 + position, 8, chart)
                            print(count)

                    # 11. MIC:
                    MIC_final_df = pd.concat(MIC_final)
                    MIC_final_summary = MIC_final_df.pivot_table(values=units, index='Bacteria', columns='Sample',
                                                                 aggfunc='first')
                    print("\n", MIC_final_summary, "\n" * 2, MIC_final_df)

                    # 12. Excel file_sheets:
                    MIC_final_summary.to_excel(writer, sheet_name="MIC")
                    MIC_final_df.to_excel(writer, sheet_name="MIC", startrow=9)
                    writer.save()


                except:
                    # Exit window when error arise:
                    print("Oops! Some input was incorrect. Please try again.")
                    window = tk.Tk()
                    window.title("Welcome to DANAMIC GUI !!!")
                    window.geometry('230x150+700+300')
                    window.configure(bg="white")
                    error_window = tk.StringVar(window)
                    msg = tk.Message(window, textvariable=error_window)
                    msg.config(bg='red', font=('times', 24, 'italic'))
                    error_window.set("Oops! Some input was incorrect. Please try again.")
                    msg.pack()

                    window.config(menu=menu)
                    window.mainloop()
                    break


            # Exit window when not errors arise:
            count = 1  # Stopping the while loop if not errors arise.
            print("Your data is now cleaned and analysed")

            window = tk.Tk()
            window.title("Welcome to DANAMIC GUI !!!")
            window.geometry('180x118+700+300')
            window.configure(bg="white")
            good_window = tk.StringVar(window)
            msg = tk.Message(window, textvariable=good_window)
            msg.config(bg='lightgreen', font=('times', 24, 'italic'))
            good_window.set("Your data is now cleaned and analysed")
            msg.pack()
            window.config(menu=menu)
            window.mainloop()


    try:
        # Tkinter GUI:
        window = tk.Tk()
        window.title("Welcome to DANAMIC GUI !!!")
        window.geometry('850x350+400+300')
        window.configure(bg="white")
        window.resizable(width=True, height=False)
        

        # MENUS AND ITEMS:
        def openfile():
            global filename
            filename = tk.filedialog.askopenfilename(initialdir="root.filename",
                                                     title="DANAMIC Choose your file for you!!!",
                                                     filetypes=(
                                                               ("Microsoft Excel Comma Separated Values File", "*.csv"),
                                                               ("Microsoft Excel Worksheet", "*.xlsx"),
                                                               ("all files", "*.*")
                                                              )
                                                     )

            print(filename)

            label_openfile = tk.Label(window, text="Open file path: ", font=("Calibri Bold", 10), bg="light sky blue",
                                      fg="black")
            label_openfile.grid(column=0, row=0, )
            label_openfile_path = tk.Label(window, text=filename, font=("Calibri Bold", 10), bg="pale green", fg="black")
            label_openfile_path.grid(column=1, row=0)

        def numberplates():
            global plates_number
            plates_number = tk.simpledialog.askinteger(prompt="Enter a number:", title="How many plates?", minvalue=0,
                                                       maxvalue=100, )
            print(plates_number)

            label_number_name = tk.Label(window, text="Number of plates: ", font=("Calibri Bold", 10), bg="light sky blue",
                                         fg="black")
            label_number_name.grid(column=0, row=1, )
            label_number = tk.Label(window, text=plates_number, font=("Calibri Bold", 10), bg="pale green", fg="black")
            label_number.grid(column=1, row=1)

        def savefiles():
            global savefile
            savefile = tk.filedialog.asksaveasfilename(initialdir="root.filename",
                                                       title="DANAMIC Save your file for you!!!",
                                                       filetypes=(
                                                                  ("Microsoft Excel Worksheet", "*.xlsx"),
                                                                  ("Microsoft Excel Comma Separated Values File", "*.csv"),
                                                                  ("all files", "*.*")
                                                                 )
                                                       )
            print(savefile)

            label_savefile = tk.Label(window, text="Saved file path: ", font=("Calibri Bold", 10), bg="light sky blue",
                                      fg="black")
            label_savefile.grid(column=0, row=2, )
            label_savefile_path = tk.Label(window, text=savefile, font=("Calibri Bold", 10), bg="pale green", fg="black")
            label_savefile_path.grid(column=1, row=2)

        def molar_unit():
            global unit
            unit = "Concentration µM"
            print(unit)
            label_molar = tk.Label(window, text="    Concentration µM    ", font=("Calibri Bold", 10), bg="light sky blue",
                                   fg="black")
            label_molar.grid(column=0, row=3)

        def milligram_unit():
            global unit
            unit = "Concentration (mg/L)"
            print(unit)
            label_milligram = tk.Label(window, text="Concentration (mg/L)", font=("Calibri Bold", 10), bg="light sky blue",
                                       fg="black")
            label_milligram.grid(column=0, row=3 )

        menu = Menu(window)
        new_item1 = Menu(menu, tearoff=0)
        new_item2 = Menu(menu, tearoff=0)
        new_item3 = Menu(menu, tearoff=0)
        new_item4 = Menu(menu, tearoff=0)
        new_item5 = Menu(menu, tearoff=0)
        new_item6 = Menu(menu, tearoff=0)

        new_item1.add_command(label='Open', command=openfile)
        new_item2.add_command(label='Number of plates', command=numberplates)
        new_item3.add_command(label='Save', command=savefiles)
        run_button = tk.Button(window, text="Run", command=window.destroy, height=2, width=10, padx=5, pady=5, bd=4).place(
            x=380, y=200)
        new_item4.add_command(command=run_button)
        new_item5.add_command(label="Concentration µM",  command=molar_unit)
        new_item5.add_command(label="Concentration mg/L", command=milligram_unit)

        # Button for closing and exit the program.
        def callback():
            # The code, changes the variable value from False to True (or reverse) every time you press the button.
            global buttonClicked
            buttonClicked = not buttonClicked
            if buttonClicked == True:
                exit()

        buttonClicked = False    # Before first click
        exit_button = tk.Button(window, text="Exit", command=callback, height=2, width=10, padx=5, pady=5, bd=4).place(
            x=380, y=270)

        menu.add_cascade(label='File', menu=new_item1)
        menu.add_cascade(label='Plates', menu=new_item2)
        menu.add_cascade(label='Save', menu=new_item3)
        menu.add_cascade(label='Units', menu=new_item5)

        window.protocol('WM_DELETE_WINDOW', callback) # Exit the program using the close button of the menu (X).
        window.config(menu=menu)
        window.mainloop()

        # FINAL: Parameters to modify each new MIC:
        MIC_data_cleaning(path=filename, num_plates=plates_number, excel_name=savefile, units=unit, png=savefile)


    except:
        # When exit_button is clicked, close the exception window exiting the program.
        if buttonClicked == True:
            exit()

        # Exit window when error arise:
        print("Oops! Some input was incorrect. Please try again.")
        window = tk.Tk()
        window.title("Welcome to DANAMIC GUI !!!")
        window.geometry('230x150+700+300')
        window.configure(bg="white")
        error_window = tk.StringVar(window)
        msg = tk.Message(window, textvariable=error_window)
        msg.config(bg='red', font=('times', 24, 'italic'))
        error_window.set("Oops! Some input was incorrect. Please try again.")
        msg.pack()
        window.config(menu=menu)
        window.mainloop()