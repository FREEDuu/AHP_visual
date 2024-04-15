
import pandas as pd
import os
import customtkinter
from tkinter import filedialog
import numpy as np
from matplotlib import pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors


desktop = os.path.normpath(os.path.expanduser("~/Desktop"))

def aggiungi_colonna_nomi(df):
    global param_button_weight
    for elementi in range(param_button_weight):
        text_griglia3 = customtkinter.CTkTextbox(scrollable_frame1, width=500//param_button_weight, height=500//param_button_weight)
        if elementi == 0:
            text_griglia3.insert("0.0", 'confronto')  # insert at line 0 character 0
        else:
            text_griglia3.insert("0.0", df.columns[elementi])  # insert at line 0 character 0
        text_griglia3.grid( row = elementi, column = 0, sticky="ew")
        text_griglia3.configure(state="disabled") 

def aggiungi_riga_nomi(df):
    global param_button_weight
    for elementi in range(param_button_weight):
        if elementi != 0:
            text_griglia3 = customtkinter.CTkTextbox(scrollable_frame1, width=500//param_button_weight, height=500//param_button_weight)
            text_griglia3.insert("0.0", df.columns[elementi])  # insert at line 0 character 0
            text_griglia3.grid( row = 0, column = elementi, sticky="ew")
            text_griglia3.configure(state="disabled") 

def ahp_attributes(ahp_df):
    # Creating an array of sum of values in each column
    sum_array = np.array(ahp_df.sum(numeric_only=True))
    # Creating a normalized pairwise comparison matrix.
    # By dividing each column cell value with the sum of the respective column.
    cell_by_sum = ahp_df.div(sum_array,axis=1)
    # Creating Priority index by taking avg of each row
    priority_df = pd.DataFrame(cell_by_sum.mean(axis=1),
                               index=ahp_df.index,columns=['priority index'])
    priority_df = priority_df.transpose()
    return priority_df

def create_matrice_confronto(pesi_nomi):
    df = {}
    df1 = {}
    nomi = []
    
    for el in range(len(pesi_nomi)):
        if el % param_button_weight == 0:
            nomi.append(pesi_nomi[el][1])
    df1['contro'] = nomi

    to_app = []
    for el in range(len(pesi_nomi)):
        to_app.append(float(pesi_nomi[el][0]))
        if el % (param_button_weight-1) == 0 and el != 0:
            initi = to_app.pop(-1)
            df[pesi_nomi[el-1][1]] = to_app
            df1[pesi_nomi[el-1][1]] = to_app
            to_app = []
            to_app.append(initi)
    df[pesi_nomi[-1][1]] = to_app
    df1[pesi_nomi[-1][1]] = to_app
    df1 = pd.DataFrame(df1)
    return ahp_attributes(pd.DataFrame(df))
def da_st_a_int(st):
    if st == 'basso':
        return 3
    if st == 'alto':
        return 5
    if st == 'medio':
        return 8
    return 0

def browseFiles():
    filename = filedialog.askopenfilename(initialdir = desktop,
                                            title = "Select a File",
                                            filetypes = [("EXCEL", "*.xlsx"),
                                                ("CSV", "*.csv"),
                                                ("All files", "*")])
    textbox = customtkinter.CTkTextbox(app, height=32)
    textbox.insert("0.0", filename)  # insert at line 0 character 0
    textbox.grid(row=2, column=0, sticky="ew")
    textbox.configure(state="disabled")  # configure textbox to be read-only
    button2.configure(state = 'normal')

    global path
    path = filename
        
    return filename, path


def load_projects(path):

    """ This function load the project and return a list of projects objects"""

    if 'csv' in path :
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    textbox3 = customtkinter.CTkTextbox(app)
    global param_button_weight
    global scrollable_frame1
    global scrollable_frame2
    global griglia
    global indexs

    indexs = df.index
    param_button_weight = len(df.columns)
    scrollable_frame1.grid(row =6, column = 0, pady = 10)
    scrollable_frame2.grid(row = 8, column = 0, pady = 10)

    text_griglia = customtkinter.CTkTextbox(app, height=32, width=550)
    text_griglia.insert("0.0", 'Matrice Confronti')  # insert at line 0 character 0
    text_griglia.grid(row= 4, sticky="ew")
    text_griglia.configure(state="disabled") 
    
    for colonna in range(param_button_weight):
        if colonna == 0:
            aggiungi_colonna_nomi(df)
        else:

            for riga in range(param_button_weight):
                if riga == 0:
                    aggiungi_riga_nomi(df)
                else:
                    confronto = customtkinter.CTkEntry(scrollable_frame1, width=500//param_button_weight, height=500//param_button_weight)
                    confronto.grid(row = riga, column=colonna, padx=1, pady=1, sticky="ew")
                    griglia.append([confronto, df.columns[colonna]])

    text_griglia2 = customtkinter.CTkTextbox(app, height=32, width=550)
    text_griglia2.insert("0.0", 'Relativi Pesi')  # insert at line 0 character 0
    text_griglia2.grid(row= 7, sticky="ew")
    text_griglia2.configure(state="disabled") 
    for i in range(param_button_weight):
        if i != 0:
            textbox3 = customtkinter.CTkTextbox(scrollable_frame2, height=32, width=100)
            textbox3.insert("0.0", df.columns[i])  # insert at line 0 character 0
            textbox3.grid(row=i*3, column=0, sticky="ew")
            textbox3.configure(state="disabled")  # configure textbox to be read-only
            
            for el in range(len(df.index)):
                slider = customtkinter.CTkEntry(scrollable_frame2, placeholder_text= 'importanza da 1 a 10',width = 480/df.shape[0])
                slider.grid(row=(i*3)+1, column=el+1, padx =0, pady=0, sticky="ew")
                sliders.append([slider, df.columns[i]])
                textbox4 = customtkinter.CTkTextbox(scrollable_frame2, height=32, width = 480/df.shape[0])
                textbox4.insert("0.0", df.iloc[:,0][el])  # insert at line 0 character 0
                textbox4.grid(row=i*3, column= el+1, sticky="ew")
                textbox4.configure(state="disabled")  # configure textbox to be read-only
                

    button3 = customtkinter.CTkButton(app, text= 'Calcola il Miglior Progetto' , command= calcola)
    button3.grid(row=9, column=0, padx=20, pady=20, sticky="ew")

def calcola():
    valori = []
    progetti = {}
    for entry in griglia:
        valori.append([entry[0].get(), entry[1]])

    weights = create_matrice_confronto(valori)
    weights1 = weights.iloc[0]
    ww = []
    for el in weights1:
        ww.append(el)
    peso = []
    to_dict = {}
    for pesi in sliders:
        peso.append([pesi[0].get(), pesi[1]])
        if pesi[1] not in to_dict.keys():
            to_dict[pesi[1]] = [pesi[0].get()]
        else:
            to_dict[pesi[1]].append(pesi[0].get())

    dictionaries = []
    for dizionari in to_dict.keys():
        to_dict1 = {} 
        to_app = []
        for el in range(len(to_dict[dizionari])):
            for el1 in to_dict[dizionari]:
                to_app.append(int(el1)/int(to_dict[dizionari][el]))
            to_dict1[el] = to_app
            to_app = []
        dictionaries.append(to_dict1)
        to_dict1 = {}
    vettori = []    
    for dic in dictionaries:
        dic = pd.DataFrame(dic)
        normalized_df=(dic)/dic.sum()
        vettori.append(ahp_attributes(normalized_df))
    
    if 'csv' in path :
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    for colonne in range(len(df.columns)):
        if colonne != 0:
            df[df.columns[colonne]] = df[df.columns[colonne]] * ww[colonne-1]
    nomi_progetti = df['NOME PROGETTO']
    final_df = {}
    
    for el in range(len(df.columns)):
        if df.columns[el] != 'NOME PROGETTO':
            vettori[el-1] = vettori[el-1] * ww[el-1]
    for vet in range(len(vettori)):
        final_df[df.columns[vet+1]] = vettori[vet].squeeze(axis = 0)
    final_df = pd.DataFrame(final_df)

    dataframe_plot = {}
    for colonne in final_df.columns:
        dataframe_plot[colonne] = final_df[colonne]

    final_df['Total'] = final_df.sum(axis=1)
    final_df['progetti'] = nomi_progetti

    fig, ax = plt.subplots()

    progetti = nomi_progetti

    bottom = np.zeros(len(nomi_progetti))

    for boolean, weight_count in dataframe_plot.items():
        p = ax.bar(progetti, weight_count,width = 0.5 ,label=boolean, bottom=bottom)
        bottom += weight_count
    ax.legend(loc="upper right")
    pdf = SimpleDocTemplate("Risultati Finali", pagesize=letter)
    pdf2 = SimpleDocTemplate("Risultati Finali", pagesize=letter)

    table_data = []
    to_app = []
    df = df.round(3)
    weights = weights.round(3)

    table_data2 = []

    final_df = final_df.round(3)
    for el in df.columns:
        if el != 'NOME PROGETTO':
            to_app.append(el)
    table_data2.append(list(to_app))
    to_app.append('TOTALE')
    to_app.append('PROGETTI')
    table_data.append(list(to_app))
    for i, row in final_df.iterrows():
        table_data.append(list(row))
    for i, row in weights.iterrows():
        table_data2.append(list(row))

    table = Table(table_data)
    table2 = Table(table_data2)

    table_style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 14),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 1), (-1, -1), 12),
    ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
])

    table.setStyle(table_style)
    table2.setStyle(table_style)
    pdf_table = []
    pdf_table.append(table)
    pdf_table.append(table2)

    textbox3 = customtkinter.CTkTextbox(app, height=32)
    textbox3.insert("0.0",'I risultati sono stati calcolati')  # insert at line 0 character 0
    textbox3.grid(row=10, column=0, sticky="ew")
    textbox3.configure(state="disabled")  # configure textbox to be read-only

    pdf.build(pdf_table)
    plt.title('Risultati Finali')
    plt.show()
    plt.savefig("Pesi.pdf", format="pdf", bbox_inches="tight")

if __name__ == "__main__":
        
    path = '/'
    param_button_weight = 0
    indexs = 0
    griglia = []
    pesi = []
    sliders = []

    app = customtkinter.CTk()
    app.title("Match Progetti/Strumenti")
    app.geometry("700x1000")

    textbox1 = customtkinter.CTkTextbox(app, height=32)
    textbox1.insert("0.0", 'Carica il File Dove ci sono i Progetti')  # insert at line 0 character 0
    textbox1.grid(row=0, column=0, sticky="ew")
    textbox1.configure(state="disabled")  # configure textbox to be read-only

    button = customtkinter.CTkButton(app, text= 'Path file da scegliere' , command=browseFiles)
    button.grid(row=1, column=0, padx=20, pady=20, sticky="ew")
    scrollable_frame1 = customtkinter.CTkScrollableFrame(master=app, height=250, width=600)
    scrollable_frame2 = customtkinter.CTkScrollableFrame(master=app, height=250, width=600)

    button2 = customtkinter.CTkButton(app, text= 'Genera Le Matrici' , command= lambda : load_projects(path), state = 'disabled')
    button2.grid(row=4, column=0, padx=20, pady=20, sticky="ew")

    app.grid_columnconfigure(0, weight=1)
    app.mainloop()