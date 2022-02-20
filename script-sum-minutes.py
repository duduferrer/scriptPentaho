import pandas as pd
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
from tkinter import *
import datetime


def execute():
    showinfo(
        title='Escolha um arquivo',
        message='Selecione a tabela de horas do PENTAHO'
    )
    filename = select_file()
    df_pentaho = open_file(filename)
    df_pentaho = group_by_operator(df_pentaho)
    showinfo(
        title='Escolha um arquivo',
        message='Selecione a tabela base'
    )
    filename = select_file()
    df_from_sheet = open_file(filename)
    print(df_from_sheet)
    result = join(df_pentaho, df_from_sheet)
    print(result)
    df = add_sum(result)
    # filtered_df = filter_active_operators(df)
    filepath = create_output(df)
    showinfo(
         title='Arquivo Criado',
        message=filepath
    )


def join(df_pentaho, df_from_sheet):
    #merge two DB, inner_right, df_from_sheet has priority
    df_from_sheet = df_from_sheet[['-', 'ATCO', 'IND. OP.', 'HE (h)']]
    result = pd.merge(df_pentaho,
                          df_from_sheet,
                          on=['IND. OP.'],
                          how='right')
    result = result[['-', 'ATCO', 'IND. OP.', 'HE (h)', 'HL (h)']]
    result['HL (h)'] = result[['HL (h)']].fillna("0")
    result['HE (h)'] = result[['HE (h)']].fillna("0")
    return result

#sum and add to the DF the total amount of hours of HE/HL and IDBR
def add_sum(df):
    #count the number of rows
    count_row = df.shape[0]
    #row containing the sum of all operators hours
    totalammount_row = count_row-2

    #sum the HE column
    df_he = df[['HE (h)']].copy().astype(float)
    he_sum = df_he.sum()
    he_sum = he_sum[0]

    # sum the HL column
    df_hl = df[['HL (h)']].copy().astype(float)
    #df_hl = df_hl.applymap(lambda entry: make_delta(entry))
    hl_sum = df_hl.sum()
    hl_sum = hl_sum[0]
    print(hl_sum)

    #calculate IDBR
    idbr = hl_sum/he_sum*100
    idbr = "%.2f%%" % idbr
    print(idbr)

    #put HL and HE sum in the right format for the table
    #hl_sum = time_formatter(hl_sum)
    #he_sum = time_formatter(he_sum)

    #put the data calculated into the original DF
    df.at[totalammount_row+1, 'HE (h)'] = idbr
    df.at[totalammount_row+1, 'HL (h)'] = ""
    df.at[totalammount_row, 'HL (h)'] = hl_sum
    df.at[totalammount_row, 'HE (h)'] = he_sum
    return df

#FORMAT TIME FROM DAYS/HOURS/MINUTES TO H:MM
def time_formatter(duration):
    totsec = duration.total_seconds()
    h = totsec / 3600
    #m = (totsec % 3600) // 60
    #return "%d:%02d" % (h, m)
    return "%.2f" % (h)

#TRANSFORM H:MM INTO DAYS/HOURS/MINUTES
def make_delta(entry):
    h, m = entry.split(':')
    return datetime.timedelta(hours=int(h), minutes=int(m))


def open_file(filename):
    df = pd.DataFrame(pd.read_excel(filename))
    return(df)

def group_by_operator(df):
    # show only operators and login time
    df = df[['Unnamed: 0', 'Unnamed: 4']]
    df.columns = ["IND. OP.", "HL (h)"]
    pd.set_option('display.max_rows', None, 'display.max_columns', None)
    #remove nan
    df = df[df['HL (h)'].notnull()]
    #remove Horas (HH:MM) line
    df = df[df['HL (h)'] != "Horas (HH:MM)"]
    #transform STR into datetime
    df['HL (h)'] = df[['HL (h)']].applymap(lambda entry: make_delta(entry))
    # group hours by operator
    df = df.groupby('IND. OP.')[['HL (h)']].sum()
    #format date time into hours
    df['HL (h)'] = df[['HL (h)']].applymap(lambda entry: time_formatter(entry))
    print(df)
    return df


def create_output(df):
    now = datetime.datetime.now()
    dt_string = now.strftime("%d-%m-%Y_%H-%M-%S")
    filepath = f'Relatorio Agrupado {dt_string}.xlsx'
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(filepath, engine='xlsxwriter',
                            engine_kwargs={'options': {'strings_to_numbers': True}})
    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    return filepath


def select_file():
    filetypes = (
        #file types shown on the file search
        ('XLS/XLSX files', '*.xls'),
        ('XLS/XLSX files', '*.xlsx'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Escolha um arquivo',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='Arquivo Selecionado',
        message=filename
    )
    file.set(filename)
    return filename


#open window
window = Tk()
window.geometry("360x540+700+300")
window.title("Script Agrupar Carga de Console")
btn_selectfile = Button(text='Iniciar', command=execute)
btn_selectfile.grid(row=0, column=0, pady=20)
file = StringVar()
lbl_filedesc = Label(window, text="Arquivo selecionado: ")
lbl_filedesc.grid(row=1, column=0, pady=50)
lbl_filename = Label(window, textvariable=file)
lbl_filename.grid(row=2, column=0, pady=50)
lbl_instructions = Label(window, text="Instruçoes:\n"
                                      "1) Selecione a base de dados do PENTAHO \n"
                                      "2)Selecione a base de dados do mês de destino.(Modelo na pasta)\n"
                                      "3) Um novo arquivo é gerado na pasta onde está o script.")
lbl_instructions.grid(row=3, column=0, pady=50)
lbl_dev = Label(window, text="Script desenvolvido por: Eduardo Ferrer", font=("Times New Roman", 8))
lbl_dev.grid(row=4, column=0, pady=50)


#maintain window opened
window.mainloop()
