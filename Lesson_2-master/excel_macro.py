import pandas as pd
import openpyxl
import os

def automation():
    print('Start the automation...')

    #Update the os.getcwd() to the dirctory where the excel (or CSV) is saved
    original_path = os.path.join(r'C:\Users\hilton.netta\py\Lesson_2-master\','Summary_12-23-2020.xlsx')
    # original_path = os.path.join(os.getcwd(),'Summary_12-23-2020.xlsx')
    rates_path = os.path.join(r'C:\Users\hilton.netta\py\Lesson_2-master\', 'Inputs_12-24-2020.xlsx')
    # rates_path = os.path.join(os.getcwd(),'Inputs_12-24-2020.xlsx')

    #Creates a df - treasury rates
    df_rates = pd.read_excel(rates_path, sheet_name='Rates', engine='openpyxl', index_col=0)
    #df_rates = pd.read_csv(rates_path) to read the same file in csv format

    #Sets val_date
    #This is due to the way datetime values are stored in pandas: using the numpy datetime64[ns] dtype. 
    val_date = df_rates.columns[0].strftime('%m-%d-%Y')

    print(df_rates)
    print(val_date)
    quit()
    #Opens the summary_sheet
    summary_workbook = openpyxl.load_workbook(filename = original_path)
    input_sheet = summary_workbook['Inputs']

    input_sheet['B2'].value = val_date

    row_max = len(df_rates.index)

    for i in range(0, row_max):
        #print(i)
        #print(df_rates.iloc[i, 0])          #row, column
        input_sheet.cell(i + 9, 3).value = df_rates.iloc[i, 0]      #row, columns

    destination = os.path.join(os.getcwd(), r'Summary_{}.xlsx'.format(val_date))

    summary_workbook.save(filename = destination)
    print('workbook saved as {}'.format(destination))
    print('Success!')

#Set special variable before running the code
#If this python file is being imported
# Python is able to call function from different python files (.py),
# So it's best practice to keep "if __name__ == '__main__':"
if __name__ == '__main__':
    automation()

