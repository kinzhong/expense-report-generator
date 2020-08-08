import pandas as pd
import dateutil
import numpy as np
import sys
from dateutil.relativedelta import relativedelta
import xlsxwriter


if (len(sys.argv)==2):
    print("Reading: ", str(sys.argv[1]))

    #Excel Files
    writer = pd.ExcelWriter('expenses-report.xlsx', engine='xlsxwriter', datetime_format='dd/mm/yyyy')
    workbook  = writer.book

    #Excel Formatting Styles
    left = workbook.add_format({'align': 'left'})
    title_format = workbook.add_format({'bold': True})
    title_format.set_font_size(17)
    subheader_format = workbook.add_format({'bold': True})
    subheader_format.set_font_size(12)
    
    #Expenses Data
    data = pd.read_excel(str(sys.argv[1]))
    expenses = data.iloc[21:,:6]
    expenses.columns = data.iloc[20,:6]
    expenses.columns = [c.replace(' ', '_') for c in expenses.columns]

    #Credit Card Info
    cc_info = data.iloc[1:4,7:9]
    cc_info.columns = data.iloc[0,7:9]
    cc_info.columns = [c.replace(' ', '_') for c in cc_info.columns]
    cc_list = expenses.Payment_Method.unique()

    for cc in cc_list:

        print("Generating report for ", cc)
        condition = cc_info["Credit_Card"] == cc
        bill_cycle = cc_info[condition].iloc[0,1]
        bill_cycle.replace("Every ", "")
        
        
        condition = expenses["Payment_Method"] == cc
        #print(expenses[condition].to_string())
        cc_data = expenses[condition]
        cc_data = cc_data.iloc[:,[0,1,2,3,5]]
        
        first_date = cc_data['Date'].iloc[0]
        last_date = cc_data['Date'].iloc[-1]
        
        next_bill_cycle = first_date + relativedelta(months=1)
        next_bill_cycle = next_bill_cycle.replace(day=12)
        #print(next_bill_cycle)
        
        last_bill_cycle = last_date + relativedelta(months=1)
        last_bill_cycle = last_bill_cycle.replace(day=12)
        #print(last_bill_cycle)
        
        index = 0
        cc_data_copy = cc_data.copy()

        combined_list = []
        while next_bill_cycle <= last_bill_cycle:
            #do sth
            #print(index,": ",next_bill_cycle.strftime("%d/%m/%Y"))

            curr_cycle = cc_data_copy.loc[cc_data_copy['Date'] < next_bill_cycle]
            cc_data_copy = cc_data_copy.loc[cc_data_copy['Date'] >= next_bill_cycle]
            curr_cycle[''] = np.nan
            curr_cycle['Total Amount'] = np.nan
            curr_cycle['NIL2'] = np.nan
            total_sum = curr_cycle['Amount'].sum()
            curr_cycle.rename(columns = {'NIL2':str(total_sum)}, inplace = True)

            
            #print(list(curr_cycle.columns.values))
            #curr_cycle = curr_cycle.append(list(curr_cycle.columns.values), ignore_index=True)

            curr_cycle = curr_cycle.reset_index(drop=True)
            curr_cycle.loc[-1] = list(curr_cycle.columns.values)
            
            start_date = next_bill_cycle - relativedelta(months=1)
            end_date = next_bill_cycle - relativedelta(days=1)
            curr_cycle.loc[-2] = ["Billing Cycle:", str(start_date.strftime("%d/%m/%Y")) + " to " + str(end_date.strftime("%d/%m/%Y")), "", "", "", "", "", ""]
            curr_cycle.index = curr_cycle.index + 2
            curr_cycle = curr_cycle.sort_index()
            curr_cycle.loc[-1] = ['', '', '', '', '', '', '', '']
 
            inc = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
            curr_cycle.columns = inc
            
            #print(curr_cycle.to_string())
            combined_list.append(curr_cycle)

            
            next_bill_cycle = next_bill_cycle + relativedelta(months=1)
            index = index + 1

        title = {'': [cc]}
        title = pd.DataFrame(data=title)
        #print(df)
        title.to_excel(writer, sheet_name=cc, header=False, index=False, startrow=0)
        results = pd.concat(combined_list, ignore_index=True)
        #print(results.to_string())
        results.to_excel(writer, sheet_name=cc, header=False, index=False, startrow=2)

        worksheet = writer.sheets[cc]                                    
        worksheet.set_column('A:A', 12, left)
        worksheet.set_column('B:D', 20, left)
        worksheet.set_column('E:E', 33, left)
        worksheet.set_column('F:H', 20, left)
        worksheet.write('A1', cc, title_format)
        
        #find indexs of titles
        billing_condition = results['A'] == "Billing Cycle:"
        #billing = results[billing_condition]
        billing_indexes = results.index[results['A'] == "Billing Cycle:"].tolist()
        for i in billing_indexes:
            worksheet.set_row(i+2,None, subheader_format)
            worksheet.set_row(i+3,None, subheader_format)
        

    #Repayment
    print("Generating report for Debt Collection")
    debtor_list = expenses.Payee.unique()

    index = 1
    for debtor in debtor_list:
        if (debtor == "NIL"):
            break
        
        condition = expenses["Payee"] == debtor
        #print(expenses[condition].to_string())
        debtor_data = expenses[condition]

        if (debtor == "Shared"):
            total_sum = debtor_data["Amount"].sum()/2
        else:
            total_sum = debtor_data["Amount"].sum()

        debtor_data = debtor_data.reset_index(drop=True)
        debtor_data.loc[-1] = ['Total Sum:', str(total_sum), '', '', '', '']
        debtor_data = debtor_data.reset_index(drop=True)
        debtor_data.to_excel(writer, sheet_name="Debt Collection", index=False, startrow=index)
        debt_worksheet = writer.sheets["Debt Collection"]
        debt_worksheet.write(str("A"+str(index)), debtor, title_format)
        
        
        index = index + debtor_data.shape[0] + 3

        debt_worksheet.set_row(index-3, None, subheader_format)
        debt_worksheet.set_column('A:C', 11, left)
        debt_worksheet.set_column('D:F', 23, left)
        debt_worksheet.set_column('E:E', 35, left)
        
    writer.save()

elif (len(sys.argv)<2):
    print("Input excel file as an argument")

else:
    print("Too many arguments!")
