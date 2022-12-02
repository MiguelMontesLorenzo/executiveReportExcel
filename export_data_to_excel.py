import clean_transform_data

#generate graphs
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

#saving images process
import os

#dataframe treatment modules
import pandas as pd
import numpy as np

#other
import datetime
import xlsxwriter
import itertools






def change_day_format(date):
    date_time_obj = datetime.datetime.strptime(date, '%Y-%m-%d')
    return date_time_obj.strftime('%d-%m-%Y')



if __name__ == '__main__':


    dfs = clean_transform_data.extract()
    df_ingredients = dfs[-1]
    dfs_pizzas = clean_transform_data.data_cleaning(dfs[1:5])
    dfs_Product_Details = clean_transform_data.dataframe_for_Product_Details_graphs(dfs_pizzas, df_ingredients)


    #DEFINE THE EMPLOYEE WAGES:  (Asuming we have 1 cook and 2 waiters)
    #average cook wage (in USA): 31,630 $
    #average waiter wage (in USA): 29,010 $
    total_employee_wages = 31630+2*29010

    #DEFINE CORPORATE FEDERAL INCOME TAX:
    #feceral corporate income tax: 21%
    #average state corporate income tax: 3.38 %
    tax_percentage = 21+3.38

    df_metrics, df_metrics_by_month, df_week_day = clean_transform_data.dataframe_for_Executive_Summary_graphs( df_ingredients_costs = df_ingredients,
                                                                                                                dfs = dfs_pizzas,
                                                                                                                tax_percentage = tax_percentage,
                                                                                                                total_employee_wages = total_employee_wages)



    df_product_details = dfs_Product_Details[0]
    df_product_details_by_size = dfs_Product_Details[1]
    df_product_details_by_pizza_type_id = dfs_Product_Details[2]
    df_product_details_by_category = dfs_Product_Details[3]
    df_pizza_marginal_profits = dfs_Product_Details[4]




    df_fig5 = df_metrics[['date','cumulative_sales', 'cumulative_costs', 'gross_profit']]
    df_fig6 = df_metrics_by_month[['month', 'sales']]
    df_fig8 = df_week_day[['weekday', 'sales']]
    
    dfs_to_exel = { 'product_details_by_size':df_product_details_by_size,
                    'product_details_by_category':df_product_details_by_category,
                    'product_details':df_product_details,
                    'pizza_marginal_profits':df_pizza_marginal_profits,
                    'metrics':df_metrics,
                    'metrics_by_month':df_metrics_by_month,
                    'week_day':df_week_day
                    }




    saving_path = 'excel_data'

    if not os.path.exists(saving_path):
        os.mkdir(saving_path)

    else:
        try:
            os.remove('excel_data')
            os.mkdir(saving_path)
        except:
            finish = False
            saving_path += '_1'
            counter = 1

            while not finish:

                if not os.path.exists(saving_path):
                    os.mkdir(saving_path)
                    finish = True

                else:
                    while not saving_path[-1] == '_':
                        saving_path = saving_path[:-1]

                    counter += 1
                    saving_path = saving_path + str(counter)



    #Create excel
    saving_path = saving_path + '/'
    file_name = 'report.xlsx'
    file_path = saving_path + file_name
    writer = pd.ExcelWriter(file_path, engine='xlsxwriter')

    #Get the workbook object
    workbook = writer.book

    #Create the workbook sheets
    for key in dfs_to_exel.keys():
        #get data_frame
        df = dfs_to_exel[key]

        #CHANGE DAY FORMAT (df)
        if 'date' in df.keys():
            df['date'] = df['date'].apply(change_day_format)
  
        #save df as exel
        df.to_excel(writer, sheet_name = key)






    #Work individually with each work-sheet to create pertinent graph
    #If graph is too complex just import it as an image from plotly


    #GRAPH 1 - product_details_by_size
    #Type: barplot
    worksheet_name = 'product_details_by_size'
    chart1 = workbook.add_chart({'type': 'column'})
    # Configure the series of the chart from the dataframe data.
    data_column = 2
    chart1.add_series({
        'name':       [worksheet_name, 0, data_column],
        'categories': [worksheet_name, 1, 1, 5, 1],# + list(itertools.chain.from_iterable([[i,0] for i in range(1,5)])),
        'values':     [worksheet_name, 1, data_column, 5, data_column],# + list(itertools.chain.from_iterable([[i,0] for i in range(1,5)])),
        'gap':        300,
    })

    # Configure the chart axes.
    chart1.set_y_axis({'major_gridlines': {'visible': False}})

    # Insert the chart into the worksheet.
    writer.sheets[worksheet_name].insert_chart('D1', chart1)





    # #GRAPH 2 - product_details_by_category
    # #Type: barplot
    worksheet_name = 'product_details_by_category'
    chart2 = workbook.add_chart({'type': 'pie'})

    # # Configure the series of the chart from the dataframe data.
    chart2.add_series({
        'name':       [worksheet_name, 0, data_column],
        'values':     [worksheet_name, 1, data_column, 4, data_column],
        'categories': [worksheet_name, 1, 1, 4, 1],
        'data_labels': {'percentage': True},
    })

    # # Configure the chart axes.
    # chart.set_y_axis({'major_gridlines': {'visible': False}})

    # # Insert the chart into the worksheet.
    writer.sheets[worksheet_name].insert_chart('D1', chart2)





    # #GRAPH 3 - product_details
    # #Type: stacked barplot
    worksheet_name = 'product_details'

    #generate with plotly and insert in 

    fig = px.bar(  data_frame=df_product_details,
                    y='quantity',
                    x='pizza_type_id',
                    color='size',
                    text='quantity',
                    labels={
                        "quantity": "Orders",
                        "pizza_type_id": "Pizza flavor",
                        "size": "Sizes"
                    })
    fig.update_traces(texttemplate='%{text:.2s}', textposition='outside')
    fig.update_layout( title="PIZZA TYPE'S POPULARITY AND SIZES",
                        uniformtext_minsize=10,
                        uniformtext_mode='hide',
                        xaxis={'categoryorder': 'total ascending'},
                        autosize=False, width=1200, height=600)

    fig.write_image(saving_path + f'{worksheet_name}.jpg')
    writer.sheets[worksheet_name].insert_image('G1', saving_path + f'{worksheet_name}.jpg')




    #GRAPH 4 - product_details_by_size
    #Type: barplot
    worksheet_name = 'pizza_marginal_profits'

    chart4 = workbook.add_chart({'type': 'column'})
    colors = ['#29E542', '#D6453E', '#EFDA4F']
    # Configure the series of the chart from the dataframe data.
    for data_column in range(2,5):
        chart4.add_series({
            'name':       [worksheet_name, 0, data_column],
            'categories': [worksheet_name, 1, 1, 31, 1],
            'values':     [worksheet_name, 1, data_column, 31, data_column],
            'fill':       {'color':  colors[data_column - 2]},
            'gap':        300,
        })

    # Insert the chart into the worksheet.
    writer.sheets[worksheet_name].insert_chart('K1', chart4)




    #GRAPH 5 - product_details_by_size
    #Type: barplot
    worksheet_name = 'metrics'

    chart5 = workbook.add_chart({'type': 'line'})
    colors = ['#29E542', '#D6453E', '#EFDA4F']
    # Configure the series of the chart from the dataframe data.
    for data_column in range(4,7):
        chart5.add_series({
            'name':       [worksheet_name, 0, data_column],
            'categories': [worksheet_name, 1, 1, 358, 1],
            'values':     [worksheet_name, 1, data_column, 358, data_column],
            'fill':       {'color':  colors[data_column - 4]},
            'gap':        300,
        })

    # Insert the chart into the worksheet.
    writer.sheets[worksheet_name].insert_chart('K1', chart5)



    #GRAPH 6 - metrics_by_month
    #Type: barplot
    worksheet_name = 'metrics_by_month'

    chart6 = workbook.add_chart({'type': 'column'})
    colors = ['#29E542', '#D6453E']
    # Configure the series of the chart from the dataframe data.
    for data_column in range(2,4):
        chart6.add_series({
            'name':       [worksheet_name, 0, data_column],
            'categories': [worksheet_name, 1, 1, 12, 1],
            'values':     [worksheet_name, 1, data_column, 12, data_column],
            'fill':       {'color':  colors[data_column - 2]},
            'gap':        300
        })

    # Insert the chart into the worksheet.
    writer.sheets[worksheet_name].insert_chart('K1', chart6)




    #GRAPH 7 - week_day
    #Type: barplot
    worksheet_name = 'week_day'

    chart7 = workbook.add_chart({'type': 'column'})
    # Configure the series of the chart from the dataframe data.
    data_column = 2
    chart7.add_series({
        'name':       [worksheet_name, 0, data_column],
        'categories': [worksheet_name, 1, 1, 6, 1],
        'values':     [worksheet_name, 1, data_column, 6, data_column],
        'gap':        300,
    })

    # Insert the chart into the worksheet.
    writer.sheets[worksheet_name].insert_chart('K1', chart7)





    # Close the Pandas Excel writer and output the Excel file.
    workbook.close()