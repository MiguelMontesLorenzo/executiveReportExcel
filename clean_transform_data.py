#EXERCICE 2 - DATA CLEANING + ETL

#IMPORT MODULES
import numpy as np
import pandas as pd
import re
import datetime
#import xml.etree.ElementTree as ET


#IMPORT CSV's AS PD.DATAFRAMES (EXTRACT)
def extract():
    #Description of each field of the following datasets
    data_description = pd.read_csv('maven_pizzas/data_dictionary.csv')

    #Each order might have more than 1 pizza  (On average, there's ~ 2 pizzas by order)
    unordered_pizzas_details = pd.read_csv('maven_pizzas/order_details_2016.csv', sep=';')    #descrives number of pizzas ordered by its pizza_id 
    orders_details = pd.read_csv('maven_pizzas/orders_2016.csv', sep=';')                   #descrives date of each order

    # (pizza_id) descrives pizza type and size
    # (pizza_type_id) descrives just pizza type
    pizzas_ingredients = pd.read_csv('maven_pizzas/pizza_types.csv', encoding='latin1')    # descrives ingredients contained in each pizza_type_id
    pizza_type_price = pd.read_csv('maven_pizzas/pizzas.csv')                                 # associates pizza_id with correspondent pizza_type_id and price

    ingredients_quantity_price = pd.read_csv('maven_pizzas/ingredients.csv')
    
    return [data_description, unordered_pizzas_details, orders_details, pizzas_ingredients, pizza_type_price, ingredients_quantity_price]








#DATA QUALITY
def data_quality(dfs):

    for i, df in enumerate(dfs):

        #obtain data types
        df_types = df.dtypes

        #obtain number of null or NaN values
        df_nulls = df.isnull().sum()

        #concatenate both pd.series into dataframe
        df_quality = pd.concat([df_types, df_nulls], axis=1, keys=['data_types', 'null_values'])

        dfs[i] = df_quality


    dfs_quality = pd.concat(dfs)

    return dfs_quality






#DATA CLEANING
def correct_quantities(x):
    replacements = {1:1, '1':1, 'one':1, 'One':1, -1:1, '-1':1,
                    2:2, '2':2, 'two':2, 'Two':2, -2:2, '-2':2,
                    3:3, '3':3, 'three':3, 'Three':3, -3:3, '-3':2,
                    4:4, '4':4, 'four':4, 'Four':4, -4:4, '-4':4,}
    return replacements[x]

def correct_pizza_IDs(string):
    string = string.lower().replace(' ','_').replace('-','_')
    string = string.replace('0','o').replace('3','e').replace('@','a')
    return string

def to_datetime_format(date):
    date = str(date)
    date = date.replace(',',', ')
    date = ' '.join(date.split())



    format_1 = re.compile(r'\d{1,2}\s\w+\s\d{4}')
    format_2 = re.compile(r'\d{2}-\d{2}-\d{4}')
    format_3 = re.compile(r'\d{8}$')
    format_4 = re.compile(r'\d{4}/\d{2}/\d{2}')
    format_5 = re.compile(r'\d{4}-\d{2}-\d{2}$')
    format_6 = re.compile(r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}')
    format_7 = re.compile(r'\d{2}-\d{2}-\d{2} \d{2}:\d{2}$')
    format_8 = re.compile(r'\w+\s\d{1,2}\s\d{4}')
    format_9 = re.compile(r'\w{3}\w+,\s\d{1,2}\s\w+,\s\d{4}')
    format_10 = re.compile(r'\w+\s\d{1,2}-\w{3}\w+-\d{4}')
    format_11 = re.compile(r'\d{2}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}')
    format_12 = re.compile(r'\w{3}\s\d{1,2}\s\w{3}\s\d{4}')
    # format_13 = re.compile(r'\d{8}\d+')
    formats = { format_1:'%d %b %Y',
                format_2:'%d-%m-%Y',
                format_3:'%Y%m%d',
                format_4:'%Y/%m/%d',
                format_5:'%Y-%m-%d',
                format_6:'%Y-%m-%dT%H:%M',
                format_7:'%d-%m-%Y %H:%M',
                format_8:'%b %d %Y',
                format_9:'%A, %d %B, %Y',
                format_10:'%a %b-%d-%Y',
                format_11:'%d-%m-%y %H:%M:%S',
                format_12:'%a %d %b %Y'
              }

    # %Y year (sirve si el año está expresado en 2 dígitos)
    # %m month
    # %d day
    # %H hour
    # %M minute
    # %S second
    # %I hour in (avant/post meridiem) format  (Comes always with %p [AM/PM] at some place)
    # %a day of the week

    #Encuentra el formato que coincide con la fecha dada
    date_time_format = None
    for _,pattern in enumerate(formats):

        if (bool(re.search(pattern, date))):
            date_time_format = formats[pattern]

    #Si se ha encontrado el formato se reformatea la fecha, si no, se devuelve un NaN
    if date_time_format == None:
        return np.nan
    else:
        date_time_obj = datetime.datetime.strptime(date, date_time_format)
        return date_time_obj.strftime('%Y-%m-%d')


def to_time_format(time):

    format_1 = re.compile(r'\d{2}H \d{2}M \d{2}S')
    format_2 = re.compile(r'\d{2}:\d{2}\s\w{2}')
    format_3 = re.compile(r'\d{2}:\d{2}:\d{2}')
    formats = { format_1:'%HH %MM %SS',
                format_2:'%H:%M %p',
                format_3:'%H:%M:%S',
              }

    #Encuentra el formato que coincide con la hora dada
    for _,pattern in enumerate(formats):
        if (bool(re.search(pattern, time))):
            date_time_format = formats[pattern]

    #Reformatea la hora
    date_time_obj = datetime.datetime.strptime(time, date_time_format)
    return date_time_obj.strftime('%H:%M')


def data_cleaning(lsts):

    unordered_pizzas_details, orders_details, pizzas_ingredients, pizza_type_price = lsts[0], lsts[1], lsts[2], lsts[3]

    #Order dataframe by ID's
    ordered_pizzas_details = unordered_pizzas_details.sort_values(by='order_details_id')
    orders_details = orders_details.sort_values(by='order_id')
    
    #DATASET 1: oder_details
    #column: quantity
    #Replace nulls with 1's in 'quantity' column
    ordered_pizzas_details = ordered_pizzas_details.fillna(value={'quantity': 1})
    #Correct the rest of values in 'quantity' column
    ordered_pizzas_details['quantity'] = ordered_pizzas_details['quantity'].apply(correct_quantities)

    #columna: pizza_id
    #Eliminate rows without ID
    ordered_pizzas_details = ordered_pizzas_details.dropna()
    #Correct the rest of values in 'pizza_id' column
    ordered_pizzas_details['pizza_id'] = ordered_pizzas_details['pizza_id'].apply(correct_pizza_IDs)


    #DATASET 2: orders.csv
    orders_details = orders_details.fillna(method='bfill').fillna(method='ffill')

    orders_details['date'] = orders_details['date'].apply(to_datetime_format)
    orders_details['time'] = orders_details['time'].apply(to_time_format)

    #los que no haya sido posible convertir son eliminados
    orders_details = orders_details.fillna(method='bfill').fillna(method='ffill')


    
    
    #assuming {type of pizza ordered by people} and {date} are independent events
    #we don't need to study pizza_id's distribution throughout the year
    
    return [ordered_pizzas_details, orders_details, pizzas_ingredients, pizza_type_price]
    

    


    
#OBTAIN INGREDIENT QUANTITIES (TRANSFORM)

def ponderate_quatity_by_size(args):
    ponderations = {'S':0.75 , 'M':1, 'L':1.25, 'XL':1.5, 'XXL':1.75}
    ponderation = ponderations[args[2]]

    return ponderation*args[4]


def multiply_by(x, factor):
    return x*factor


def transform(lsts):

    ordered_pizzas_details, orders_details, pizzas_ingredients, pizza_type_price = lsts[0], lsts[1], lsts[2], lsts[3]

    #Obtain number of orders of each pizza_id
    number_pizzas_ordered_by_ID = ordered_pizzas_details[['pizza_id','quantity']].groupby('pizza_id').sum(numeric_only=False)   #pd.series
    df_temp = number_pizzas_ordered_by_ID.reset_index(level=0)


    #Incorporate quantity (of [pizza_ID pizza] ordered) to pizza_type_price dataframe
    pizza_type_price_quantity = pizza_type_price.merge(df_temp, on='pizza_id', how = 'inner')
    pizza_type_price_quantity['ponderated_quantities'] = pizza_type_price_quantity.apply(ponderate_quatity_by_size, axis = 1)

    #Obtain ponderated quantities of each pizza_type_id  (According to the sizes 'S':0.75, 'M':1, 'L':1.25, 'XL':1.5, 'XXL':1.75   of pizza_id)
    df_temp = pizza_type_price_quantity.groupby('pizza_type_id').sum(numeric_only=True)['ponderated_quantities'].to_frame()
    pizzas_ingredients = pizzas_ingredients.merge(df_temp, on='pizza_type_id', how = 'inner')

    #Obtain ponderated quantities of each ingredient
    all_ingredients = dict()
    for index, row in pizzas_ingredients.iterrows():

        row_ingredients = row['ingredients'].split(',')

        #add Tomatoe if it has not sauce
        has_sauce = False
        for ingredient in row_ingredients:
            if 'sauce' in ingredient.lower() or 'tomatoes' in ingredient.lower():
                has_sauce = True

        if not has_sauce:
            row_ingredients.append('Tomatoes')

        #add Mozzarella if it has not
        if not 'Mozzarella Cheese' in row_ingredients:
            row_ingredients.append('Mozzarella Cheese')


        for i in row_ingredients:
            i = i.strip()
            if (i in all_ingredients.keys()):
                all_ingredients[i] += row['ponderated_quantities']
            else:
                all_ingredients[i] = row['ponderated_quantities']

    #Transformation to diccionary to operate more easily
    ingredients_ponderated_quantities = pd.DataFrame.from_dict(all_ingredients, orient='index')
    ingredients_ponderated_quantities.rename({0: 'quantity'}, axis=1, inplace=True)


    #NOW WE CAN STUDY HOW COUD WE MAKE THE OPTIMAL INGREDIENT ACQUIREMENTS
    #Acording to CRITERIA: Weekly ACQUIREMENTS
    # -> We must calculate how much ingredient quantities we sould adquire in order to finish them all by the end of the week

    ingredients_ponderated_quantities_criteria1 = ingredients_ponderated_quantities
    adq_period = 7 #days  #(acquirement period)
    hole_period = 365  #days

    #we obtain igredient usage per week
    ingredients_ponderated_quantities_criteria1['per_week'] = (ingredients_ponderated_quantities_criteria1['quantity']*(adq_period/hole_period)).round(2)

    return ingredients_ponderated_quantities_criteria1




    
    
    


def add_costs_per_weeks(df_with_dates, df_ingredients_costs, df_ingredients_acquirements, monthly_employee_wages):

    #DATA:
    #ASUMPTION: Each ingredient is going to be used in equal quantities in every pizza
    #(df_ingredients_costs) contains [quantity of ingredient used in pizza] and [price of 1 kilogram of that ingredient]
    #(df_ingredients_acquirements) contains [anual ingredient quantity adquired] and [weekly ingredient quantity adquired]

    #GOAL:
    #df_costs_evolution) contains [date], [cost each day (0 except mondays, as we are attempting weekly acquirements)] and [cumulative cost]



    #1 - OBTAIN INIGREDIENT COSTS PER PIZZA
    #AS IN (df_ingredients_acquirements) WE ARE MEASSURING ingredient quantities IN amount used in medium pizza
    #IN ORDER TO OBTAIN ingredient quantities costs WE NEED TO USE (AS COSTS ARE RELATED TO kg's) RELATION BETWEEN kg' AND amount used in medium pizza:
    #Logic: ingredient[cost/pizza] = ingredient[cost/kg]*ingredient[kg/pizza]

    df_ingredients_costs['ingredient_cost_per_pizza'] = (1/1000)*df_ingredients_costs['quantity_per_pizza']*df_ingredients_costs['price_kg']
    
    df_ingredients_costs = df_ingredients_costs.set_index('ingredient')
    df_ingredients_costs_per_pizza = df_ingredients_costs['ingredient_cost_per_pizza']


    #2 - OBTAIN WEEKY COSTS
    df_ingredients_acquirements = df_ingredients_acquirements.join(df_ingredients_costs_per_pizza)
    df_ingredients_acquirements['cost_per_week'] = df_ingredients_acquirements['per_week']*df_ingredients_acquirements['ingredient_cost_per_pizza']


    #3 - OBTAIN DAILY COSTS
    costs_per_week = df_ingredients_acquirements.sum(numeric_only=True)['cost_per_week']

    def add_costs(date, costs_per_week, monthly_employee_wages):

        #En realidad febrero tendría 29 días porque 2026 es año bisiesto, pero el dataset está mal
        days_per_month = {1:31, 2:28, 3:31, 4:30, 5:31, 6:30, 7:31, 8:31, 9:30, 10:31, 11:30, 12:31}
        date_time_obj = datetime.datetime.strptime(date, '%Y-%m-%d')
        day_of_the_year = date_time_obj.day + sum([days_per_month[month] for month in range(1,date_time_obj.month)])

        daily_costs = 0
        daily_costs += costs_per_week if (day_of_the_year-1)%7 == 0 else 0
        daily_costs += monthly_employee_wages if date_time_obj.day == days_per_month[date_time_obj.month] else 0

        return daily_costs


    df_with_dates['costs'] = df_with_dates['date'].apply(add_costs, args=(costs_per_week, monthly_employee_wages))
    
    return df_with_dates







def dataframe_for_Executive_Summary_graphs(df_ingredients_costs, dfs, tax_percentage, total_employee_wages):
    ordered_pizzas_details, orders_details, pizzas_ingredients, pizza_type_price  = dfs[0], dfs[1], dfs[2], dfs[3]
    #add ing costs

    #JUTAR TODOS LOS DATOS NECESARIOS EN UN SOLO DATAFRAME
    df_metrics_evolution = ordered_pizzas_details.merge(orders_details, on='order_id', how = 'inner').drop('time', axis=1)
    df_metrics_evolution = df_metrics_evolution.merge(pizza_type_price, on='pizza_id', how = 'inner').sort_values(by='date')

    #SI UN ELIMINAR EN LUGAR DE GUARDAR EL PRECIO DEL TIPO DE PIZZA PEDIDO, ALMACENAR (PRECIO_TIPO_PEDIDO*NÚMERO_PIZZAS_PEDIDAS)
    df_metrics_evolution['order_price'] = df_metrics_evolution['quantity']*df_metrics_evolution['price']
    df_metrics_evolution = df_metrics_evolution.drop(['quantity','price','pizza_type_id','pizza_id','size'], axis=1)


    #OBTENER LOS INGRESOS DIARIOS
    df_metrics_evolution = df_metrics_evolution[['order_price', 'date']].groupby(by='date').sum(numeric_only=True)
    df_metrics_evolution = df_metrics_evolution.reset_index(level=0)
    df_metrics_evolution.rename({'order_price': 'sales'}, axis=1, inplace=True)


    #OBTENER GASTOS ACUMULATIVOS
    monthly_employee_wages = total_employee_wages/12
    df_metrics_evolution = add_costs_per_weeks(df_metrics_evolution, df_ingredients_costs, transform(dfs), monthly_employee_wages)


    #OBTENER INGRESOS Y GASTOS ACUMULATIVOS
    df_metrics_evolution['cumulative_sales'] = df_metrics_evolution['sales'].cumsum()
    df_metrics_evolution['cumulative_costs'] = df_metrics_evolution['costs'].cumsum()


    #MÁRGENES BRUTOS Y MÁRGENES NETOS
    df_metrics_evolution['gross_profit'] = (df_metrics_evolution['cumulative_sales']-df_metrics_evolution['cumulative_costs'])
    df_metrics_evolution['net_profit'] = ((100-tax_percentage)/100)*df_metrics_evolution['gross_profit']



    def add_week(date):
        date_time_obj = datetime.datetime.strptime(date, '%Y-%m-%d')
        # day_of_the_year = date_time_obj.day + sum([days_per_month[month] for month in range(1,date_time_obj.month)])
        return date_time_obj.weekday()+1


    def add_month(date):
        date_time_obj = datetime.datetime.strptime(date, '%Y-%m-%d')
        return date_time_obj.month


    df_metrics_evolution['weekday'] = df_metrics_evolution['date'].apply(add_week)
    df_metrics_evolution['month'] = df_metrics_evolution['date'].apply(add_month)

    df_week_day = df_metrics_evolution.drop(['costs', 'cumulative_sales', 'cumulative_costs', 'gross_profit', 'net_profit', 'month'], axis=1).groupby('weekday', as_index=False).sum(numeric_only=True)
    weekdays = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
    df_week_day['month'] = np.array(weekdays)
    df_week_day = df_week_day.round(2)

    agg_dic = { 'sales':'sum','costs':'sum','cumulative_sales':'last','cumulative_costs':'last','gross_profit':'last','net_profit':'last'}
    df_metrics_evolution_monthly = df_metrics_evolution.drop(['date','weekday'], axis=1).groupby('month', as_index=False).agg(agg_dic)
    months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    df_metrics_evolution_monthly['month'] = np.array(months)
    df_metrics_evolution_monthly = df_metrics_evolution_monthly.round(2)


    return df_metrics_evolution, df_metrics_evolution_monthly, df_week_day






def dataframe_for_Product_Details_graphs(dfs, df_ingredients_costs):
    
    ordered_pizzas_details, orders_details, pizzas_ingredients, pizza_type_price  = dfs[0], dfs[1], dfs[2], dfs[3]

    #DATAFRAME 1
    #Necesito un dataframe con: [pizza_id, pizza_type_id, size, quantity, category]
    #Lo obtengo como un merge de [pizza_id, pizza_type_id, size] , [pizza_id, quantity] , [pizza_type_id, category]

    #DATAFRAMES 2, 3, 4
    #dataframe 1 grouped by size
    #dataframe 1 grouped by pizza_type_id
    #dataframe 1 grouped by category

    #DATAFRAME 5
    #Márgenes por tipo de pizza


    #DATAFRAME 1
    #dataframe 1.1  ->  [pizza_id, pizza_type_id, size]
    df_1_1 = pizza_type_price.drop('price', axis=1)

    #dataframe 1.2  ->  [pizza_id, quantity]
    number_pizzas_ordered_by_ID = ordered_pizzas_details[['pizza_id', 'quantity']].groupby('pizza_id').sum(numeric_only=False)
    df_1_2 = number_pizzas_ordered_by_ID.reset_index(level=0)

    #dataframe 1.3  ->  [pizza_id, category]
    df_1_3 = pizzas_ingredients[['pizza_type_id', 'category']]

    #MERGE 3 PREVIOUS DATAFRAMES
    df_temp = df_1_1.merge(df_1_2, on='pizza_id', how = 'inner')
    df_1 = df_temp.merge(df_1_3, on='pizza_type_id', how = 'inner')



    #DATAFRAME 2
    df_2 = df_1[['size','quantity']].groupby('size').sum(numeric_only = False)
    df_2 = df_2.reset_index(level=0)

    sorter = ['S', 'M', 'L', 'XL', 'XXL']
    sorterIndex = dict(zip(sorter, range(len(sorter))))
    df_2['size_Rank'] = df_2['size'].map(sorterIndex)
    df_2 = df_2.sort_values(by='size_Rank', ascending = True)
    df_2 = df_2.drop('size_Rank', axis=1)
    
    

    #DATAFRAME 3
    df_3 = df_1[['pizza_type_id','quantity']].groupby('pizza_type_id').sum(numeric_only = False)
    df_3 = df_3.reset_index(level=0)



    #DATAFRAME 4
    df_4 = df_1[['category','quantity']].groupby('category').sum(numeric_only = False)
    df_4 = df_4.reset_index(level=0)



    #DATAFRAME 5
    def round_number(number, decimals):
         return np.round(number, decimals)    

    #Obtains costs of all ingredients in passed list
    def obtain_pizza_type_cost(ingredients_string, ingredient_cost_per_pizza):

        ingredients_list = ingredients_string.split(',')

        #add Tomatoe if it has not sauce
        has_sauce = False
        for i, ingredient in enumerate(ingredients_list):
            ingredients_list[i] = ingredient.strip()
            if 'sauce' in ingredient.lower() or 'tomatoes' in ingredient.lower():
                has_sauce = True

        if not has_sauce:
            ingredients_list.append('Tomatoes')

        #add Mozzarella if it has not
        if not 'Mozzarella Cheese' in ingredients_list:
            ingredients_list.append('Mozzarella Cheese')

        return np.round(sum([ingredient_cost_per_pizza[ingredient] for ingredient in ingredients_list]),4)



    #Obtain dictionary with each ingredient cost
    df_ingredients_costs['ingredient_cost_per_pizza'] = (1/1000)*df_ingredients_costs['quantity_per_pizza']*df_ingredients_costs['price_kg']
    df_ingredients_costs['ingredient_cost_per_pizza'] = df_ingredients_costs['ingredient_cost_per_pizza'].apply(round_number, args=(4,))
    df_ingredients_costs = df_ingredients_costs.set_index('ingredient')
    ingredient_cost_per_pizza = df_ingredients_costs.to_dict()['ingredient_cost_per_pizza']


    #Get rid of non medium sized pizzas
    df_5 = pizza_type_price[pizza_type_price['size']=='M'].groupby('pizza_type_id', as_index=False).last()

    #Include ingredients column
    df_5 = df_5.merge(pizzas_ingredients[['pizza_type_id', 'ingredients']], on='pizza_type_id', how = 'inner')

    #Obtain pizza cost fron pizza ingredients
    df_5['pizza_type_cost'] = df_5['ingredients'].apply(obtain_pizza_type_cost, args=(ingredient_cost_per_pizza,))
    #Get rid of fields which I dont need
    df_5 = df_5.drop(['pizza_id', 'size', 'ingredients'], axis = 1)

    #Get margins
    df_5['pizza_marginal_profit'] = df_5['price'] - df_5['pizza_type_cost']
    df_5['pizza_marginal_profit'] = df_5['pizza_marginal_profit'].apply(round_number, args=(4,))



    return [df_1, df_2, df_3, df_4, df_5]





#EXPORT RESULTS TO CSV (LOAD)
def load(df):
    df = df.reset_index(level=0)
    df.rename({'index': 'pizza_id'}, axis=1, inplace=True)


    register = open('register.csv', 'w')

    #write the fields names
    register.write('ingredient;quantity/week\n')

    for index, row in df.iterrows():
        #print(row)
        register.write( str(row['pizza_id']) + ';' + str(row['quantity']) + '\n') 

    register.close()
















    # #Obtain number of orders of each pizza_id
    # number_pizzas_ordered_by_ID = ordered_pizzas_details.groupby('pizza_id').sum(numeric_only=False)['quantity']   #pd.series

    # #pd.series to pd.dataframe
    # df_temp = pd.DataFrame(data = [number_pizzas_ordered_by_ID.values], columns = number_pizzas_ordered_by_ID.index).T 
    # df_temp = df_temp.reset_index(level=0)
    # df_temp.rename({0: 'quantity'}, axis=1, inplace=True)


