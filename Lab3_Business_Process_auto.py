from sys import argv, exit
import os
from datetime import date
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv() 
    order_dir = create_order_dir()
    return 

#Get Path of Sales Data Files CSV
def get_sales_csv():

#Check command line parameter
    num_params = len(argv) - 1
    if num_params >= 1: 
        sales_csv = argv[1]
        #Check Provider Parameter for valid Path 
        if os.path.isfile(sales_csv):
            return sales_csv
        else: 
            print('Error: Invalid Path to sales data CSV file')
            exit(1)
    else: 
        print('Error: Missing Path to sales data CSV File path')
        exit(1)
    #return 

def create_order_dir(sales_csv): 

#Get Directory in sales data CSV file resides 
    sales_dir = os.path.dirname(os.path.abspath(sales_csv))

#Determine path of the directory to hold order files 
    todays_date = date.today().isoformat()
    orders_dir = os.path.join(sales_dir, f'Orders_{todays_date}')
#Create other Directory if it exist 
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)

    return orders_dir

#Split data into individual order and save it to Excel sheets 
def process_sales_data(sales_csv, orders_dir):
#Import sales data from CSV File 
    sales_df = pd.read_csv(sales_csv)
#Insert a new colum "TOTAL PRICE"
    sales_df.insrt(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])
#remove colums from the dataframe 
    sales_df.drop(columns=['ADDRESS','CITY','STATE','POSTAL CODE', 'COUNTRY'], inplace=True)

#Group Sales Data and save
    for order_id, order_db in sales_df.groupby('ORDER ID'):
        # remove the 'order n d' colum 
        order_df.drop(columns=['ORDER ID'], inplace=True)
        #sort order by item number 
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        # add the grand total row 
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df =pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]}) 
        order_df = pd.concat([order_df, grand_total_df])

        export_order_to_excel(order_id, order_df, orders_dir)

        break
#return 
def export_order_to_excel(order_df, order_id, order_dir):
    # Determine File and Path of Order Excel sheet
    customer_name = order_df['CUSTOMER NAME'].values[0] 
    customer_name = re.sub(r'\W','', customer_name) 
    order_file = f'Order{order_id}_{customer_name}.xlsx' 
    order_path = os.path.join(order_dir, order_file) 
    
    sheet_name = f'order#{order_id}' 
    order_df.to_excel(order_path, index=False, sheet_name=sheet_name) 
#return 
if '__name__' == '__main___':
    main()