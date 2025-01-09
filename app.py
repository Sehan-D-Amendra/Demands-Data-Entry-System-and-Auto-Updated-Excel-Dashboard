from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

# Path to the Excel file
EXCEL_FILE = 'transportation_data.xlsx'

# Check if the Excel file exists, if not create it with headers
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=[
        'Route Name', 'Vehicle ID', 'City', 'Shop Name', 
        'Shirt Type 1 Quantity', 'Shirt Type 2 Quantity', 
        'Shirt Type 3 Quantity', 'Shirt Type 4 Quantity',
        'Shirt Type 5 Quantity', 'Shirt Type 6 Quantity', 
        'Date', 'Time'])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Collect data from the form
    route_name = request.form['route_name']
    vehicle_id = request.form['vehicle_id']
    city = request.form['city']
    shop_name = request.form['shop_name']
    shirt_type_1 = request.form['shirt_type_1']
    shirt_type_2 = request.form['shirt_type_2']
    shirt_type_3 = request.form['shirt_type_3']
    shirt_type_4 = request.form['shirt_type_4']
    shirt_type_5 = request.form['shirt_type_5']
    shirt_type_6 = request.form['shirt_type_6']

    # Automatically generate current date and time
    current_date = datetime.now().strftime('%Y-%m-%d')
    current_time = datetime.now().strftime('%H:%M:%S')

    # Create a DataFrame with the new data
    new_data = pd.DataFrame([[route_name, vehicle_id, city, shop_name,
                              shirt_type_1, shirt_type_2, shirt_type_3, 
                              shirt_type_4, shirt_type_5, shirt_type_6,
                              current_date, current_time]],
                            columns=['Route Name', 'Vehicle ID', 'City', 'Shop Name', 
                                     'Shirt Type 1 Quantity', 'Shirt Type 2 Quantity', 
                                     'Shirt Type 3 Quantity', 'Shirt Type 4 Quantity',
                                     'Shirt Type 5 Quantity', 'Shirt Type 6 Quantity',
                                     'Date', 'Time'])

    # Append data to the existing Excel file
    with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='overlay') as writer:
        existing_data = pd.read_excel(EXCEL_FILE)
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
        updated_data.to_excel(writer, index=False)

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
