from io import BytesIO
from scipy.interpolate import make_interp_spline
import anvil.server
import pandas as pd
import os
import io
import time
from io import StringIO
import statistics
#import win32com.client
import math
import numpy as np
from anvil import BlobMedia 
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import date
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor
from anvil import *
import random

generated_ids = set()
def generate_unique_id():
    while True:
        new_id = str(random.randint(10000000, 99999999))
        if new_id not in generated_ids:
            generated_ids.add(new_id)
            return new_id

def get_temp_range(folder_path, set_point):
    if set_point == 20:
        return 15, 25
    elif set_point == 5:
        return 2, 8
    else:
        return -1, -1   
      
def update_heading_style(heading):
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)

def update_style(table):
    for cell in table.rows[0].cells:
    # Set background color to black
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), '000000')  # Black color
        cell._element.get_or_add_tcPr().append(shading_elm)
    # Set text color to white
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor(255, 255, 255)  # White text
def extract_temperatures_from_csv(file_path, start_datetime, end_datetime):
    temperatures = []
    # Read the CSV file into a DataFrame
    bytes_data = file_path.get_bytes()
    df = pd.read_csv(BytesIO(bytes_data)) 
    df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
    filtered_df = df[(df['DateTime'] >= start_datetime) & (df['DateTime'] <= end_datetime)]
    # Extract the temperature values from the 'Temperature' column if it exists
    if 'Temperature' in filtered_df.columns:
        temperatures = filtered_df['Temperature'].dropna().tolist()
    else:
        print(f"Column 'Temperature' not found in file: {file_path}")
    return temperatures
  
def calculate_diff(start_time, end_time):
    start_datetime = pd.to_datetime(start_time)
    end_datetime = pd.to_datetime(end_time)
    diff_time = end_datetime - start_datetime
    # Extract total days, hours, and minutes
    days = diff_time.days
    total_seconds = diff_time.seconds
    hours, remainder = divmod(total_seconds, 3600)
    minutes = remainder // 60
    seconds = minutes // 60
    # Format the output to include days, hours, and minutes
    if days > 0:
        formatted_diff = f"{days} day(s), {int(hours)} hour(s),  {int(minutes)} minute(s) and {int(seconds)} second(s)"
    else:
        formatted_diff = f"{int(hours)} hour(s) and {int(minutes)} minute(s) and {int(seconds)} second(s)"
    return formatted_diff

def process_file(doc, files, start_datetime, end_datetime):
    index=1
    total_avg_time=0
    table = doc.add_table(rows=1, cols=4)
    table.autofit = True
    table.allow_autofit = True
    table.style = 'Table Grid'
    overall_min, overall_max = 0, 0
    update_style(table)
    table.cell(0, 0).text = 'Measuring Point'
    table.cell(0, 1).text = 'Average [°C]'
    table.cell(0, 2).text = 'Minimum [°C]'
    table.cell(0, 3).text = 'Maximum [°C]'
    for filename in files:
        temperatures = extract_temperatures_from_csv(filename, start_datetime, end_datetime)
        max_temp = max(temperatures)
        min_temp = min(temperatures)
        if index == 1:
            overall_max = max_temp
            overall_min = min_temp
        else:
            overall_min = min(overall_min, min_temp)
            overall_max = max(overall_max, max_temp)
        avg_temp = round((max_temp+min_temp)/2, 2)
        total_avg_time += int(avg_temp)
        row_cells = table.add_row().cells
        row_cells[0].text = str(filename.name)
        row_cells[1].text = str(avg_temp) 
        row_cells[2].text = str(min_temp) 
        row_cells[3].text = str(max_temp)
        index+=1
    total_avg_time/=(index-1)
    return total_avg_time, overall_min, overall_max
  
def get_combined_graph(folder_path, start_datetime, end_datetime):
    # Step 1: Get all CSV files in the specified folder
    csv_files = folder_path
    num_files = len(csv_files)
    colors = plt.cm.viridis(np.linspace(0, 1, num_files))
    # Step 2: Create a new figure for plotting
    plt.figure(figsize=(12, 6))
    ax = plt.gca() 
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d/%Y\n%I:%M:%S %p'))
    # Initialize an empty DataFrame for combined data
    combined_data = pd.DataFrame()  
    # Step 3: Loop through each CSV file and filter data
     
    for idx, file in enumerate(csv_files):
          # Read the CSV file
          bytes_data = file.get_bytes()
          df = pd.read_csv(BytesIO(bytes_data)) 
          #df = pd.read_csv(BytesIO(bytes_data.decode('utf-8')))
          #df = pd.read_csv(StringIO(file.get_bytes().decode('utf-8')))
          #df = pd.read_csv(file)
          df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
          filtered_data = df[(df['DateTime'] >= start_datetime) & (df['DateTime'] <= end_datetime)]
          filtered_data['SmoothedTemperature'] = filtered_data['Temperature'].rolling(window=5).mean()
          # Check if filtered data is not empty before plotting
          if not filtered_data.empty:
              # Append the filtered data to the combined DataFrame
              combined_data = pd.concat([combined_data, filtered_data], ignore_index=True)
              # Plot the filtered data using the new DateTime column for x-axis
              plt.plot(filtered_data['DateTime'], 
                      filtered_data['SmoothedTemperature'], 
                      linestyle='-', color=colors[idx])
          else:
              print(f"No data in {file} for the given time period.")
    # Step 4: Add titles and labels
    plt.xlabel('Temperature (°C)')
    plt.ylabel('Date and Time')
    min_temp = filtered_data['Temperature'].min()
    max_temp = filtered_data['Temperature'].max()
    ax.set_yticks(np.arange(np.floor(min_temp), np.ceil(max_temp) + 1, 1))
    plt.xticks(rotation=45)  # Rotate x-axis labels for better visibility
    plt.legend()
    plt.grid()
    # Save and show the plot
    plt.tight_layout()
    plt.savefig("combined_data_plot.jpg")
    return "combined_data_plot.jpg", num_files

def low_high_calculator(folder_path, doc, temp, str_value, start_datetime, end_datetime):
    table = doc.add_table(rows=1, cols=3)
    update_style(table)
    table.style = 'Table Grid'
    table.autofit = True
    table.allow_autofit = True
    table.cell(0, 0).text = 'Measuring Point(s) ' + str_value + ' Value'
    table.cell(0, 1).text = 'Value[°C]'
    table.cell(0, 2).text = 'Date / Time'
    for filename in folder_path:
          bytes_data = filename.get_bytes()
          df = pd.read_csv(BytesIO(bytes_data)) 
          df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
          filtered_df = df[(df['DateTime'] >= start_datetime) & (df['DateTime'] >= end_datetime) & (df['Temperature'] == temp)]
          if not filtered_df.empty:
              row = filtered_df.iloc[0]
              row_cells = table.add_row().cells
              row_cells[0].text = filename.name
              row_cells[1].text = str(row['Temperature'])
              row_cells[2].text = str(row['Date'] + " " + row['Time'])
            
def extract_temp_from_csv(file_path):
    temperatures = []
    # Read the CSV file into a DataFrame
    bytes_data = file_path.get_bytes()
    df = pd.read_csv(BytesIO(bytes_data)) 
    temperatures = df['Temperature'].dropna().tolist()
    return temperatures

def process_all_files(folder_path, start_datetime, end_datetime):
    # Convert start and end date/time strings to datetime objects
    all_temperatures = []
    # Iterate over each CSV file in the given folder
    for filename in folder_path:
            temperatures = extract_temp_from_csv(filename)
            all_temperatures.extend(temperatures)  # Collect temperatures from each file
    # Calculate the maximum and minimum temperature across all files for the given range
    if all_temperatures:
        max_temp = max(all_temperatures)
        min_temp = min(all_temperatures)
        avg_temp = round(((max_temp+min_temp)/2), 2)
        return max_temp, min_temp, avg_temp
    else:
        print(f"No temperature data found between {start_datetime} and {end_datetime} across all files.")

def generate_graph_for_time_range(filtered_df, image_path):
    if not filtered_df.empty:
        plt.figure(figsize=(10, 6))
        ax = plt.gca() 
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d/%Y\n%I:%M:%S %p'))
        filtered_df.loc[:, 'DateTime'] = pd.to_datetime(filtered_df['Date'] + ' ' + filtered_df['Time'])
        filtered_df['SmoothedTemperature'] = filtered_df['Temperature'].rolling(window=5).mean()
    
        #plt.plot(filtered_df.loc[:, 'DateTime'], filtered_df['Temperature'], linestyle='-', color='b')
        plt.xlabel("Date and Time")
        plt.ylabel("Temperature (°C)")
        min_temp = filtered_df['Temperature'].min()
        max_temp = filtered_df['Temperature'].max()
        ax.set_yticks(np.arange(np.floor(min_temp), np.ceil(max_temp) + 1, 1))
        # Plot smoothed line, dropping NaN values that might arise from the rolling window
        plt.plot(filtered_df['DateTime'], filtered_df['SmoothedTemperature'], linestyle='-', color='black')
        #plt.grid(True)
        #plt.tight_layout(pad=1.5, w_pad=3.5, h_pad=1.0)
        img_stream = BytesIO()
        plt.savefig(img_stream, format='jpg')
        plt.close()
        img_stream.seek(0)  # Move to the start of the BytesIO object
        return img_stream
    else:
        print("No data available for the specified time range.")
        return ""
      
def mean_kinetic_temperature(temperatures, delta_h=83144, r=8.314):
    # Convert temperatures to Kelvin
    temps_in_kelvin = [temp + 273.15 for temp in temperatures]
    # Calculate the exponentials for each temperature
    exponentials = [math.exp(-delta_h / (r * temp)) for temp in temps_in_kelvin]
    # Calculate MKT using the formula
    mkt_kelvin = (-delta_h / r) / math.log(sum(exponentials) / len(exponentials))
    # Convert MKT back to °C
    mkt_celsius = mkt_kelvin - 273.15
    return round(mkt_celsius, 2)

def read_and_filter_data(doc, folder_path, start_datetime, end_datetime):
    counter=4.1
    value=0.1
    for filename in folder_path:
        if counter + value >= 5.0:
          counter = 4.1
          value = 0.1/10
          counter += value
        else:
          counter += value
        if value == 0.1:
          counter = round(counter, 1)
        elif value == 0.01:
          counter = round(counter, 2)
        heading = doc.add_heading(f'{str(counter)}: {filename.name}', level=2)
        update_heading_style(heading)
        bytes_data = filename.get_bytes()
        df = pd.read_csv(BytesIO(bytes_data)) 
        df['DateTime'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])
        filtered_df = df[(df['DateTime'] >= start_datetime) & (df['DateTime'] <= end_datetime)]
        file_name = filename.name 
        base = os.path.splitext(file_name)[0]
        new_filename = base + ".jpg"
        new_file_path = new_filename
        image_path = generate_graph_for_time_range(filtered_df, new_file_path)
        doc.add_picture(image_path, width=Inches(6.0)) 
        doc.add_heading('', level=2)
        table = doc.add_table(rows=8, cols=2)
        table.style = 'Table Grid'
        table.autofit = True
        table.allow_autofit = True
        update_style(table)
        table.cell(0, 0).text = 'Statistics'
        table.cell(0, 1).text = 'Value'
        table.cell(1, 0).text = 'Range'
        table.cell(1, 1).text = str(start_datetime) + '-' + str(end_datetime)
        temperatures = extract_temperatures_from_csv(filename, start_datetime, end_datetime)
        max_temp = max(temperatures)
        min_temp = min(temperatures)
        avg_temp = round((max_temp+min_temp)/2, 2)
        table.cell(2, 0).text = 'Average'
        table.cell(2, 1).text = str(avg_temp) + ' °C'
        table.cell(3, 0).text = 'Min/Max Thresholds'
        table.cell(3, 1).text = str(min_temp) + ' / ' + str(max_temp) +  ' °C'
        calculated_mkt = mean_kinetic_temperature(temperatures)
        table.cell(4, 0).text = 'MKT'
        table.cell(4, 1).text = str(calculated_mkt) +  ' °C'
        variance = round(statistics.variance(temperatures), 2)
        table.cell(5, 0).text = 'Variance'
        table.cell(5, 1).text = f"{variance:.2f}"
        std_deviation = statistics.variance(temperatures)
        table.cell(6, 0).text = 'Std. Deviation'
        table.cell(6, 1).text = f"{std_deviation:.4f}"
        table.cell(7, 0).text = 'Time: Within Min/Max'
        diff_time = calculate_diff(start_datetime, end_datetime)
        table.cell(7, 1).text = diff_time
    
def create_document(files, start_datetime, end_datetime, start_input, set_point, company_name, author_name, app_name):
    doc = Document()
    if start_input == 1:
        title = doc.add_heading('Field Shipment Test – 6-Hour Mapping-Empty Trailer', level=1)
    elif start_input == 2:
        title = doc.add_heading('Field Shipment Test – 6-Hour Mapping-Loaded Trailer', level=1)
    elif start_input == 3:
        title = doc.add_heading('Field Shipment Test – 1-Hour Temperature-Control Power Failure-Loaded Trailer', level=1)
    elif start_input == 4:
        title = doc.add_heading('Field Shipment Test – 1-Minute Open Door-Loaded Trailer', level=1)
    else:
        title = doc.add_heading('Field Shipment Test – 2-Hour Field Shipment Test-Loaded Trailer', level=1)
    update_heading_style(title)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    second_title = doc.add_heading(company_name + ' : Unit 324 Set Point ' +  str(set_point) + ' °C', level=2)
    second_title.italic = True
    second_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    update_heading_style(second_title)
    heading=doc.add_heading('1: General', level=2)
    update_heading_style(heading)
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    table.autofit = True
    table.allow_autofit = True
    table.cell(0, 0).text = 'Author'
    table.cell(0, 1).text = author_name
    table.cell(1, 0).text = 'Creation Date'
    table.cell(1, 1).text = str(date.today())
    table.cell(2, 0).text = 'Application'
    table.cell(2, 1).text = app_name
    table.cell(3, 0).text = 'Unique Mapping ID'
    table.cell(3, 1).text = generate_unique_id()
    table.cell(4, 0).text = 'Report Time Base'
    table.cell(4, 1).text = 'UTC-05:00'
    heading = doc.add_heading('2: Thresholds', level=2)
    update_heading_style(heading)
    table = doc.add_table(rows=3, cols=2)
    table.style = 'Table Grid'
    table.autofit = True
    table.allow_autofit = True
    update_style(table)
    table.cell(0, 0).text = 'Limit'
    table.cell(0, 1).text = 'Value'
    min_temp, max_temp = get_temp_range(files, set_point)
    table.cell(1, 0).text = 'Upper Temperature Limit'
    table.cell(1, 1).text = str(max_temp) + " °C"
    table.cell(2, 0).text = 'Lower Temperature Limit'
    table.cell(2, 1).text = str(min_temp) + " °C"
    heading = doc.add_heading('3: Conclusion', level=2)
    update_heading_style(heading)
    heading = doc.add_heading('3.1: Average/Minimum/Maximum Temperature', level=2)
    update_heading_style(heading)
    avg_temp, overall_min, overall_max = process_file(doc, files, start_datetime, end_datetime)
    heading = doc.add_heading('3.2: Overall Temperature', level=2)
    update_heading_style(heading)
    table = doc.add_table(rows=2, cols=2)
    update_style(table)
    table.style = 'Table Grid'
    table.autofit = True
    table.allow_autofit = True
    table.cell(0, 0).text = 'Measuring Point'
    table.cell(0, 1).text = 'Value[°C]'
    table.cell(1, 0).text = 'Overall average of all single average temperature values'
    avg_temp = round(avg_temp, 2)
    table.cell(1, 1).text = str(avg_temp)
    heading = doc.add_heading('3.3: Lowest/Highest Temperatures', level=2)
    update_heading_style(heading)
    low_high_calculator(files, doc, overall_min, 'Lowest', start_datetime, end_datetime)
    doc.add_heading('', level=2)
    low_high_calculator(files, doc, overall_max, 'Highest', start_datetime, end_datetime)
    heading = doc.add_heading('4: Mapping Results', level=2)
    update_heading_style(heading)
    combine_image_path, number_of_files = get_combined_graph(files, start_datetime, end_datetime)
    heading=doc.add_heading('4.1: Overview', level=2)
    update_heading_style(heading)
    doc.add_picture(combine_image_path, width=Inches(6.0)) 
    doc.add_heading('', level=2)
    table = doc.add_table(rows=6, cols=2)
    update_style(table)
    table.style = 'Table Grid'
    table.autofit = True
    table.allow_autofit = True
    table.cell(0, 0).text = 'Statistics'
    table.cell(0, 1).text = 'Value'
    table.cell(1, 0).text = 'Range'
    table.cell(1, 1).text = str(start_datetime) + '-' + str(end_datetime)
    total_max, total_min, total_avg = process_all_files(files, start_datetime, end_datetime)
    table.cell(2, 0).text = 'Average'
    table.cell(2, 1).text = str(total_avg) + ' °C'
    table.cell(3, 0).text = 'Min/Max Thresholds'
    table.cell(3, 1).text = str(total_max) + ' / ' + str(total_min) +  ' °C'
    table.cell(4, 0).text = 'Number of files'
    table.cell(4, 1).text = str(number_of_files)
    table.cell(5, 0).text = 'Number of sensors'
    table.cell(5, 1).text = str(number_of_files)
    read_and_filter_data(doc, files,start_datetime, end_datetime)
    doc_stream = io.BytesIO()
    doc.save(doc_stream) 
    doc_stream.seek(0)
    return BlobMedia(
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 
        doc_stream.read(), 
        'Temp Report.docx'
    ) 
@anvil.server.callable
def save_user_choice(start_digit, author_name, start_date, start_time, end_date, end_time, temp_value, application_name, company_name, files):
    start_date_str = start_date.strftime('%Y-%m-%d')
    start_str = start_date_str + " " + start_time
    start_datetime = pd.to_datetime(start_str)
    end_date_str = end_date.strftime('%Y-%m-%d')
    end_str = end_date_str + " " + end_time
    end_datetime = pd.to_datetime(end_str)
    return create_document(files, start_datetime, end_datetime, start_digit, temp_value, company_name, author_name, application_name)

