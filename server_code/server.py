import anvil.server
import pandas as pd
@anvil.server.callable
def save_user_choice(selected_value, author_name, start_date, start_time, end_date, end_time, files, temp_value):
    print("Selected option: ", selected_value)
    print("Author name: ", author_name)
    print("start date: ", start_date)
    print("start time: ", start_time)
    start_date_str = start_date.strftime('%Y-%m-%d')
    start_str = start_date_str + " " + start_time
    start_datetime = pd.to_datetime(start_str)
    print("Start date time: ", start_datetime)
    end_date_str = end_date.strftime('%Y-%m-%d')
    end_str = end_date_str + " " + end_time
    end_datetime = pd.to_datetime(end_str)
    print("End date time: ", end_datetime)
    print("Temp: ", temp_value)
    
    # Print file names (if any)
    if files:
        for file in files:
            print(f"Uploaded file: {file.name}")  # Print the name of each file
    else:
        print("No files uploaded.")
