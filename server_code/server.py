import anvil.server

@anvil.server.callable
def save_user_choice(selected_value, author_name, start_date, start_time, end_date, end_time, files, temp_value):
    print("Selected option: ", selected_value)
    print("Author name: ", author_name)
    print("Start date: ", start_date)
    print("Start time: ", start_time)
    print("End date: ", end_date)
    print("End time: ", end_time)
    print("Temp: ", temp_value)
    
    # Print file names (if any)
    if files:
        for file in files:
            print(f"Uploaded file: {file.name}")  # Print the name of each file
    else:
        print("No files uploaded.")
