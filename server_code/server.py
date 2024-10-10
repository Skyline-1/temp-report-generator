import anvil.server

@anvil.server.callable
def save_user_choice(selected_value, author_name, start_datetime, end_datetime, files, temp_value):
    print("Selected option: ", selected_value)
    print("Author name: ", author_name)
    print("Start date time: ", start_datetime)
    print("End date time: ", end_datetime)
    print("Temp: ", temp_value)
    
    # Print file names (if any)
    if files:
        for file in files:
            print(f"Uploaded file: {file.name}")  # Print the name of each file
    else:
        print("No files uploaded.")
