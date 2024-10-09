import anvil.server

@anvil.server.callable
def save_user_choice(author_name, start_date, start_time, end_date, end_time, files):
    #print("Selected option: ", selected_option)
    print("Author name: ", author_name)
    print("Start date: ", start_date)
    print("Start time: ", start_time)
    print("End date: ", end_date)
    print("End time: ", end_time)
    
    # Print file names (if any)
    if files:
        for file in files:
            print(f"Uploaded file: {file.name}")  # Print the name of each file
    else:
        print("No files uploaded.")
