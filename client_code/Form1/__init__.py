from ._anvil_designer import Form1Template
from anvil import *
import re
import anvil.server

class Form1(Form1Template):
    def __init__(self, **properties):
        # Set Form properties and Data Bindings.
        self.init_components(**properties)

    def submit_button_click(self, **properties):
        # Collect data from radio buttons
        print("ON submit button click method")
        error_messages = []
        all_valid = True
        # Check if trailer is selected or not
        if self.radio_button_1.selected:
          selected_value = 1
        elif self.radio_button_2.selected:
          selected_value = 2
        elif self.radio_button_5.selected:
          selected_value = 3
        elif self.radio_button_6.selected:
          selected_value = 4
        elif self.radio_button_7.selected:
          selected_value = 5
        if not (self.radio_button_1.selected or self.radio_button_2.selected or self.radio_button_5.selected or self.radio_button_6.selected or self.radio_button_7.selected):
          error_messages.append("Select one option from trailer group")
          all_valid = False
         #Check if Author's name is added or not 
        author_name = self.text_box_1.text
        if not author_name:
          error_messages.append("Author name is required")
          all_valid = False
        #check if start_date is picked or not  
        start_date = self.date_picker_1.date
        if not start_date:
          error_messages.append("Start date is required")
          all_valid = False
        #check if start_time is picked or not and is in 24 hour format or not
        start_time = self.text_box_3.text
        time_pattern = r"^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$|^([01]\d|2[0-3]):([0-5]\d)$|^([01]\d|2[0-3])$"
        if not re.match(time_pattern, start_time):
            error_messages.append("Invalid time format! Please enter a valid time in HH:MM:SS, HH:MM, or HH in 24 hour format.")
            all_valid = False
        #check if end_date is picked or not  
        end_date = self.date_picker_2.date
        if not end_date:
          error_messages.append("End date is required")
          all_valid = False
        end_time = self.text_box_4.text
      #check if end_time is picked or not and is in 24 hour format or not
        if not re.match(time_pattern, end_time):
            error_messages.append("Invalid time format! Please enter a valid time in HH:MM:SS, HH:MM, or HH in 24 hour format.")
            all_valid = False
        #check if user has added files or not and whether they have extension as csv or not
        files = self.file_loader_1.files  # Get the list of files
        if len(self.file_loader_1.files) <= 0:
          error_messages.append("Atleast add one more csv file")
          all_valid = False
        else:
          # Check if all added files are CSV
          for file in files:
              if not file.name.lower().endswith('.csv'):
                 error_messages.append(f"{file.name} is not a CSV file.")
                 all_valid = False
        #check if temperature is added or not
        if self.radio_button_3.selected:
          temp_value = 5
        elif self.radio_button_4.selected:
          temp_value = 20
        elif self.radio_button_8.selected:
          temp_value = -20
        if not (self.radio_button_3.selected or self.radio_button_4.selected or self.radio_button_8.selected):
          error_messages.append("Temp value is required")
          all_valid = False
        company_name = self.text_box_5.text
        if not company_name:
          error_messages.append("Company name is required")
          all_valid = False
        application_name = self.text_box_2.text
        if not application_name:
          error_messages.append("Application name is required")
          all_valid = False
        # Call server function to save user choice
        if all_valid:
            #anvil.server.call('say_hello', 'Parneet')
            anvil.server.call('save_user_choice', selected_value, author_name, start_date, start_time, end_date, end_time, temp_value, application_name, company_name)
        else:
            print(error_messages)
            print("Please check the errors and retry")

    def text_box_5_pressed_enter(self, **event_args):
      """This method is called when the user presses Enter in this text box"""
      pass
    


