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
        # Collect text input
        if self.radio_button_1.selected:
          selected_value = 'Empty' 
        elif self.radio_button_2.selected:
          selected_value = 'Loaded'
        if not (self.radio_button_1.selected or self.radio_button_2.selected):
          error_messages.append("Select one option, loaded or empty")
          all_valid = False
        author_name = self.text_box_1.text
        if not author_name:
          error_messages.append("Author name is required")
          all_valid = False
        start_date = self.date_picker_1.date
        if not start_date:
          error_messages.append("Start date is required")
          all_valid = False
        start_time = self.text_box_3.text
        time_pattern = r"^([01]\d|2[0-3]):([0-5]\d):([0-5]\d)$|^([01]\d|2[0-3]):([0-5]\d)$|^([01]\d|2[0-3])$"
        if not re.match(time_pattern, start_time):
            error_messages.append("Invalid time format! Please enter a valid time in HH:MM:SS, HH:MM, or HH in 24 hour format.")
            all_valid = False
        end_date = self.date_picker_2.date
        if not end_date:
          error_messages.append("End date is required")
          all_valid = False
        end_time = self.text_box_4.text
        if not re.match(time_pattern, end_time):
            error_messages.append("Invalid time format! Please enter a valid time in HH:MM:SS, HH:MM, or HH in 24 hour format.")
            all_valid = False
        # Collect data from the second set of radio buttons
        # Collect uploaded files
        files = self.file_loader_1.files  # Get the list of files
        if self.file_loader_1.files.count <= 0:
          error_messages.append("Atleast add one more csv file")
          all_valid = False
        if self.radio_button_3.selected:
          temp_value = 5
        elif self.radio_button_4.selected:
          temp_value = 20
        if not (self.radio_button_3.selected or self.radio_button_3.selected):
          error_messages.append("Temp value is required")
          all_valid = False
        # Call server function to save user choice
        if all_valid:
          anvil.server.call('save_user_choice', selected_value, author_name, start_date, start_time, end_date, end_time, files, temp_value)
        else:
          print(error_messages)
    def save_user_choice(self, **event_args):
      """This method is called when the button is clicked"""
      pass


