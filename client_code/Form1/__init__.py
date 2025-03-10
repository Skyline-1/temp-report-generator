from ._anvil_designer import Form1Template
from anvil import *
import re
from datetime import datetime
from anvil import BlobMedia
import anvil.server

class Form1(Form1Template):
    def __init__(self, **properties):
        # Set Form properties and Data Bindings.
        self.init_components(**properties)

    def submit_button_click(self, **properties):
        # Collect data from radio buttons
        error_messages = []
        all_valid = True
        # Check if trailer is selected or not
        selected_value=0
        print(self.radio_button_1.selected)
        print(self.radio_button_1.selected)
        print(self.radio_button_5.selected)
        print(self.radio_button_6.selected)
        print(self.radio_button_7.selected)
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
        print(selected_value)
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
        if all_valid:
          start_date_str = f"{start_date} {start_time}"  # Combine date and time
          end_date_str = f"{end_date} {end_time}"          # Combine date and time
      
          # Define the format you expect (adjust based on your date format)
          date_format = "%Y-%m-%d %H:%M:%S"
      
          # Convert strings to datetime objects
          start_datetime = datetime.strptime(start_date_str, date_format)
          end_datetime = datetime.strptime(end_date_str, date_format)
          if start_datetime > end_datetime:
            error_messages.append("Start date should be less than or equal to end date")
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
        operating_conditions=temp_value
        if not (self.radio_button_3.selected or self.radio_button_4.selected or self.radio_button_8.selected):
          error_messages.append("Temp value is required")
          all_valid = False
        season = self.text_box_7.text
        if not season:
          error_messages.append("Season is required")
          all_valid = False
        trailer_no = self.text_box_6.text
        if not trailer_no:
          error_messages.append("Trailer number is required")
          all_valid = False
        company_name = self.text_box_5.text
        if not company_name:
          error_messages.append("Company name is required")
          all_valid = False
        application_name = self.text_box_2.text
        if not application_name:
          error_messages.append("Application name is required")
          all_valid = False
        document_number = "CAL-F02"
        if not document_number:
          error_messages.append("Document number is required")
          all_valid = False  
        make = self.text_box_10.text
        if not make:
          error_messages.append("Make is required")
          all_valid = False 
        model_number = self.text_box_11.text
        if not model_number:
          error_messages.append("Model number is required")
          all_valid = False 
        vin_number = self.text_box_12.text
        if not vin_number:
          error_messages.append("Vin number is required")
          all_valid = False
        equipment_number = self.text_box_13.text
        if not equipment_number:
          error_messages.append("Equipment Id is required")
          all_valid = False
        creation_date = self.date_picker_3.date
        if not creation_date:
          error_messages.append("Creation date is required")
          all_valid = False
        equipment_type= "Refrigerator Trailer"
        # Call server function to save user choice
        date_of_revision = self.date_picker_4.date
        if not date_of_revision:
          error_messages.append("Revision date is required")
          all_valid = False
        if all_valid:
             print("Start Date=", start_date)
             print("End Date=", end_date)
             print("Start Time=", start_time)
             print("End Time=", end_time)
             blob_media = anvil.server.call('save_user_choice', selected_value, author_name, start_date, start_time, end_date, end_time, temp_value, application_name, company_name, files, trailer_no, season, protocol_number, document_number, make,model_number,vin_number, equipment_number, operating_conditions, creation_date, equipment_type, date_of_revision)
             self.download_link.url = blob_media.url  # Set the URL
             self.download_link.text = "Download your file"  # Link text
             self.download_link.visible = True  # Show the link
             anvil.media.download(blob_media)
        else:
            print(error_messages)
            print("Please check the errors and retry")

    


