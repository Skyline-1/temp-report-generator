from ._anvil_designer import Form1Template
from anvil import *
import anvil.server

class Form1(Form1Template):
    def __init__(self, **properties):
        # Set Form properties and Data Bindings.
        self.init_components(**properties)

    def submit_button_click(self, **properties):
        # Collect data from radio buttons
        print("ON submit button click method")
        # Collect text input
        author_name = self.text_box_1.text
        start_date = self.date_picker_1.date
        start_time = self.text_box_3.text
        end_date = self.date_picker_2.date
        end_time = self.text_box_4.text

        # Collect data from the second set of radio buttons
        # Collect uploaded files
        files = self.file_loader_1.files  # Get the list of files

        # Call server function to save user choice
        anvil.server.call('save_user_choice', author_name, start_date, start_time, end_date, end_time, files)

    def save_user_choice(self, **event_args):
      """This method is called when the button is clicked"""
      pass


