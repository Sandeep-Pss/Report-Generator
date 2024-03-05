from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.filechooser import FileChooserIconView
from report import json_to_excel

class JSONToExcelConverter(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'

        self.label = Label(text="Select JSON file:")
        self.add_widget(self.label)

        self.file_chooser = FileChooserIconView()
        self.add_widget(self.file_chooser)

        self.select_button = Button(text="Select JSON file")
        self.select_button.bind(on_press=self.select_json_file)
        self.add_widget(self.select_button)

        self.convert_button = Button(text="Convert to Excel")
        self.convert_button.bind(on_press=self.convert_to_excel)
        self.add_widget(self.convert_button)

    def select_json_file(self, instance):
        self.file_chooser.path = '/'
        self.file_chooser.filters = ['*.json']  # Show only JSON files

    def convert_to_excel(self, instance):
        selected_file_path = self.file_chooser.selection and self.file_chooser.selection[0]
        if selected_file_path:
            json_to_excel(selected_file_path)
            self.label.text = "Conversion completed successfully."
        else:
            self.label.text = "No file selected."

