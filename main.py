from kivy.app import App
from json_to_excel_converter import JSONToExcelConverter

class MyApp(App):
    def build(self):
        return JSONToExcelConverter()

if __name__ == "__main__":
    MyApp().run()
