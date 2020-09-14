import logging
import os
from datetime import date
from openpyxl import load_workbook
import api
from api import EditAPI
from api import SelectCircles
from api import IndividualEdit
from kivy.app import App
from kivy.uix.screenmanager import Screen, ScreenManager, FadeTransition
from kivy.core.window import Window
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button


#TODO: Add this inside the API
# Loading the workbook from the current directory
workbook = load_workbook(os.path.join(os.getcwd(), "4-08-2020 (2).xlsx"))

Window.clearcolor = (135/255,206/255,235/255,1)
class MyScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        root = BoxLayout(orientation="vertical")
 
        self.individualedit = IndividualEdit(cols=1, spacing=10, size_hint_y=None)
 
        Back_And_Update_Buttons = BoxLayout(size_hint_y=0.2)
        back_button = Button(text="Go Back")
        back_button.bind(on_press=self.individualedit.go_back)
        Back_And_Update_Buttons.add_widget(back_button)

        update_date_button = Button(text = "Update All")
        update_date_button.bind(on_press = self.individualedit.update_all)
        Back_And_Update_Buttons.add_widget(update_date_button)
        
        root.add_widget(Back_And_Update_Buttons)
        
        # Make sure the height is such that there is something to scroll.
        self.individualedit.bind(minimum_height=self.individualedit.setter('height'))
        scroll = ScrollView(size_hint=(1, 0.8))
        scroll.add_widget(self.individualedit)
        
        root.add_widget(scroll)
        self.add_widget(root)
    

    def on_pre_enter(self, *args):
        self.individualedit.on_pre_enter()


class ReportGenerator(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)

    def build(self):

        # Initializing a screen manager
        self.screen_manager = ScreenManager()
        api.root = self.screen_manager
        
        # Adding select circles(Main page) to the screen_manager
        self.select_circles_page = SelectCircles()
        screen = Screen(name="Main Page")
        screen.add_widget(self.select_circles_page)
        self.screen_manager.add_widget(screen)

        # Adding Individual editing Screen to the screen_manager
        screen = MyScreen(name="Next Page")
        self.screen_manager.add_widget(screen)

        return self.screen_manager


if __name__ == "__main__":
    Application = ReportGenerator()
    Application.run()
