import logging
import os
from datetime import date
from openpyxl import load_workbook
import api
from api import EditAPI
from api import SelectCircles
from api import IndividualEdit
from api import InfoLabel, AreaInput, CircleButton
from kivy.app import App
from kivy.core.window import Window
from kivy.lang import Builder

KV = """
<InfoLabel>:
    size_hint_x: None
    width: 20
    bold: True
    color: 0,0,0,1

<AreaInput>:
    multiline: False
    width: 20

<CircleButton>:
    on_release:
        self.parent.clicked(self)
    size_hint: .5, .5
    font_size: "20sp"
    background_color: 177/255, 126/255, 5/255, 1


ScreenManager:
    Screen:
        name: "Main Page"
        SelectCircles:
            sm: root
    Screen:
        on_pre_enter:
            ie.on_pre_enter()
        name: "Next Page"
        BoxLayout:
            orientation: "vertical"
            BoxLayout:
                size_hint_y: .2
                Button:
                    text: "Go Back"
                    on_release:
                        ie.go_back()
                Button:
                    text: "Update All"
                    on_release:
                        ie.update_all()
            ScrollView:
                size_hint: 1, 0.8
                IndividualEdit:
                    id: ie
                    size_hint_y: None
                    height: self.minimum_height
                    sm: root
"""

Window.clearcolor = (135/255,206/255,235/255,1)

class ReportGenerator(App):

    def build(self):
        return Builder.load_string(KV)


if __name__ == "__main__":
    Application = ReportGenerator
    Application().run()
