import logging
import os
from datetime import date
from openpyxl import load_workbook
import api
from api import EditAPI
from api import IndividualEdit
 
from kivy.properties import StringProperty
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.app import App
 
 
KV = """
<Viewclass@BoxLayout>:
    padding: 20
    l1: ""
    l2: ""
    l3: ""
    km_input1: ""
    area_input1: ""
    km_input2: ""
    area_input2: ""
    vmf_input: ""
    vmf_km_input: ""
    vmf_area_input: ""
    index: 0
    orientation: "vertical"
    Label:
        text: root.l1
        font_size : 20
        color : (0,0,0,1)
        bold: True
    BoxLayout:
        InfoLabel:
            text: "PFM-1"
        NameLabel:
            text: root.l2
        AreaInput:
            text: root.km_input1
            on_text:
                app.root.ie.data[root.index]["km_input1"] = self.text
        AreaInput:
            text: root.area_input1
            on_text:
                app.root.ie.data[root.index]["area_input1"] = self.text
    BoxLayout:
        InfoLabel:
            text: "PFM-2"
        NameLabel:
            text: root.l3
        AreaInput:
            text: "KM"
            on_text:
                app.root.ie.data[root.index]["km_input2"] = self.text
        AreaInput:
            text: "Areas Covered"
            on_text:
                app.root.ie.data[root.index]["area_input2"] = self.text
    BoxLayout:
        InfoLabel:
            text: "VMF"
        AreaInput:
            text: "Name"
            width: 5
            on_text:
                app.root.ie.data[root.index]["vmf_input"] = self.text
        AreaInput:
            text: "KM"
            on_text:
                app.root.ie.data[root.index]["vmf_km_input"] = self.text
        AreaInput:
            text: "Areas Covered"
            on_text:
                app.root.ie.data[root.index]["vmf_area_input"] = self.text
    

<NameLabel@Label>:
    color : (0,0,0,1)
    text_size: root.width, None
    size: self.texture_size
 
<InfoLabel@Label>:
    size_hint_x: 0.5
    bold: True
    color: 0,0,0,1
 
<AreaInput@TextInput>:
    multiline: False
 
<CircleButton@Button>:
    on_release:
        app.circle = self.text
        self.parent.sm.current = "Next Page"
    size_hint: .5, .5
    font_size: "20sp"
    background_color: 177/255, 126/255, 5/255, 1
 
<SelectCircles@GridLayout>:
    cols: 1
    spacing: 10
    padding: 40
    Label:
        color:(0,0,0,1)
        font_size: "20sp"
        bold:True
        text: "TAP BELOW TO EDIT YOUR CIRCLE"
    CircleButton:
        text: "Circle 6"
    CircleButton:
        text: "Circle 7"
    CircleButton:
        text: "Circle 8"
    CircleButton:
        text: "Circle 9"
    CircleButton:
        text: "Circle 10"
    CircleButton:
        text: "Circle 11"
 
<BackUpdateButton@Button>:
    font_size: "20sp"
 
ScreenManager:
    ie: ie
    Screen:
        name: "Main Page"
        SelectCircles:
            sm: root
    Screen:
        id: nextpage
        on_pre_enter:
            ie.init_data()
        name: "Next Page"
        BoxLayout:
            orientation: "vertical"
            BoxLayout:
                size_hint_y: .2
                BackUpdateButton:
                    text: "Go Back"
                    on_release:
                        root.current = "Main Page"
                BackUpdateButton:
                    text: "Update All"
                    on_release:
                        print(ie.data)
                        ie.update_all()
            IndividualEdit:
                id: ie
                sm: root
                viewclass: 'Viewclass'
                RecycleBoxLayout:
                    default_size: None, 350
                    default_size_hint: 1, None
                    size_hint_y: None
                    height: self.minimum_height
                    orientation: 'vertical'
"""
 
Window.clearcolor = (135/255,206/255,235/255,1)
 
class ReportGenerator(App):
    circle = StringProperty("")
    sheet_editors = { f"Circle {i}":EditAPI(f"4-08-2020 (2).xlsx", f"Circle {i}") for i in range(6,12) }
 
    def build(self):
        return Builder.load_string(KV)
 
 
if __name__ == "__main__":
    Application = ReportGenerator
    Application().run()