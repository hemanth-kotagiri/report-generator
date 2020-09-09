from kivy.app import App
from kivy.uix.button import Button


class TestApp(App):
    "Test Application"

    def build(self):
        b1 = Button(text = "Test Button")
        return b1


TestApp().run()
