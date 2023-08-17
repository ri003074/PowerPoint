import datetime
import os
from functools import partial

import wx
from PIL import ImageGrab
import sys

from app.app import add_all_slides
from app.app import add_picture_to_placeholder
from lib.get_window_info import (
    get_window_titles,
    get_window_rect_adj,
    get_monitor_full_size_info,
    set_foreground_window,
)
from lib.gui_lib_wx import GuiLibWx, resource_path
from lib.ss import get_screen_shot

sys.path.append("'C:\\Users\\ri003\\Documents\\Programming\\GUI")
ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)


class MyGui(GuiLibWx):
    def __init__(self, parent, title, height, width, pos, layout):
        super(MyGui, self).__init__(
            parent,
            title=title,
            height=height,
            width=width,
            pos=pos,
            layout=layout,
        )
        self.selected_radio_buttons = {"app_name": 0}
        self.text_data = {"slide_title": "", "slide_layout_number": 1}
        self.window_titles = get_window_titles()

        self.SetIcon(wx.Icon(resource_path("python.png")))
        self.taskbar = wx.adv.TaskBarIcon()
        self.taskbar.SetIcon(
            wx.Icon(resource_path("python.png"), wx.BITMAP_TYPE_PNG),
            "",
        )
        self.init()

    def init_ui(self):
        self.create_radio_button_layout(
            key="app_name",
            buttons=self.window_titles,
            label="Application List",
            alignment="vertical",
            height_coefficient=1.0,
        )
        monitor_size, x_left, x_right = get_monitor_full_size_info()
        self.create_text_layout(
            key="monitor left",
            label=" monitor left, right : "
            + str(x_left)
            + ", "
            + str(x_right),
        )
        self.create_choice_layout(
            key="slide_layout_number",
            choices=["select slide layout"] + [str(i) for i in range(1, 50)],
            default=int(self.text_data.get("slide_layout_number")),
        )
        self.create_text_control_layout(key="slide_title", hint="slide title")

        self.create_button_layout(
            key="execute",
            label="get screen shot",
            func="self.execute",
        )
        self.create_button_layout(
            key="update_window",
            label="update window",
            func="self.update_window",
        )
        self.create_button_layout(
            key="add_all_slides",
            label="add all slides",
            func="self.add_all_slides",
        )

        self.setup_layout()

    def update_window(self, event=None):
        self.window_titles = get_window_titles()
        self.init()

    def add_all_slides(self, event=None):
        add_all_slides(30)

    def execute(self, event=None):
        app_name = self.window_titles[
            self.selected_radio_buttons.get("app_name")
        ]
        title = self.text_data.get("slide_title")
        slide_layout_number = int(self.text_data.get("slide_layout_number"))
        get_screen_shot_and_insert_to_pptx(
            app_title=app_name,
            slide_title=title,
            slide_layout_number=slide_layout_number,
        )


def get_screen_shot_and_insert_to_pptx(
    app_title, slide_title="", slide_layout_number=1
):
    set_foreground_window(window_title=app_title)
    rect = get_window_rect_adj(window_title=app_title)
    monitor_full_size, x_left, x_right = get_monitor_full_size_info()
    file_name = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "_tmp.png"
    file_path = os.getcwd() + "/" + file_name
    get_screen_shot(
        save_file_path=file_path,
        x_top_left=monitor_full_size.get(rect[0]),
        y_top_left=rect[1],
        x_bottom_right=monitor_full_size.get(rect[2]),
        y_bottom_right=rect[3],
    )
    if slide_title != "":
        text = [slide_title]
    else:
        text = None

    add_picture_to_placeholder(
        [file_path],
        slide_layout=slide_layout_number,
        title_placeholder_numbers=[1],
        titles=text,
    )


def ss_2_ppt():
    app = wx.App()
    MyGui(
        None,
        title="Screen Shot",
        height=300,
        width=500,
        pos=(100, 100),
        layout="static",
    )
    app.MainLoop()


if __name__ == "__main__":
    ss_2_ppt()
