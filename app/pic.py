from lib.gui_lib_wx import GuiLibWx, resource_path
from app.app import add_all_slides
from app.app import delete_all_slides
from app.app import add_pictures2
from app.app import add_picture_to_placeholder
import wx
import glob


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
        self.selected_radio_buttons = {
            "vertical_number": 0,
            "horizontal_number": 0,
            "picture_layout": 0,
            "picture_per_slide": 0,
            "picture_placeholder1": 0,
            "picture_placeholder2": 0,
            "picture_placeholder3": 0,
            "picture_placeholder4": 0,
        }
        self.text_data = {
            "slide_title": "",
            "slide_layout_number": 11,
            "pic_top1": 150,
            "pic_top2": 350,
            "pic_width": 300,
        }

        self.SetIcon(wx.Icon(resource_path("python.png")))
        self.taskbar = wx.adv.TaskBarIcon()
        self.taskbar.SetIcon(
            wx.Icon(resource_path("python.png"), wx.BITMAP_TYPE_PNG),
            "",
        )
        self.init()

    def init_ui(self):
        self.create_radio_button_layout(
            key="picture_layout",
            buttons=["placeholder", "custom"],
            label="Picture Layout",
        )

        self.create_text_control_layout(
            key="slide_layout_number",
            hint="Slide Layout (default "
            + str(self.text_data.get("slide_layout_number"))
            + ")",
        )

        if self.selected_radio_buttons["picture_layout"] == 0:
            self.create_radio_button_layout(
                key="picture_per_slide",
                buttons=["1", "2", "3", "4"],
                label="Picture Per Slide",
            )

            pic_per_slide = (
                self.selected_radio_buttons["picture_per_slide"] + 1
            )
            for i in range(1, pic_per_slide + 1):
                self.create_radio_button_layout(
                    key="picture_placeholder" + str(i),
                    buttons=["None", "1", "2", "3", "4", "5"],
                    label="Picture" + str(i) + " Title Placeholder",
                )
                self.create_text_ctrl_and_browse_layout(
                    key="folder_browse" + str(i),
                    func="self.folder_browse",
                    label="Pic" + str(i),
                )
                self.set_drop_target(
                    key="folder_browse" + str(i), mode="folder"
                )

            self.create_button_layout(
                key="execute",
                label="Add Picture",
                func="self.add_picture_to_placeholder",
            )
        else:
            self.create_text_control_layout(
                key="pic_top1",
                hint="Picture Top1 (default "
                + str(self.text_data.get("pic_top1"))
                + ")",
            )
            self.create_text_control_layout(
                key="pic_top2",
                hint="Picture Top2 (default "
                + str(self.text_data.get("pic_top2"))
                + ")",
            )
            self.create_text_control_layout(
                key="pic_width",
                hint="Picture Width (default "
                + str(self.text_data.get("pic_width"))
                + ")",
            )
            self.create_radio_button_layout(
                key="vertical_number", buttons=["1", "2"], label="vertical"
            )
            self.create_radio_button_layout(
                key="horizontal_number",
                buttons=["1", "2", "3", "4"],
                label="horizontal",
            )

            vn = self.selected_radio_buttons["vertical_number"] + 1
            hn = self.selected_radio_buttons["horizontal_number"] + 1

            for i in range(1, hn * vn + 1):
                self.create_text_ctrl_and_browse_layout(
                    key="folder_browse" + str(i),
                    func="self.folder_browse",
                    label="Pic" + str(i),
                )
                self.set_drop_target(
                    key="folder_browse" + str(i), mode="folder"
                )

            self.create_button_layout(
                key="execute",
                label="Add Picture",
                func="self.add_picture",
            )
        self.create_button_layout(
            key="add_slide",
            label="Add All Slides",
            func="self.add_all_slides",
        )
        self.create_button_layout(
            key="delete_slide",
            label="Delete All Slides",
            func="self.delete_all_slides",
        )

        self.setup_layout()

    def add_picture(self, event=None):
        vn = self.selected_radio_buttons["vertical_number"] + 1
        hn = self.selected_radio_buttons["horizontal_number"] + 1
        pic_top1 = int(self.text_data["pic_top1"])
        pic_top2 = int(self.text_data["pic_top2"])
        pic_width = int(self.text_data["pic_width"])
        slide_layout_number = int(self.text_data["slide_layout_number"])
        file_paths = []
        for i in range(1, hn * vn + 1):
            file_paths.append(
                glob.glob(
                    self.text_data["folder_browse" + str(i)] + "/**/*.png",
                    recursive=True,
                )
            )

        add_pictures2(
            file_paths,
            slide_layout=slide_layout_number,
            horizontal=hn,
            vertical=vn,
            pic_top1=pic_top1,
            pic_top2=pic_top2,
            pic_width=pic_width,
        )

    def add_picture_to_placeholder(self, event=None):
        pic_per_slide = self.selected_radio_buttons["picture_per_slide"] + 1
        slide_layout_number = int(self.text_data["slide_layout_number"])
        file_paths = []
        title_placeholders = []
        for i in range(1, pic_per_slide + 1):
            file_paths.append(
                glob.glob(
                    self.text_data["folder_browse" + str(i)] + "/**/*.png",
                    recursive=True,
                )
            )
            title_placeholders.append(
                self.selected_radio_buttons.get("picture_placeholder" + str(i))
            )

        if 0 in title_placeholders:
            title_placeholders = None

        add_picture_to_placeholder(
            file_paths=file_paths,
            slide_layout=slide_layout_number,
            title_placeholder_numbers=title_placeholders,
        )

    def add_all_slides(self, event=None):
        add_all_slides(100, placeholder_number=True)

    def delete_all_slides(self, event=None):
        delete_all_slides()


def pic():
    app = wx.App()
    MyGui(
        None,
        title="Pic",
        height=300,
        width=500,
        pos=(100, 100),
        layout="dynamic",
    )
    app.MainLoop()
