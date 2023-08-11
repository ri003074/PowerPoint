import os
from lib.powerpoint import PowerPoint


def add_picture(file_paths, slide_layout, top=0, left=0, font_size=14):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for file_path in file_paths:
        ppt.add_slide(slide_layout=slide_layout)
        ppt.add_picture(file_path=file_path, top=top, left=left)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        ppt.add_text_to_placeholder(
            text=file_name, placeholder_number=1, font_size=font_size
        )


def add_pictures(
    left_file_paths,
    right_file_paths,
    picture_top,
    picture_width,
    slide_layout,
    font_size=14,
):
    ppt = PowerPoint()
    ppt.setup_active_presentation()

    left_picture_left = ppt.slide_width * 0.25 - picture_width * 0.5
    right_picture_left = ppt.slide_width * 0.75 - picture_width * 0.5

    for left_file_path, right_file_path in zip(
        left_file_paths, right_file_paths
    ):
        ppt.add_slide(slide_layout=slide_layout)
        file_name = os.path.splitext(os.path.basename(left_file_path))[0]
        ppt.add_textbox(
            text=file_name,
            top=picture_top,
            left=0,
            width=ppt.slide_width / 2,
            height=10,
            font_size=font_size,
        )
        ppt.add_picture(
            file_path=left_file_path,
            top=picture_top,
            left=left_picture_left,
            width=picture_width,
        )
        file_name = os.path.splitext(os.path.basename(right_file_path))[0]
        ppt.add_textbox(
            text=file_name,
            top=picture_top,
            left=ppt.slide_width / 2,
            width=ppt.slide_width / 2,
            height=10,
            font_size=font_size,
        )
        ppt.add_picture(
            file_path=right_file_path,
            top=picture_top,
            left=right_picture_left,
            width=picture_width,
        )


def add_all_slides(count):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for i in range(1, count):
        ppt.add_slide(i)


def delete_all_slides():
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    print(ppt.slide_count())
    for i in range(ppt.slide_count(), 0, -1):
        ppt.delete_slide(i)
