import os
import csv

import pywintypes

from lib.powerpoint import PowerPoint
from lib.powerpoint import get_file_name
from lib.powerpoint import get_dir_name


def add_picture_to_placeholder(
    file_paths,
    slide_layout,
    title_placeholder_numbers=None,
    titles=None,
    file_name_to_title=True,
):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for i in range(len(file_paths[0])):
        ppt.add_slide(slide_layout=slide_layout)
        for j in range(len(file_paths)):
            if title_placeholder_numbers is not None:
                if titles is not None:
                    title = titles[j]
                else:
                    if file_name_to_title:
                        title = get_file_name(file_path=file_paths[j][i])
                    else:
                        title = get_dir_name(file_path=file_paths[j][i])
                ppt.add_text_to_placeholder(
                    text=title,
                    placeholder_number=title_placeholder_numbers[j],
                )
            ppt.add_picture(
                file_path=file_paths[j][i], slide_number=ppt.slide_count()
            )


def add_picture_to_custom_layout(
    file_paths,
    slide_layout=11,
    picture_width=300,
    picture_top1=100,
    picture_top2=300,
    vertical=2,
    horizontal=1,
    font_size=14,
    file_name_to_title=True,
):
    ppt = PowerPoint()
    ppt.setup_active_presentation()

    left_list = []
    top_list = []
    pic_left1 = ppt.slide_width * 0.125 - picture_width * 0.5
    pic_left2 = ppt.slide_width * 0.166 - picture_width * 0.5
    pic_left3 = ppt.slide_width * 0.25 - picture_width * 0.5
    pic_left4 = ppt.slide_width * 0.375 - picture_width * 0.5
    pic_left5 = ppt.slide_width * 0.5 - picture_width * 0.5
    pic_left6 = ppt.slide_width * 0.625 - picture_width * 0.5
    pic_left7 = ppt.slide_width * 0.75 - picture_width * 0.5
    pic_left8 = ppt.slide_width * 0.834 - picture_width * 0.5
    pic_left9 = ppt.slide_width * 0.875 - picture_width * 0.5

    if vertical == 1 and horizontal == 1:
        left_list = [pic_left5]
        top_list.append(picture_top1)

    if vertical == 1 and horizontal == 2:
        left_list = [pic_left3, pic_left7]
        top_list = [picture_top1, picture_top1]

    if vertical == 1 and horizontal == 3:
        left_list = [pic_left2, pic_left5, pic_left8]
        top_list = [picture_top1, picture_top1, picture_top1]

    if vertical == 1 and horizontal == 4:
        left_list = [
            pic_left1,
            pic_left4,
            pic_left6,
            pic_left9,
        ]
        top_list = [picture_top1, picture_top1, picture_top1, picture_top1]

    if vertical == 2 and horizontal == 1:
        left_list = [pic_left5, pic_left5]
        top_list = [picture_top1, picture_top2]

    if vertical == 2 and horizontal == 2:
        left_list = [
            pic_left3,
            pic_left7,
            pic_left3,
            pic_left7,
        ]
        top_list = [picture_top1, picture_top1, picture_top2, picture_top2]

    if vertical == 2 and horizontal == 3:
        left_list = [
            pic_left2,
            pic_left5,
            pic_left8,
            pic_left2,
            pic_left5,
            pic_left8,
        ]
        top_list = [
            picture_top1,
            picture_top1,
            picture_top1,
            picture_top2,
            picture_top2,
            picture_top2,
        ]

    if vertical == 2 and horizontal == 4:
        left_list = [
            pic_left1,
            pic_left4,
            pic_left6,
            pic_left9,
            pic_left1,
            pic_left4,
            pic_left6,
            pic_left9,
        ]
        top_list = [
            picture_top1,
            picture_top1,
            picture_top1,
            picture_top1,
            picture_top2,
            picture_top2,
            picture_top2,
            picture_top2,
        ]

    for i in range(len(file_paths[0])):
        ppt.add_slide(slide_layout=slide_layout)
        for j in range(len(file_paths)):
            if file_name_to_title:
                file_name = get_file_name(file_paths[j][i])
            else:
                file_name = get_dir_name(file_paths[j][i])
            ppt.add_textbox(
                text=file_name,
                top=top_list[j],
                left=left_list[j],
                width=picture_width,
                height=10,
                font_size=font_size,
            )
            ppt.add_picture(
                file_path=file_paths[j][i],
                slide_number=ppt.slide_count(),
                top=top_list[j],
                left=left_list[j],
                width=picture_width,
            )


def add_table(data, cell_width=None):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    ppt.add_slide(slide_layout=11)
    ppt.add_table(data, cell_width)


def add_all_slides(count, placeholder_number=False):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for i in range(1, count):
        ppt.add_slide(i)
        if placeholder_number:
            for j in range(1, ppt.placeholder_count(ppt.slide_count()) + 1):
                try:
                    ppt.add_text_to_placeholder(text=j, placeholder_number=j)
                except pywintypes.com_error:
                    pass


def delete_all_slides():
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for i in range(ppt.slide_count(), 0, -1):
        ppt.delete_slide(i)


def add_picture_to_active_slide(file_path):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    ppt.add_picture(file_path, slide_number=ppt.active_slide_number())


def replace_text(slide_number, file):
    datas = []
    with open(file, "r") as f:
        reader = csv.reader(f)
        for line in reader:
            datas.append(line)

    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for placeholder_number in range(
        1, ppt.placeholder_count(slide_number) + 1
    ):
        for data in datas:
            ppt.replace_placeholder_text(
                slide_number=slide_number,
                placeholder_number=placeholder_number,
                before=data[0],
                after=data[1],
            )
