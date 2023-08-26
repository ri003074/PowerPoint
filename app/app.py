import os
import csv

import pywintypes

from lib.powerpoint import PowerPoint
from lib.powerpoint import get_file_name
from lib.powerpoint import get_dir_name


def add_picture(file_paths, slide_layout, top=0, left=0, font_size=14):
    ppt = PowerPoint()
    ppt.setup_active_presentation()
    for file_path in file_paths:
        ppt.add_slide(slide_layout=slide_layout)
        ppt.add_picture(
            file_path=file_path,
            slide_number=ppt.slide_count(),
            top=top,
            left=left,
        )
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        ppt.add_text_to_placeholder(
            text=file_name, placeholder_number=1, font_size=font_size
        )


def add_pictures(
    left_file_paths,
    right_file_paths,
    picture_top=0,
    picture_width=0,
    slide_layout=1,
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
        file_name = get_file_name(left_file_path)
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
            slide_number=ppt.slide_count(),
            top=picture_top,
            left=left_picture_left,
            width=picture_width,
        )
        file_name = get_file_name(right_file_path)
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
            slide_number=ppt.slide_count(),
            top=picture_top,
            left=right_picture_left,
            width=picture_width,
        )


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


def add_pictures2(
    file_paths,
    slide_layout=11,
    pic_width=300,
    pic_top1=100,
    pic_top2=300,
    vertical=2,
    horizontal=1,
    font_size=14,
    file_name_to_title=True,
):
    ppt = PowerPoint()
    ppt.setup_active_presentation()

    left_list = []
    top_list = []
    pic_left1 = ppt.slide_width * 0.125 - pic_width * 0.5
    pic_left2 = ppt.slide_width * 0.166 - pic_width * 0.5
    pic_left3 = ppt.slide_width * 0.25 - pic_width * 0.5
    pic_left4 = ppt.slide_width * 0.375 - pic_width * 0.5
    pic_left5 = ppt.slide_width * 0.5 - pic_width * 0.5
    pic_left6 = ppt.slide_width * 0.625 - pic_width * 0.5
    pic_left7 = ppt.slide_width * 0.75 - pic_width * 0.5
    pic_left8 = ppt.slide_width * 0.834 - pic_width * 0.5
    pic_left9 = ppt.slide_width * 0.875 - pic_width * 0.5

    if vertical == 1 and horizontal == 1:
        left_list = [pic_left5]
        top_list.append(pic_top1)

    if vertical == 1 and horizontal == 2:
        left_list = [pic_left3, pic_left7]
        top_list = [pic_top1, pic_top1]

    if vertical == 1 and horizontal == 3:
        left_list = [pic_left2, pic_left5, pic_left8]
        top_list = [pic_top1, pic_top1, pic_top1]

    if vertical == 1 and horizontal == 4:
        left_list = [
            pic_left1,
            pic_left4,
            pic_left6,
            pic_left9,
        ]
        top_list = [pic_top1, pic_top1, pic_top1, pic_top1]

    if vertical == 2 and horizontal == 1:
        left_list = [pic_left5, pic_left5]
        top_list = [pic_top1, pic_top2]

    if vertical == 2 and horizontal == 2:
        left_list = [
            pic_left3,
            pic_left7,
            pic_left3,
            pic_left7,
        ]
        top_list = [pic_top1, pic_top1, pic_top2, pic_top2]

    if vertical == 2 and horizontal == 3:
        left_list = [
            pic_left2,
            pic_left5,
            pic_left8,
            pic_left2,
            pic_left5,
            pic_left8,
        ]
        top_list = [pic_top1, pic_top1, pic_top1, pic_top2, pic_top2, pic_top2]

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
            pic_top1,
            pic_top1,
            pic_top1,
            pic_top1,
            pic_top2,
            pic_top2,
            pic_top2,
            pic_top2,
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
                width=pic_width,
                height=10,
                font_size=font_size,
            )
            ppt.add_picture(
                file_path=file_paths[j][i],
                slide_number=ppt.slide_count(),
                top=top_list[j],
                left=left_list[j],
                width=pic_width,
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
