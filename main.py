from app.app import add_pictures
from app.app import add_all_slides
from app.app import add_pictures
from app.app import add_picture_to_placeholder
from app.app import delete_all_slides
from app.app import add_table
from app.app import add_picture_to_active_slide
from app.app import replace_text
from app.ss_2_ppt import ss_2_ppt
import os
from app.pic import pic

files1 = [
    os.getcwd() + "/data/a.png",
    os.getcwd() + "/data/a.png",
    os.getcwd() + "/data/a.png",
    os.getcwd() + "/data/a.png",
]
files2 = [
    os.getcwd() + "/data/b.png",
    os.getcwd() + "/data/b.png",
    os.getcwd() + "/data/b.png",
    os.getcwd() + "/data/b.png",
]
files3 = [
    os.getcwd() + "/data/a.png",
    os.getcwd() + "/data/b.png",
    os.getcwd() + "/data/a.png",
    os.getcwd() + "/data/b.png",
]

data1 = [["a", "b"], [1, 2], [4, 5]]


def main1():
    delete_all_slides()
    # add_all_slides(30, placeholder_number=True)

    # to only 1 placeholder
    add_pictures(file_paths=[files1], slide_layout=16, font_size=50)
    add_pictures(
        file_paths=[files1, files2],
        pic_top1=200,
        pic_top2=400,
        slide_layout=11,
        font_size=40,
    )
    add_picture_to_placeholder([files1, files2], slide_layout=34)
    add_picture_to_placeholder(
        [files1, files2], slide_layout=34, title_placeholder_numbers=[2, 4]
    )
    add_picture_to_placeholder([files1, files2, files3], slide_layout=24)

    add_table(data=data1, cell_width=[100, 200])


def main2():
    add_picture_to_active_slide(os.getcwd() + "/data/a.png")


def main3():
    add_picture_to_placeholder(
        [files1],
        title_placeholder_numbers=[1],
        slide_layout=29,
        file_name_to_title=False,
    )


def main4():
    replace_text(slide_number=15, file="sample.txt")


if __name__ == "__main__":
    delete_all_slides()
    main1()
    main2()
    main3()
    main4()
    # ss_2_ppt()
    # pic()
