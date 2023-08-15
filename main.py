from app.app import add_picture
from app.app import add_all_slides
from app.app import add_pictures
from app.app import add_picture_to_placeholder
from app.app import delete_all_slides
from app.app import add_table
from app.ss_2_ppt import ss_2_ppt
import os


def main1():
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
    delete_all_slides()
    add_all_slides(30, placeholder_number=True)

    # to only 1 placeholder
    add_picture(file_paths=files1, slide_layout=16, font_size=50)
    add_pictures(
        left_file_paths=files1,
        right_file_paths=files2,
        picture_top=200,
        picture_width=400,
        slide_layout=11,
        font_size=40,
    )
    add_picture_to_placeholder(files1, files2, slide_layout=34)
    add_picture_to_placeholder(
        files1, files2, slide_layout=34, title_placeholder_numbers=[2, 4]
    )
    add_picture_to_placeholder(files1, files2, files3, slide_layout=24)

    add_table(data=data1, cell_width=[100, 200])


if __name__ == "__main__":
    ss_2_ppt()
