from app.app import add_picture
from app.app import add_all_slides
from app.app import add_pictures
from app.app import add_picture_to_placeholder
from app.app import delete_all_slides
import os

if __name__ == "__main__":
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
    delete_all_slides()
    add_all_slides(30, placeholder_number=True)

    add_picture(
        file_paths=files1, slide_layout=16, font_size=50
    )  # to only 1 placeholder
    add_pictures(
        left_file_paths=files1,
        right_file_paths=files2,
        picture_top=200,
        picture_width=400,
        slide_layout=11,
        font_size=40,
    )
    add_picture_to_placeholder(files1, files2, slide_layout=29)
    add_picture_to_placeholder(files1, files2, files3, slide_layout=24)
