from app.app import add_picture
from app.app import add_all_slides
from app.app import add_pictures
from app.app import delete_all_slides
import os

if __name__ == "__main__":
    files = [os.getcwd() + "/data/a.png"]
    delete_all_slides()
    add_all_slides(30)
    add_picture(file_paths=files, slide_layout=24, font_size=50)
    add_pictures(
        left_file_paths=files,
        right_file_paths=files,
        picture_top=200,
        picture_width=400,
        slide_layout=11,
        font_size=40,
    )
