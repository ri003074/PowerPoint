import win32com.client
import os
import pywintypes


def get_file_name(file_path):
    return os.path.splitext(os.path.basename(file_path))[0]


def get_dir_name(file_path):
    return os.path.dirname(file_path).split("\\")[-1].split("/")[-1]


class PowerPoint:
    def __init__(self):
        self.application = None
        self.active_presentation = None
        self.slide_width = None
        self.slide_height = None

    def setup_active_presentation(self):
        self.application = win32com.client.Dispatch("PowerPoint.Application")
        self.active_presentation = self.application.ActivePresentation
        self.slide_width = self.active_presentation.PageSetup.SlideWidth
        self.slide_height = self.active_presentation.PageSetup.SlideHeight

    def add_slide(self, slide_layout):
        slide = self.active_presentation.Slides.Add(
            Index=self.slide_count() + 1,
            Layout=slide_layout,
        )
        slide.Select()

    def delete_slide(self, slide_number):
        self.active_presentation.Slides.Item(slide_number).delete()

    def add_picture(self, file_path, slide_number, left=0, top=0, width=0):
        picture = self.active_presentation.Slides(
            slide_number
        ).Shapes.AddPicture(
            FileName=file_path,
            LinkToFile=-1,
            SaveWithDocument=-1,
            Left=left,
            Top=top,
        )
        if width != 0:
            picture.Width = width

    def add_textbox(
        self, text, top, left, width, height, alignment=2, font_size=14
    ):
        textbox = self.active_presentation.Slides(
            self.slide_count()
        ).Shapes.AddTextbox(1, Top=top, Left=left, Width=width, Height=height)
        textbox.TextFrame.TextRange.Text = text
        textbox.TextFrame.TextRange.ParagraphFormat.Alignment = alignment
        textbox.TextFrame.VerticalAnchor = 3
        textbox.TextFrame.TextRange.Font.Size = font_size
        textbox.Top = top - textbox.Height

    def add_text_to_placeholder(self, text, placeholder_number, font_size=14):
        textframe = (
            self.active_presentation.Slides(self.slide_count())
            .Shapes(placeholder_number)
            .TextFrame
        )
        textframe.TextRange.Text = text
        textframe.TextRange.Font.Size = font_size
        textframe.TextRange.ParagraphFormat.Alignment = 2  # center
        textframe.VerticalAnchor = 3  # middle, 4->bottom

    def add_table(self, data, cell_width=None):
        row = len(data)
        column = len(data[0])
        table = (
            self.active_presentation.Slides(self.slide_count())
            .Shapes.AddTable(row, column)
            .Table
        )

        for i in range(row):
            for j in range(column):
                text_range = table.Cell(i + 1, j + 1).Shape.TextFrame.TextRange
                text_range.Text = data[i][j]
                text_range.ParagraphFormat.Alignment = 2
                if cell_width is not None:
                    table.Columns(j + 1).Width = cell_width[j]

        shape = self.active_presentation.Slides(self.slide_count()).Shapes(2)
        shape.Left = self.slide_width / 2 - shape.width / 2
        shape.Top = self.slide_height / 3

    def replace_placeholder_text(
        self, slide_number, placeholder_number, before, after
    ):
        textframe = (
            self.active_presentation.Slides(slide_number)
            .Shapes(placeholder_number)
            .TextFrame
        )
        try:
            text = str(textframe.TextRange.Text)
            textframe.TextRange.Text = text.replace(before, after)
        except pywintypes.com_error:
            pass

    def slide_count(self):
        return self.active_presentation.Slides.Count

    def placeholder_count(self, slide_number):
        return self.active_presentation.Slides(slide_number).Shapes.Count

    def active_slide_number(self):
        return self.application.ActiveWindow.Selection.SlideRange.SlideIndex
