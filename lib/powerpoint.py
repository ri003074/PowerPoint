import win32com.client


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

    def add_picture(self, file_path, left=0, top=0, width=0):
        picture = self.active_presentation.Slides(
            self.slide_count()
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

    def slide_count(self):
        return self.active_presentation.Slides.Count
