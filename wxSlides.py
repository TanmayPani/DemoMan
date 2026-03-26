from collections import defaultdict

import wx
from wx.media import MediaCtrl, MC_NO_AUTORESIZE
from wx.grid import Grid

import pptx
from pptx.util import Pt, Cm
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR


def set_table_font_size(table, size_pt):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size_pt)


def set_text_font_size(tbox, size_pt):
    for paragraph in tbox.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size_pt)


class wxShape(wx.StaticBoxSizer):
    def __init__(self, parent, orient=wx.HORIZONTAL, title=""):
        super().__init__(orient, parent, label=title)
        # if title is not None:
        #    self.Add(
        #        wx.StaticText(self.StaticBox, id=wx.ID_ANY, label=title), 0, wx.ALL
        #    )

    # def SaveToSlide(self, slide, *args, **kwargs):
    #    return slide.shapes.add_group_shape()


class wxTextBox(wxShape):
    def __init__(self, parent, orient=wx.HORIZONTAL, title="", **kwargs):
        super().__init__(parent, orient=orient, title=title)
        self.textCtrl = wx.TextCtrl(self.StaticBox, **kwargs)
        self.Add(self.textCtrl, wx.SizerFlags(1).Expand().Border())

    @property
    def Text(self):
        return self.textCtrl.GetValue()

    @Text.setter
    def Text(self, text):
        self.textCtrl.SetValue(text)

    def SaveToSlide(self, slide, left, top, width, height, wrap=True, font_size=None):
        tbox = slide.shapes.add_textbox(left, top, width, height)
        tbox.text_frame.text = self.Text
        tbox.text_frame.word_wrap = wrap
        if font_size is not None:
            set_text_font_size(tbox, font_size)
        return tbox


class wxMovieShape(wxShape):
    def __init__(
        self,
        parent,
        orient=wx.HORIZONTAL,
        title="",
        file_name=None,
        style=MC_NO_AUTORESIZE,
        **kwargs,
    ):
        super().__init__(parent, orient=orient, title=title)
        self.movieCtrl = MediaCtrl(self.StaticBox, style=style, **kwargs)
        self.movieCtrl.ShowPlayerControls()
        self.Add(self.movieCtrl, wx.SizerFlags(1).Expand().Border())
        if file_name is not None:
            self.LoadVideo(file_name)

    def LoadVideo(self, file_name):
        self.fileName = file_name
        status = self.movieCtrl.Load(self.fileName)
        if not status:
            raise ValueError(f"Can't load {self.fileName}, not a valid video file")

    def SaveToSlide(
        self,
        slide,
        left,
        top,
        width,
        height,
        mime_type="video/mp4",
        poster_frame_image=None,
    ):
        movie = slide.shapes.add_movie(
            self.fileName,
            left,
            top,
            width,
            height,
            mime_type=mime_type,
            poster_frame_image=poster_frame_image,
        )


class wxTableShape(wxShape):
    def __init__(self, parent, orient=wx.HORIZONTAL, title="", **kwargs):
        super().__init__(parent, orient=orient, title=title)
        self.tableCtrl = Grid(self.StaticBox, **kwargs)
        self.tableCtrl.UseNativeColHeader()
        self.Add(self.tableCtrl, wx.SizerFlags(1).Expand().Border())

    def CreateGrid(self, num_rows, num_cols):
        self.tableCtrl.CreateGrid(num_rows, num_cols)

    def SetColumnNames(self, column_names):
        numCols = self.tableCtrl.GetNumberCols()
        if len(column_names) != numCols:
            raise ValueError(
                f"Sequence of column_names (given length {len(column_names)}) must be of same length as num_cols ({numCols})"
            )
        for icol, colName in enumerate(column_names):
            self.tableCtrl.SetColLabelValue(icol, str(colName))

    def SetTableData(self, tabular_data):
        for irow, row_data in enumerate(tabular_data):
            for icol, cell_data in enumerate(row_data):
                self.tableCtrl.SetCellValue(irow, icol, str(cell_data))

    def SetFromDataframe(self, df):
        numRows, numCols = df.shape
        self.CreateGrid(numRows, numCols)
        self.SetColumnNames(df.columns)
        self.SetTableData(df.values.tolist())

    def SaveToSlide(self, slide, left, top, width, height, font_size=None):
        rows = self.tableCtrl.GetNumberRows()
        cols = self.tableCtrl.GetNumberCols()
        tableFrame = slide.shapes.add_table(rows + 1, cols, left, top, width, height)
        table = tableFrame.table
        for icol in range(cols):
            table.cell(0, icol).text = self.tableCtrl.GetColLabelValue(icol)

        for irow in range(rows):
            for icol in range(cols):
                table.cell(irow + 1, icol).text = self.tableCtrl.GetCellValue(
                    irow, icol
                )

        if font_size is not None:
            set_table_font_size(table, font_size)

        return tableFrame


class wxSlide(wx.Panel):
    LAYOUT = 6

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.shapes = defaultdict(list)
        self.SlideLayout()

    def MakeSlideLayout(self):
        return wx.BoxSizer(wx.VERTICAL)

    def SlideLayout(self):
        lSizer = self.MakeSlideLayout()
        mainSizer = self.GetSizer()

        if mainSizer is None:
            self.SetSizer(lSizer)
        else:
            mainSizer.Add(lSizer)
            self.Layout()

    def AddTextBox(self, *args, **kwargs):
        self.shapes["textbox"].append(wxTextBox(self, *args, **kwargs))

    def AddMovie(self, *args, **kwargs):
        self.shapes["movie"].append(wxMovieShape(self, *args, **kwargs))

    def AddTable(self, *args, **kwargs):
        self.shapes["table"].append(wxTableShape(self, *args, **kwargs))

    def SaveToPres(self, pres):
        print(f"Slide: {len(pres.slides)}, layout: {self.LAYOUT}")
        return pres.slides.add_slide(pres.slide_layouts[self.LAYOUT])


class wxTitleSlide(wxSlide):
    LAYOUT = 0

    def MakeSlideLayout(self):
        mainSizer = super().MakeSlideLayout()
        self.AddTextBox(title="Title")
        self.AddTextBox(title="Subtitle")
        self.Title = "Example title"
        self.SubTitle = "Example subtitle"
        for shape in self.shapes["textbox"]:
            mainSizer.Add(shape, wx.SizerFlags(0).Expand().DoubleBorder())
        return mainSizer

    @property
    def Title(self):
        return self.shapes["textbox"][0].Text

    @Title.setter
    def Title(self, value):
        self.shapes["textbox"][0].Text = str(value)

    @property
    def SubTitle(self):
        return self.shapes["textbox"][1].Text

    @SubTitle.setter
    def SubTitle(self, value):
        self.shapes["textbox"][1].Text = str(value)

    def SaveToPres(self, pres):
        slide = super().SaveToPres(pres)
        slide.shapes.title.text = self.Title
        slide.placeholders[1].text = self.SubTitle
        return slide


class wxTitleOnlySlide(wxSlide):
    LAYOUT = 5

    def MakeSlideLayout(self):
        mainSizer = super().MakeSlideLayout()
        self.AddTextBox(title="Title")
        self.shapes["textbox"][0].Text = "Example title"
        mainSizer.Add(
            self.shapes["textbox"][0], wx.SizerFlags(0).Expand().DoubleBorder()
        )
        return mainSizer

    @property
    def Title(self):
        return self.shapes["textbox"][0].Text

    @Title.setter
    def Title(self, value):
        self.shapes["textbox"][0].Text = str(value)

    def SaveToPres(self, pres):
        slide = super().SaveToPres(pres)
        slide.shapes.title.text = self.Title
        return slide


class wxStepSlide(wxTitleOnlySlide):
    def __init__(
        self,
        parent,
        title,
        text,
        movie_file_name=None,
        table_data1=None,
        table_data2=None,
        **kwargs,
    ):
        super().__init__(parent, **kwargs)
        self.Title = title
        self.shapes["textbox"][1].Text = text

        if movie_file_name is not None:
            self.shapes["movie"][0].LoadVideo(movie_file_name)

        if table_data1 is not None:
            self.shapes["table"][0].SetFromDataframe(table_data1)
            self.shapes["table"][0].tableCtrl.AutoSizeColumns()
            self.shapes["table"][0].tableCtrl.AutoSizeRows()

        if table_data2 is not None:
            self.shapes["table"][1].SetFromDataframe(table_data2)
            self.shapes["table"][1].tableCtrl.AutoSizeColumns()
            self.shapes["table"][1].tableCtrl.AutoSizeRows()

    def MakeSlideLayout(self):
        mainSizer = super().MakeSlideLayout()

        self.AddTextBox(title="Text:", style=wx.TE_MULTILINE | wx.TE_BESTWRAP)
        self.AddMovie(title="Video:")
        self.AddTable(title="Components:")
        self.AddTable(title="Tool:")

        # for key, val in self.shapes.items():
        #    print(key, len(val))

        newSizer = wx.BoxSizer(wx.HORIZONTAL)

        leftSizer = wx.BoxSizer(wx.VERTICAL)
        leftSizer.Add(
            self.shapes["textbox"][1], wx.SizerFlags(1).Expand().DoubleBorder()
        )
        leftSizer.Add(self.shapes["table"][0], wx.SizerFlags(1).Expand().DoubleBorder())
        newSizer.Add(leftSizer, wx.SizerFlags(1).Expand())

        rightSizer = wx.BoxSizer(wx.VERTICAL)
        rightSizer.Add(
            self.shapes["movie"][0], wx.SizerFlags(1).Expand().DoubleBorder()
        )
        rightSizer.Add(
            self.shapes["table"][1], wx.SizerFlags(1).Expand().DoubleBorder()
        )
        newSizer.Add(rightSizer, wx.SizerFlags(1).Expand())

        mainSizer.Add(newSizer, wx.SizerFlags(1).Expand().Border())

        return mainSizer

    def SaveToPres(self, pres):
        slide = super().SaveToPres(pres)
        _ = self.shapes["textbox"][1].SaveToSlide(
            slide, Cm(0.5), Cm(4), Cm(16), Cm(2), wrap=True, font_size=12
        )

        _ = self.shapes["table"][0].SaveToSlide(
            slide, Cm(0.2), Cm(8), Cm(12), Cm(2.4), font_size=10
        )

        toolTable = self.shapes["table"][1].SaveToSlide(
            slide, Cm(12.4), Cm(8), Cm(12), Cm(2.4), font_size=10
        )
        tblf = toolTable._element.graphic.graphicData.tbl
        tblf[0][-1].text = "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}"

        _ = self.shapes["movie"][0].SaveToSlide(
            slide, Cm(33.87 / 2.0), Cm(0.37), Cm(8), Cm(6)
        )


class wxPresentation(wx.Notebook):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.AddPage(wxTitleSlide(self), "Setup")

    def AddStepSlide(self, title, text, file_name, table1, table2):
        slideIdx = str(self.GetPageCount())
        self.AddPage(
            wxStepSlide(self, title, text, file_name, table1, table2), slideIdx
        )

    def Save(self, destination):
        pres = pptx.Presentation()
        for islide in range(self.GetPageCount()):
            _ = self.GetPage(islide).SaveToPres(pres)

        pres.save(destination)
