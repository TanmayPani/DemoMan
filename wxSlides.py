import wx
from wx.media import MediaCtrl
from wx.grid import Grid

import pptx
from pptx.util import Pt


def set_table_font_size(table, size_pt):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size_pt)


class wxSlide(wx.Panel):
    def __init__(self, parent, layout=6, **kwargs):
        super().__init__(parent, **kwargs)
        self.layout = layout
        self.placeholders = []
        self.controls = []

    def makeSlideLayoutSizer(self):
        raise NotImplementedError

    def makeSlideLayout(self):
        lSizer = self.makeSlideLayoutSizer()
        mainSizer = self.GetSizer()

        if mainSizer is None:
            self.SetSizer(lSizer)
        else:
            mainSizer.Add(lSizer)
            self.Layout()

    def addTextField(self, label, value):
        self.placeholders.append(wx.StaticBoxSizer(wx.HORIZONTAL, self))
        box = self.placeholders[-1].GetStaticBox()
        self.placeholders[-1].Add(wx.StaticText(box, label=label), 0, wx.ALL)
        self.controls.append(wx.TextCtrl(box, value=value))
        self.placeholders[-1].Add(self.controls[-1], wx.SizerFlags(1).Expand())

    def addMovie(self, file_name):
        self.controls.append(MediaCtrl(self, fileName=file_name))

    def addTable(self, data, col_names=None):
        grid = Grid(self)
        rows = data.shape[0]
        cols = data.shape[1]
        grid.CreateGrid(rows, cols)
        col_names = col_names or data.columns
        for col, col_name in enumerate(col_names):
            grid.SetColLabelValue(col, col_names)
        grid.UseNativeColHeader()
        for i, row_data in enumerate(data.values.tolist()):
            for j, cell_data in enumerate(row_data):
                grid.SetCellValue(i, j, str(cell_data))
        self.controls.append(grid)


class wxTitleSlide(wxSlide):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, layout=1, **kwargs)
        self.addTextField("Title", "Example Title")
        self.addTextField("Subtitle", "Example SubTitle")

    def makeSlideLayoutSizer(self):
        lSizer = wx.BoxSizer(wx.VERTICAL)
        for ph in self.placeholders:
            lSizer.Add(ph, wx.SizerFlags(1).Expand())
        return lSizer


class wxTitleOnlySlide(wxSlide):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, layout=5, **kwargs)
        self.addTextField("Title", "Example Title")
        self.makeSlideLayout()

    def makeSlideLayoutSizer(self):
        lSizer = wx.BoxSizer(wx.VERTICAL)
        lSizer.Add(self.placeholders[0], wx.SizerFlags(1).Expand().Border())
        return lSizer


class wxStepSlide(wxSlide):
    def __init__(self, parent, text, file_name, table1, table2, title="", **kwargs):
        super().__init__(parent, layout=5, **kwargs)
        self.addTextField("Title", title)
        self.placeholders.append(wx.BoxSizer(wx.HORIZONTAL))

        self.placeholders.append(wx.StaticBoxSizer(wx.VERTICAL, self))
        self.placeholders[-2].Add(self.placeholders[-1], wx.SizerFlags(1).Expand())
        self.controls.append(wx.TextCtrl(self, value=text))
        self.placeholders[-1].Add(self.controls[-1], wx.SizerFlags(1).Expand())
        self.addTable(table1)
        self.placeholders[-1].Add(self.controls[-1], wx.SizerFlags(1).Expand())

        self.placeholders.append(wx.StaticBoxSizer(wx.VERTICAL, self))
        self.placeholders[-3].Add(self.placeholders[-1], wx.SizerFlags(1).Expand())
        self.addMovie(file_name)
        self.placeholders[-1].Add(self.controls[-1], wx.SizerFlags(1).Expand())
        self.addTable(table2)
        self.placeholders[-1].Add(self.controls[-1], wx.SizerFlags(1).Expand())

        self.makeSlideLayout()

    def makeSlideLayoutSizer(self):
        lSizer = wx.BoxSizer(wx.VERTICAL)
        lSizer.Add(self.placeholders[0], wx.SizerFlags(1).Expand().Border())
        return lSizer

    def addToPresentation(self, pres):
        slide = super().addToPresentation(pres)
        slide.shapes.title.text = self.controls[0].GetValue()
        return slide


class wxPresentation(wx.Notebook):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.AddPage(wxTitleSlide(self), "Setup")

    def AddStep(self, text, file_name, table1, table2):
        title = f"Step {self.GetPageCount()}"
        self.AddPage(wxStepSlide(self, text, file_name, table1, table2, title), title)
