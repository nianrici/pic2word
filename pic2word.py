import win32com.client as win32
import os
import math

wordApp = win32.gencache.EnsureDispatch('Word.Application')
wordApp.Visible = True
doc = wordApp.Documents.Add()

doc.PageSetup.RightMargin = 20
doc.PageSetup.LeftMargin = 20
doc.PageSetup.Orientation = win32.constants.wdOrientLandscape
# A4 en píxeles: 595x842
doc.PageSetup.PageWidth = 595
doc.PageSetup.PageHeight = 842

my_dir = "."
piccount = 0

for root, dirs, files in os.walk(my_dir):
    for filename in files:
        if filename.upper().endswith(('.PNG', '.JPG')):
            piccount += 1
print(piccount, " images will be inserted")

total_column = 2
total_row = math.ceil(piccount / total_column) + 2
rng = doc.Range(0, 0)
rng.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
table = doc.Tables.Add(rng, total_row, total_column)
table.Borders.Enable = False
if total_column > 1:
    table.Columns.DistributeWidth()

frame_max_width = 167
frame_max_height = 125

piccount = 1

for root, dirs, files in os.walk(my_dir):
    for filename in files:
        if filename.upper().endswith(('.PNG', '.JPG')):
            print(filename)
            cell_column = (piccount % total_column + 1)
            cell_row = (piccount // total_column + 1)
            print('cell_column={}, cell_row={}'.format(cell_column, cell_row))

            cell_range = table.Cell(cell_row, cell_column).Range
            cell_range.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceSingle
            cell_range.ParagraphFormat.SpaceBefore = 0
            cell_range.ParagraphFormat.SpaceAfter = 3

            current_pic = cell_range.InlineShapes.AddPicture(
                os.path.join(root, filename))
            width, height = (frame_max_height * frame_max_width /
                             frame_max_height, frame_max_height)

            current_pic.Height = height
            current_pic.Width = width

            table.Cell(cell_row, cell_column).Range.InsertAfter(
                "\n{}".format(filename))
            piccount += 1
