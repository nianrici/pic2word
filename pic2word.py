import win32com.client as win32
import os
import math

# Creando el word
wordApp = win32.gencache.EnsureDispatch(
    'Word.Application')  # Abrir un word nuevo
# Mantiene abierto el word para verificar que todo se va haciendo correctamente
wordApp.Visible = True
doc = wordApp.Documents.Add()  # crear word intermedio

# Darle formato
doc.PageSetup.RightMargin = 20
doc.PageSetup.LeftMargin = 20
doc.PageSetup.Orientation = win32.constants.wdOrientLandscape
# A4 en píxeles: 595x842
doc.PageSetup.PageWidth = 595
doc.PageSetup.PageHeight = 842

# Insertar tabla
my_dir = "."  # si se va a ejecutar desde otra carpeta, cámbialo
filenames = [f.name for f in os.scandir(
    my_dir) if f.is_file() and f.name.upper().endswith(('.PNG', '.JPG'))]
piccount = len(filenames)
print(piccount, " images will be inserted")

# mostrar los nombres de las imágenes
total_column = 2
total_row = math.ceil(piccount / total_column) + 2
rng = doc.Range(0, 0)
rng.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
table = doc.Tables.Add(rng, total_row, total_column)
table.Borders.Enable = False
if total_column > 1:
    table.Columns.DistributeWidth()

# coger todas las imágenes de la carpeta y meterlas en el mismo documento
frame_max_width = 167  # ancho máximo
frame_max_height = 125  # alto máximo

piccount = 1

for filename in filenames:
    print(filename)
    # calcular la posición de las imágenes en las celdas
    cell_column = (piccount % total_column + 1)
    cell_row = (piccount // total_column + 1)
    print('cell_column={}, cell_row={}'.format(cell_column, cell_row))

    # Dar formato a las celdas
    cell_range = table.Cell(cell_row, cell_column).Range
    cell_range.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceSingle
    cell_range.ParagraphFormat.SpaceBefore = 0
    cell_range.ParagraphFormat.SpaceAfter = 3

    # Insertar las imágenes
    current_pic = cell_range.InlineShapes.AddPicture(
        os.path.join(os.path.abspath(my_dir), filename))
    width, height = (frame_max_height * frame_max_width /
                     frame_max_height, frame_max_height)

    # Modificar el tamaño de la celda
    current_pic.Height = height
    current_pic.Width = width

    # añadir el nombre del archivo a cada celda
    table.Cell(cell_row, cell_column).Range.InsertAfter(
        "\n{}".format(filename))
    piccount += 1
