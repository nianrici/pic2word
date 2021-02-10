import win32com.client as win32
import os

# Creando el word
wordApp = win32.gencache.EnsureDispatch('Word.Application')  # Abrir un word nuevo
wordApp.Visible = True  # Mantiene abierto el word para verificar que todo se va haciendo correctamente
doc = wordApp.Documents.Add()  # crear word intermedio

# Darle formato
doc.PageSetup.RightMargin = 20
doc.PageSetup.LeftMargin = 20
doc.PageSetup.Orientation = win32.constants.wdOrientLandscape
# A4 en p?xeles: 595x842
doc.PageSetup.PageWidth = 595
doc.PageSetup.PageHeight = 842

# Insertar tabla
my_dir = "."  # si se va a ejecutar desde otra carpeta, c?mbialo
filenames = os.listdir(my_dir)
piccount = 0
file_count = 0

for i in filenames:
    if i[len(i) - 3: len(i)].upper() == 'PNG':  # Cambiar por el formato de las im?genes
        piccount = piccount + 1
    elif i[len(i) - 3: len(i)].upper() == 'JPG':
        piccount = piccount + 1
print(piccount, " images will be inserted")

# mostrar los nombres de las im?genes
total_column = 2
total_row = int(piccount / total_column) + 2
rng = doc.Range(0, 0)
rng.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
table = doc.Tables.Add(rng, total_row, total_column)
table.Borders.Enable = False
if total_column > 1:
    table.Columns.DistributeWidth()

# coger todas las im?genes de la carpeta y meterlas en el mismo documento
frame_max_width = 167  # ancho m?ximo
frame_max_height = 125  # alto m?ximo

piccount = 1

# for isdir in os.walk("."):  # bucle 1 (directorios)
for index, filename in enumerate(filenames):  # bucle 2 (archivos)
    if os.path.isfile(
            os.path.join(os.path.abspath(my_dir), filename)):  # verifica si el objeto doc ya ha sido guardado
        if filename[len(filename) - 3: len(filename)].upper() == 'PNG':  # cambiar por el formato
            piccount = piccount + 1
            print(filename, len(filename), filename[len(filename) - 3: len(filename)].upper())

            cell_column = (piccount % total_column + 1)  # calcular la posici?n de las im?genes en las celdas
            cell_row = (piccount / total_column + 1)
            print('cell_column=%s,cell_row=%s' % (cell_column, cell_row))

            # Dar formato a las celdas
            cell_range = table.Cell(cell_row, cell_column).Range
            cell_range.ParagraphFormat.LineSpacingRule = win32.constants.wdLineSpaceSingle
            cell_range.ParagraphFormat.SpaceBefore = 0
            cell_range.ParagraphFormat.SpaceAfter = 3

            # Insertar las im?genes
            current_pic = cell_range.InlineShapes.AddPicture(os.path.join(os.path.abspath(my_dir), filename))
            width, height = (frame_max_height * frame_max_width / frame_max_height, frame_max_height)

            # Modificar el tama?o de la celda
            current_pic.Height = height
            current_pic.Width = width

            # a?adir el nombre del archivo a cada celda
            table.Cell(cell_row, cell_column).Range.InsertAfter("\n" + filename)
        else:
            continue
