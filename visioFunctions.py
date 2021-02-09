# добавляем объект в файл visio
def add(page, type, x, y, text):
    visioShape = page.Drop(type, x, y)
    visioShape.Text = text
    return visioShape

# соединяем снизу два блока
def connectDefault(visio, page, shape, shape2):
    con = visio.Application.ConnectorToolDataObject
    shapeConnect = page.Drop(con, 0, 0)
    shapeConnect.CellsU("BeginX").GlueTo(shape.CellsU("PinX"))
    shapeConnect.CellsU("EndX").GlueTo(shape2.CellsU("PinX"))

# соединяем два блока
def connect(visio, page, shape, shape2, glueBegin, glueEnd):
    con = visio.Application.ConnectorToolDataObject
    shapeConnect = page.Drop(con, 0, 0)
    shapeConnect.CellsU("BeginX").GlueTo(shape.CellsU(glueBegin))
    shapeConnect.CellsU("EndX").GlueTo(shape2.CellsU(glueEnd))

# получение названия элемента
def getTemplateName(visio, name):
    obj = visio.Documents("BASFLO_M.VSSX")
    #return obj.Masters(name)
    return obj.Masters(name)