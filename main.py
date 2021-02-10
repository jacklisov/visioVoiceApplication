# импорт библиотек
from vosk import Model, KaldiRecognizer
import os
import pyaudio
import json
import win32com.client as win32

# проверка наличия языковой модели
# выход из программы в случае отсутсвия модели с ошибкой 1
if not os.path.exists("model-ru"):
    print("No models found!")
    exit(1)



# visio support library
# добавляем объект в файл visio
def add(page, type, x, y, text):
    visioShape = page.Drop(type, x, y)
    visioShape.Text = text
    globals()['lastCmd'] = visioShape
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



# блоки
# блок начало-конец
def beginEnd(x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Начало/окончание"
    visio = globals()['visio']
    beginEnd = getTemplateName(visio, mast)
    add(page, beginEnd, x, y, '')

# блок процесса
def proccess(x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Процесс"
    visio = globals()['visio']
    beginEnd = getTemplateName(visio, mast)
    add(page, beginEnd, x, y, '')

# блок подпроцесса
def subProccess(x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Подпроцесс"
    visio = globals()['visio']
    beginEnd = getTemplateName(visio, mast)
    add(page, beginEnd, x, y, '')

# блок решение
def decision(x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Решение"
    visio = globals()['visio']
    beginEnd = getTemplateName(visio, mast)
    add(page, beginEnd, x, y, '')

# блок данные
def dataEl(x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Данные"
    visio = globals()['visio']
    beginEnd = getTemplateName(visio, mast)
    add(page, beginEnd, x, y, '')

# блок данные
def document(x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Документ"
    visio = globals()['visio']
    beginEnd = getTemplateName(visio, mast)
    add(page, beginEnd, x, y, '')



# команды
def call(a):
    # тестовое условие
    if a == "текст":
        print("Result: ", a)
    # ветка запуска программа
    if a == "запустить программу":
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        globals()['visio'] = visio
        visio.Visible = True
    # ветка открыть файл
    if a == "открыть файл":
        f = os.path.abspath(os.curdir) + "/result/task.vsdx"
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        visio.Documents.Open(f)
    # ветка создание документа
    if a == "создать документ":
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        globals()['visio'] = visio
        f = os.path.abspath(os.curdir) + "/temp.vsdx"
        globals()['doc'] = visio.Documents.Add("BASFLO_M.VSTX")
        doc = globals()['doc']
    # ветка сохранить документа
    if a == "сохранить документ":
        f = os.path.abspath(os.curdir) + "/result/task2.vsdx"
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        doc = globals()['doc']
        doc.SaveAs(f)
    # ветка завершения работы
    if a == "завершить программу":
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        visio.Visible = False
        exit(0)
    # присвоение текста последней фигуре
    if a == "присвоить текст":
        print("Произнесите текст для фигуры")
        pRes = ''
        while True:
            stream = globals()['stream']
            data = stream.read(2000, exception_on_overflow=False)
            if globals()['rec'].AcceptWaveform(data):
                pRes = json.loads(globals()['rec'].Result())['text']
                if pRes != None and pRes != '':
                    globals()['lastCmd'].Text = pRes
                    break
        print("В случае некоретного распознавания произнесите: 'Изменить текст последней фигуры'")
    # изменение текста последней фигуре
    if a == "изменить текст":
        print("Произнесите новый текст для фигуры")
        pRes = ''
        while True:
            stream = globals()['stream']
            data = stream.read(2000, exception_on_overflow=False)
            if globals()['rec'].AcceptWaveform(data):
                pRes = json.loads(globals()['rec'].Result())['text']
                if pRes != None and pRes != '':
                    globals()['lastCmd'].Text = pRes
                    break
        print("В случае некоретного распознавания произнесите: 'Изменить текст последней фигуры'")
    # соединение элементов снизу
    if a == "соединить элементы":
        if "end" in globals():
            if "root1" in globals():
                connectDefault(globals()['visio'], globals()['doc'].Pages.Item(1), globals()['lastElement1'], globals()['lastCmd'])
            if "root2" in globals():
                connectDefault(globals()['visio'], globals()['doc'].Pages.Item(1), globals()['lastElement2'], globals()['lastCmd'])
            if "root3" in globals():
                connectDefault(globals()['visio'], globals()['doc'].Pages.Item(1), globals()['lastTree'], globals()['lastCmd'])

        if "lastCmd" in globals():
            connectDefault(globals()['visio'], globals()['doc'].Pages.Item(1), globals()['preLastCmd'], globals()['lastCmd'])
        else:
            print ("Должно быть добавлено больше двух элементов")
    # ветка добавления начала
    if a == "добавить начала" or a == "добавить конец" or a == "добавить начало":
        if "lastCmd" in globals():
            globals()['preLastCmd'] = globals()['lastCmd'] 

        beginEnd(globals()['x'], globals()['y'])
        
        globals()['y'] = globals()['y'] - 1
        print("Чтобы присвоить текст произнесите команду: 'Присвоить текст'")
        print("Для соединения элементов произнесите команду: 'Соединить элементы'")
        return 1
    # ветка добавления процесса
    if a == "добавить процесс":
        if "lastCmd" in globals():
            globals()['preLastCmd'] = globals()['lastCmd'] 

        proccess(globals()['x'], globals()['y'])
        
        globals()['y'] = globals()['y'] - 1
        print("Чтобы присвоить текст произнесите команду: 'Присвоить текст'")
        print("Для соединения элементов произнесите команду: 'Соединить элементы'")
        return 1
    # ветка добавления подпроцесса
    if a == "добавить под процесс":
        if "lastCmd" in globals():
            globals()['preLastCmd'] = globals()['lastCmd'] 

        subProccess(globals()['x'], globals()['y'])
        
        globals()['y'] = globals()['y'] - 1
        print("Чтобы присвоить текст произнесите команду: 'Присвоить текст'")
        print("Для соединения элементов произнесите команду: 'Соединить элементы'")
        return 1
    # ветка добавления данных
    if a == "добавить данные":
        if "lastCmd" in globals():
            globals()['preLastCmd'] = globals()['lastCmd'] 

        dataEl(globals()['x'], globals()['y'])
        
        globals()['y'] = globals()['y'] - 1
        print("Чтобы присвоить текст произнесите команду: 'Присвоить текст'")
        print("Для соединения элементов произнесите команду: 'Соединить элементы'")
        return 1
    # ветка добавления документа
    if a == "добавить документ":
        if "lastCmd" in globals():
            globals()['preLastCmd'] = globals()['lastCmd'] 

        document(globals()['x'], globals()['y'])
        
        globals()['y'] = globals()['y'] - 1
        print("Чтобы присвоить текст произнесите команду: 'Присвоить текст'")
        print("Для соединения элементов произнесите команду: 'Соединить элементы'")
        return 1
    # ветка добавления документа
    if a == "добавить решение":
        if "lastCmd" in globals():
            globals()['preLastCmd'] = globals()['lastCmd'] 

        decision(globals()['x'], globals()['y'])
        
        globals()['main'] = globals()['lastCmd']
        globals()['y'] = globals()['y'] - 1
        globals()['rootX'] = globals()['x']
        print("Чтобы присвоить текст произнесите команду: 'Присвоить текст'")
        print("Для соединения элементов произнесите команду: 'Соединить элементы'")
        print("Для завершения блока решения произнесите команду: 'Завершить решение'")
        return 1
    # ветка добавления первой ветки
    if a == "добавить в первую ветку":
        globals()['root1'] = True
        globals()['lastCmd'] = globals()['main']
        globals()['x'] = globals()['rootX'] - 1
        print('Теперь голосом добавьте элементы')
        print('Для соединения используйте команду: "соединить элементы"')
    # ветка добавления первой ветки
    if a == "добавить во вторую ветку":
        globals()['lastElement1'] = globals()['lastCmd']
        globals()['lastCmd'] = globals()['main']
        globals()['root2'] = True
        globals()['x'] = globals()['rootX'] + 1
        print('Теперь голосом добавьте элементы')
        print('Для соединения используйте команду: "соединить элементы"')
    # ветка добавления первой ветки
    if a == "добавить в третью ветку":
        globals()['lastElement2'] = globals()['lastCmd']
        globals()['lastCmd'] = globals()['main']
        globals()['root3'] = True
        globals()['x'] = globals()['rootX']
        print('Теперь голосом добавьте элементы')
        print('Для соединения используйте команду: "соединить элементы"')
    # ветка завершения решения
    if a == "завершить решение":
        globals()['lastTree'] = globals()['lastCmd']
        globals()['end'] = True
        globals()['x'] = globals()['rootX']

# координаты первого элемента
globals()['x'] = 6
globals()['y'] = 7

# инициализация микрофона
audio = pyaudio.PyAudio()
stream = audio.open(format=pyaudio.paInt16, channels=1,
                    rate=16000, input=True, frames_per_buffer=8000)
stream.start_stream()
globals()['stream'] = stream

# инициализация русской языковой модели
model = Model("model-ru")
recognizer = KaldiRecognizer(model, 16000)
globals()['rec'] = recognizer

# бесконечный цикл прослушки с микрофона
while True:
    # прослушка микрофона
    data = stream.read(2000, exception_on_overflow=False)

    # обработка слов
    if len(data) == 0:
        # прервать цикл в случае отсутвия сигнала от микрофонв
        break
    if recognizer.AcceptWaveform(data):
        # раскодировка JSON 
        a = json.loads(recognizer.Result())['text']
        
        # вызов функции обработки слов
        res = call(a)
