# импорт библиотек
from vosk import Model, KaldiRecognizer
import os
import pyaudio
import json
import win32com.client as win32
from pathlib import Path
import visioFunctions as funs
import visioBlocks as blocks

# проверка наличия языковой модели
# выход из программы в случае отсутсвия модели с ошибкой 1
if not os.path.exists("model-ru"):
    print("No models found!")
    exit(1)

# функция обработки команды
def call(a):
    # тестовое условие
    if a == "текст":
        print("Result: ", a)
        return 1
    # ветка запуска программа
    if a == "запустить программу":
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        visio.Visible = True
        return 1
    # ветка открыть файл
    if a == "открыть файл":
        f = os.path.abspath(os.curdir) + "/result/task.vsdx"
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        visio.Documents.Open(f)
        return 1
    # ветка создание документа
    if a == "создать документ":
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        f = os.path.abspath(os.curdir) + "/temp.vsdx"
        globals()['doc'] = visio.Documents.Add("BASFLO_M.VSTX")
        doc = globals()['doc']
    # ветка сохранить документа
    if a == "сохранить документ":
        f = os.path.abspath(os.curdir) + "/result/task2.vsdx"
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        doc = globals()['doc']
        doc.SaveAs(f, FileFormat = 16)
    # ветка завершения работы
    if a == "завершить программу":
        visio = win32.gencache.EnsureDispatch("Visio.Application")
        visio.Visible = False
        exit(0)

# инициализация микрофона
audio = pyaudio.PyAudio()
stream = audio.open(format=pyaudio.paInt16, channels=1,
                    rate=16000, input=True, frames_per_buffer=8000)
stream.start_stream()

# инициализация русской языковой модели
model = Model("model-ru")
recognizer = KaldiRecognizer(model, 16000)

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
        call(a)
    else:
        # Debug
        print(json.loads(recognizer.PartialResult())['partial'])
