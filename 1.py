# импорт библиотек
from vosk import Model, KaldiRecognizer
import os
import pyaudio
import json
import win32com.client as win32
from pathlib import Path

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
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = True
        return 1
    # ветка открыть файл
    if a == "открыть файл":
        f = os.path.abspath(os.curdir) + "/test.docx"
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Documents.Open(f)
        return 1
    # ветка создание документа
    if a == "создать документ":
        word = win32.gencache.EnsureDispatch("Word.Application")
        globals()['doc'] = word.Documents.Add()
        doc = globals()['doc']
        doc.Content.Text = "Привет мир!"
    # ветка сохранить документа
    if a == "сохранить документ":
        f = os.path.abspath(os.curdir) + "/test3.docx"
        word = win32.gencache.EnsureDispatch("Word.Application")
        doc = globals()['doc']
        doc.SaveAs(f, FileFormat = 16)

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
