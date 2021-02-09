from vosk import Model, KaldiRecognizer
import os
import pyaudio
import json
import win32com.client as win32
from pathlib import Path

word = None

if not os.path.exists("model-ru"):
    print("No models found!")
    exit(1)


def call(a):
    if a == "текст":
        print("Result: ", a)
        return 1
    if a == "запустить программу":
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Visible = True
        return 1
    if a == "открыть файл":
        f = os.path.abspath(os.curdir) + "/test.docx"
        word = win32.gencache.EnsureDispatch("Word.Application")
        word.Documents.Open(f)
        return 1

audio = pyaudio.PyAudio()
stream = audio.open(format=pyaudio.paInt16, channels=1,
                    rate=16000, input=True, frames_per_buffer=8000)
stream.start_stream()

model = Model("model-ru")
recognizer = KaldiRecognizer(model, 16000)

while True:
    data = stream.read(2000, exception_on_overflow=False)

    if len(data) == 0:
       break
    if recognizer.AcceptWaveform(data):
        a = json.loads(recognizer.Result())['text']
        call(a)
    else:
        print(json.loads(recognizer.PartialResult())['partial'])
