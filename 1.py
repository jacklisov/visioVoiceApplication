from vosk import Model, KaldiRecognizer 
import os
import pyaudio

if not os.path.exists("model-ru"):
    print("No models found!")
    exit(1)

audio = pyaudio.PyAudio()
stream = audio.open(format=pyaudio.paInt16, channels=1, rate=16000, input=True, frames_per_buffer=8000)
stream.start_stream()

model = Model("model-ru")
recognizer = KaldiRecognizer(model, 16000)

while True:
    data = stream.read(2000, exception_on_overflow=False)
   
    if len(data) == 0:
       break
    if recognizer.AcceptWaveform(data):
        print(recognizer.Result())
    else:
        print(recognizer.PartialResult())

print(recognizer.FinalResult())