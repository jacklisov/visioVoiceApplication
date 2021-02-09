import visioFunctions as fun
import win32com.client as win32

# блок начало-конец
def beginEnd(text, x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Начало/окончание"
    visio = globals()['visio']
    beginEnd = fun.getTemplateName(visio, mast)
    fun.add(page, beginEnd, x, y, text)

# блок процесса
def proccess(text, x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Процесс"
    visio = globals()['visio']
    beginEnd = fun.getTemplateName(visio, mast)
    fun.add(page, beginEnd, x, y, text)

# блок подпроцесса
def subProccess(text, x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Подпроцесс"
    visio = globals()['visio']
    beginEnd = fun.getTemplateName(visio, mast)
    fun.add(page, beginEnd, x, y, text)

# блок решение
def decision(text, x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Решение"
    visio = globals()['visio']
    beginEnd = fun.getTemplateName(visio, mast)
    fun.add(page, beginEnd, x, y, text)

# блок данные
def data(text, x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Данные"
    visio = globals()['visio']
    beginEnd = fun.getTemplateName(visio, mast)
    fun.add(page, beginEnd, x, y, text)

# блок данные
def document(text, x, y):
    doc = globals()['doc']
    page = doc.Pages.Item(1)
    mast = "Документ"
    visio = globals()['visio']
    beginEnd = fun.getTemplateName(visio, mast)
    fun.add(page, beginEnd, x, y, text)