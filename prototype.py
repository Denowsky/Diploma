import tkinter as tk
from tkinter import ttk
from datetime import datetime
import pandas as pd

# Функции
def read_file(fileName):
    with open (f"{fileName}.txt", "r", encoding="utf-8") as f:
        data = {}
        product = []
        for line in f:
            key, value = line.strip().split(": ")
            data[key] = value
        for key in data:
            product.append(key)
    f.close
    return product, data

def on_select(value, index):
    if workFrame.grid_size()[1]<=7:
        workFrame.addRowButton.config(state="normal")
    info = option_menus[index].grid_info()
    workFrame.grid_slaves(info["row"],info["column"]+1)[0].destroy()
    #
    data = {}
    parts = []
    data = read_file("file")[1]
    parts = data[value].split(", ")
    #
    var = tk.StringVar(workFrame)
    var.set("Выберите деталь")
    #
    workFrame.partMenu = tk.OptionMenu(workFrame, var, *parts)
    workFrame.partMenu.config(width=30)
    #
    workFrame.partMenu.grid(row=info["row"],column=info["column"]+1)
    
def add_rows():
    #
    workFrame.addRowButton.config(state="disabled")
    workFrame.delRowButton.config(state="norm")
    #
    rowsNum = workFrame.grid_size()[1]
    #
    workFrame.products = read_file("file")[0]
    #
    var = tk.StringVar(workFrame)
    var.set("Выберите изделие")
    #
    workFrame.productMenu = tk.OptionMenu(workFrame, var, *workFrame.products, command=lambda value, 
                                          index=len(option_menus): on_select(value, index))
    workFrame.productMenu.config(width=20)
    option_menus.append(workFrame.productMenu)
    #
    workFrame.productMenu.grid(row=rowsNum, column=0)

    #
    workFrame.partLable = tk.Label(workFrame,text="Выберите изделие")
    #
    workFrame.partLable.grid(row=rowsNum, column=1)

    # Создание полей ввода
    workFrame.lengthEntry = tk.Entry(workFrame)
    workFrame.widthEntry = tk.Entry(workFrame)
    workFrame.thicknessEntry = tk.Entry(workFrame)
    workFrame.planEntry = tk.Entry(workFrame)
    workFrame.okEntry = tk.Entry(workFrame)
    workFrame.defectEntry = tk.Entry(workFrame)
    # Позиционирование элементов
    workFrame.productMenu.grid(row=rowsNum, column=0)
    workFrame.lengthEntry.grid(row=rowsNum, column=2, padx=6, pady=6)
    workFrame.widthEntry.grid(row=rowsNum, column=3, padx=6, pady=6)
    workFrame.thicknessEntry.grid(row=rowsNum, column=4, padx=6, pady=6)
    workFrame.planEntry.grid(row=rowsNum, column=5, padx=6, pady=6)
    workFrame.okEntry.grid(row=rowsNum, column=6, padx=6, pady=6)
    workFrame.defectEntry.grid(row=rowsNum, column=7, padx=6, pady=6)


def delete_rows():
    rowsNum = workFrame.grid_size()[1]-1
    if rowsNum>1:
        for j in range(0, workFrame.grid_size()[0]):
            workFrame.grid_slaves(rowsNum,j)[0].destroy()
    else:
        workFrame.addRowButton.config(state="normal")
        workFrame.delRowButton.config(state="disabled")
        return

def makeReport():
    date = []
    shift = []
    equipment = []
    operator = []
    stuff1 = [date, shift, equipment, operator]
    product = []
    part = []
    stuff2 = [product,part]
    width = []
    length = []
    thickness = []
    planPcs = []
    goodPcs = []
    defectPcs = []
    column = 0
    stuff3 = [length,width,thickness,planPcs,goodPcs,defectPcs]
    row = workFrame.grid_size()[1]
    for list in stuff1:
        for i in range(2,row):
            list.append(infoFrame.grid_slaves(0,column)[0]["text"])
        column+=1
    column = 0
    for list in stuff2:
        for i in range (2, row):
            list.append(workFrame.grid_slaves(i,column)[0]["text"])
        column+=1
    for list in stuff3:
        for i in range (2, row):
            list.append(workFrame.grid_slaves(i,column)[0].get())
        column+=1
    data = pd.DataFrame({
        'Дата': date,
        'Смена': shift,
        'Оборудование': equipment,
        'Оператор': operator,
        'Изделие': product,
        'Деталь': part,
        'Длина': length,
        'Ширина': width,
        'Толщина': thickness,
        'План шт.': planPcs,
        'Ок шт.': goodPcs,
        'Брак шт.': defectPcs
    })
    data.to_excel('./work.xlsx', sheet_name='mounth', index=False)
    



def makeEstimation():
    timeAndperformance=read_file("workTime")[1]
    shiftTime=int(timeAndperformance["Продолжительность смены в минутах"])
    planTime=int(timeAndperformance["Планируемое время в минутах"])
    planInpact=int(timeAndperformance["Производительность шт/ч"])
    result = []
    sum = 0
    ok = 0
    defect = 0
    quality = 0
    performance = 0
    availability = 0
    currentInpact = 0
    try:
        for pos in range(2,workFrame.grid_size()[1]):
            ok+=int(workFrame.grid_slaves(pos, 6)[0].get())
            defect+=int(workFrame.grid_slaves(pos, 7)[0].get())
        sum = ok+defect
        if sum>0:
            quality = ok/sum
            currentInpact = sum/(shiftTime/60)
        performance = currentInpact/planInpact
        availability = shiftTime/planTime
        oee = quality*performance*availability
        result.append(round(quality*100))
        result.append(round(performance*100))
        result.append(round(availability*100))
        result.append(round(oee*100))
        qualityVar.set(f'Качество: {result[0]}%')
        performanceVar.set(f'Производительность: {result[1]}%')
        availabilityVar.set(f'Доступность: {result[2]}%')
        oeeVar.set(f'Общая Эффективность Оборудования: {result[3]}%')
        timeVar.set(f'Время подсчёта: {str(datetime.now().time()).split(".")[0]}')
        statusVar.set(f'Статус: все данные занесены, подсчёт произведён')
        reportFrame.reportButton.config(state="norm")
    except:
        statusVar.set(f'Статус: ошибка в занесении данных, не все данные занесены')
        reportFrame.reportButton.config(state="disabled")
        return
    

#Tkinter
root = tk.Tk()     # создаем корневой объект - окно
root.title("Приложение на Tkinter")     # устанавливаем заголовок окна
root.geometry("1366x768")    # устанавливаем размеры окна 

#
logo = tk.PhotoImage(file="./icon.png")
#
root.logoLabel = ttk.Label(image=logo)
root.text1Label = tk.Label(text="Основная информация:", font=("Arial",18))
root.text2Label = tk.Label(text="Производственные данные:", font=("Arial",18))
root.text3Label = tk.Label(text="Производственные показатели:", font=("Arial",18))
#
root.logoLabel.grid(row=0,column=0, sticky="n")
root.text1Label.grid(row=1,column=0,pady=20, sticky="w")
root.text2Label.grid(row=3,column=0,pady=20, sticky="w")
root.text3Label.grid(row=5,column=0,pady=20, sticky="w")
#
option_menus = [] #

##infoFrame
infoFrame = tk.Frame(root)
infoFrame.grid(row=2, column=0, sticky="W")
# Создаём окно даты
infoFrame.dateAndTime = tk.Label(infoFrame,text=datetime.now().date())
# Размещаем дату в пространстве infoFrame
infoFrame.dateAndTime.grid(row=0,column=0)

# создаем 3 варианта для выпадающих списков
equipment = read_file("equipment")[1]
infoFrame.operators = equipment["Операторы"].split(", ")
infoFrame.machines = equipment["Оборудование"].split(", ")
infoFrame.shifts = equipment["Сменность"].split(", ")
var1 = tk.StringVar(infoFrame)
var1.set(infoFrame.operators[0])
var2 = tk.StringVar(infoFrame)
var2.set(infoFrame.machines[0])
var3 = tk.StringVar(infoFrame)
var3.set(infoFrame.shifts[0])

# Создаем выпадающие списки
infoFrame.operatorMenu = ttk.OptionMenu(infoFrame, var1, *infoFrame.operators)
infoFrame.operatorMenu.config(width=25)
infoFrame.machineMenu = ttk.OptionMenu(infoFrame, var2, *infoFrame.machines)
infoFrame.machineMenu.config(width=25)
infoFrame.shiftMenu = ttk.OptionMenu(infoFrame, var3, *infoFrame.shifts)
infoFrame.shiftMenu.config(width=10)
# Размещаем выпадающие списки в пространстве infoFrame
infoFrame.shiftMenu.grid(row=0,column=1)
infoFrame.machineMenu.grid(row=0,column=2)
infoFrame.operatorMenu.grid(row=0,column=3)


##workFrame
workFrame = tk.Frame(root) 
workFrame.grid(row=4,column=0,sticky="W")

# создаем поясняющие надписи
workFrame.productLabel = tk.Label(workFrame, text="Изделие")
workFrame.partLabel = tk.Label(workFrame, text="Деталь")
workFrame.lengthLabel = tk.Label(workFrame, text="Длина")
workFrame.widthLabel = tk.Label(workFrame, text="Ширина")
workFrame.thicknessLabel = tk.Label(workFrame, text="Толщина")
workFrame.planLabel = tk.Label(workFrame, text="План шт.")
workFrame.okLabel = tk.Label(workFrame, text="Сделано Хороших шт.")
workFrame.defectLabel = tk.Label(workFrame, text="Отправлено в брак шт.")
# размещаем поясняющие надписи в ряд
workFrame.productLabel.grid(row=1, column=0)
workFrame.partLabel.grid(row=1, column=1)
workFrame.lengthLabel.grid(row=1, column=2, padx=6, pady=6)
workFrame.widthLabel.grid(row=1, column=3, padx=6, pady=6)
workFrame.thicknessLabel.grid(row=1, column=4, padx=6, pady=6)
workFrame.planLabel.grid(row=1, column=5, padx=6, pady=6)
workFrame.okLabel.grid(row=1, column=6, padx=6, pady=6)
workFrame.defectLabel.grid(row=1, column=7, padx=6, pady=6)

# Создаём кнопки
workFrame.addRowButton = ttk.Button(workFrame, text="Добавить строку", command=add_rows)
workFrame.delRowButton = ttk.Button(workFrame, text="Удалить строку", command=delete_rows)
# Размещаем кнопоки в пространстве infoFrame
workFrame.addRowButton.grid(row=0, column=0, padx = 20)
workFrame.delRowButton.grid(row=0, column=1, padx=20)
#
reportFrame = tk.Frame(root)
reportFrame.grid(row=6,column=0, sticky="W")
#
qualityVar = tk.StringVar(reportFrame, value=f'Качество: 0%')
performanceVar = tk.StringVar(reportFrame, value=f'Производительность: 0%')
availabilityVar = tk.StringVar(reportFrame, value=f'Доступность: 0%')
oeeVar = tk.StringVar(reportFrame, value=f'Общая Эффективность Оборудования: 0%')
timeVar = tk.StringVar(reportFrame, value=f'Время подсчёта: ')
statusVar = tk.StringVar(reportFrame, value=f'Статус: подсчёт не произведён')
#
reportFrame.qualityLabel = tk.Label(reportFrame, textvariable=qualityVar)
reportFrame.performanceLabel = tk.Label(reportFrame, textvariable=performanceVar)
reportFrame.availabilityLabel = tk.Label(reportFrame, textvariable=availabilityVar)
reportFrame.oeeLabel = tk.Label(reportFrame, textvariable=oeeVar)
reportFrame.timeLabel = tk.Label(reportFrame, textvariable=timeVar)
reportFrame.statusLabel = tk.Label(reportFrame, textvariable=statusVar)
#
reportFrame.qualityLabel.grid(row=0,column=0, padx=6, pady=6)
reportFrame.performanceLabel.grid(row=0,column=1, padx=6, pady=6)
reportFrame.availabilityLabel.grid(row=0,column=2, padx=6, pady=6)
reportFrame.oeeLabel.grid(row=0,column=3, padx=6, pady=6)
reportFrame.timeLabel.grid(row=1,column=0, columnspan=2, pady=20)
reportFrame.statusLabel.grid(row=1,column=3, columnspan=3)
#
reportFrame.estimateButton = ttk.Button(reportFrame, text="Подсчитать", command=makeEstimation)
reportFrame.reportButton = ttk.Button(reportFrame, text="Выгрузить в excel", command=makeReport, 
                                      state="disabled")
#
reportFrame.estimateButton.grid(row=1,column=2)
reportFrame.reportButton.grid(row=1,column=6)

add_rows()

# Запускаем главный цикл обработки событий
root.mainloop()