Attribute VB_Name = "Module3"
Sub Еда_на_работе()

' Сочетание клавиш: Ctrl+q


Dim foodPrice As Variant
    foodPrice = InputBox("Стоимость еды")
Dim foodDescr As Variant
Dim checkBoxState As Variant
    checkBoxState = ActiveSheet.CheckBoxes(1).Value
    If checkBoxState = 1 Then foodDescr = InputBox("Описание покупки")
    
   
    Sheets("Еда на работе").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = foodPrice
    Range("C2").Select
    If checkBoxState = 1 Then ActiveCell.Value = foodDescr
    
    
    Sheets("Март").Select
    Range("K7").Formula = "=SUM('Еда на работе'!B2:B100)"
    If Range("K7").Value > Range("F11").Value Then Range("K7").Interior.Color = vbRed
    If Range("K7").Value < Range("F11").Value Then Range("K7").Interior.Pattern = xlNone
End Sub

Sub Продукты()

Dim foodPrice As Variant
    foodPrice = InputBox("Стоимость покупки")
Dim foodDescr As Variant
Dim checkBoxState As Variant
    checkBoxState = ActiveSheet.CheckBoxes(1).Value
    If checkBoxState = 1 Then foodDescr = InputBox("Описание покупки")
    Sheets("Продукты").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = foodPrice
    Range("C2").Select
    If checkBoxState = 1 Then ActiveCell.Value = foodDescr
    
    
    Sheets("Март").Select
    Range("K9").Formula = "=SUM('Продукты'!B2:B100)"
    If Range("K9").Value > Range("F16").Value Then Range("K9").Interior.Color = vbRed
    If Range("K9").Value < Range("F16").Value Then Range("K9").Interior.Pattern = xlNone
End Sub

Sub Бензин()


Dim benzPrice As Variant
    benzPrice = InputBox("Стоимость бензина")
Dim benzDescr As Variant
Dim checkBoxState As Variant
    checkBoxState = ActiveSheet.CheckBoxes(1).Value
    If checkBoxState = 1 Then benzDescr = InputBox("Описание покупки")
    Sheets("Бензин").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = benzPrice
    Range("C2").Select
    If checkBoxState = 1 Then ActiveCell.Value = benzDescr
    
    
    Sheets("Март").Select
    Range("K8").Formula = "=SUM('Бензин'!B2:B100)"
    If Range("K8").Value > Range("F12").Value Then Range("K8").Interior.Color = vbRed
    If Range("K8").Value < Range("F12").Value Then Range("K8").Interior.Pattern = xlNone

End Sub

Sub Прочее()
'
Dim prPrice As Variant
    prPrice = InputBox("Стоимость покупки")
Dim prDescr As Variant
''Dim checkBoxState As Variant
   '' checkBoxState = ActiveSheet.CheckBoxes(1).Value
    ''If checkBoxState = 1 Then
    prDescr = InputBox("Описание покупки")
    
   
    Sheets("Прочее").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = prPrice
    Range("C2").Select
    ''If checkBoxState = 1 Then
    ActiveCell.Value = prDescr
    
    Sheets("Март").Select
    Range("K10").Formula = "=SUM('Прочее'!B2:B100)"
    If Range("K10").Value > Range("F18").Value Then Range("K10").Interior.Color = vbRed
    If Range("K10").Value < Range("F18").Value Then Range("K10").Interior.Pattern = xlNone
    
End Sub
Sub Флажок4_Щелчок()
'
' Флажок4_Щелчок Макрос
'

'
End Sub

Sub очистить()
'
Sheets("Продукты").Select
Range("A2:C100").Value = ""
Sheets("Прочее").Select
Range("A2:C100").Value = ""

Sheets("Бензин").Select
Range("A2:C100").Value = ""

Sheets("Еда на работе").Select
Range("A2:C100").Value = ""

Sheets("Март").Select
Range("K7:K10").Interior.Pattern = xlNone

End Sub

