Attribute VB_Name = "Module3"
Sub ���_��_������()

' ��������� ������: Ctrl+q


Dim foodPrice As Variant
    foodPrice = InputBox("��������� ���")
Dim foodDescr As Variant
Dim checkBoxState As Variant
    checkBoxState = ActiveSheet.CheckBoxes(1).Value
    If checkBoxState = 1 Then foodDescr = InputBox("�������� �������")
    
   
    Sheets("��� �� ������").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = foodPrice
    Range("C2").Select
    If checkBoxState = 1 Then ActiveCell.Value = foodDescr
    
    
    Sheets("����").Select
    Range("K7").Formula = "=SUM('��� �� ������'!B2:B100)"
    If Range("K7").Value > Range("F11").Value Then Range("K7").Interior.Color = vbRed
    If Range("K7").Value < Range("F11").Value Then Range("K7").Interior.Pattern = xlNone
End Sub

Sub ��������()

Dim foodPrice As Variant
    foodPrice = InputBox("��������� �������")
Dim foodDescr As Variant
Dim checkBoxState As Variant
    checkBoxState = ActiveSheet.CheckBoxes(1).Value
    If checkBoxState = 1 Then foodDescr = InputBox("�������� �������")
    Sheets("��������").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = foodPrice
    Range("C2").Select
    If checkBoxState = 1 Then ActiveCell.Value = foodDescr
    
    
    Sheets("����").Select
    Range("K9").Formula = "=SUM('��������'!B2:B100)"
    If Range("K9").Value > Range("F16").Value Then Range("K9").Interior.Color = vbRed
    If Range("K9").Value < Range("F16").Value Then Range("K9").Interior.Pattern = xlNone
End Sub

Sub ������()


Dim benzPrice As Variant
    benzPrice = InputBox("��������� �������")
Dim benzDescr As Variant
Dim checkBoxState As Variant
    checkBoxState = ActiveSheet.CheckBoxes(1).Value
    If checkBoxState = 1 Then benzDescr = InputBox("�������� �������")
    Sheets("������").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = benzPrice
    Range("C2").Select
    If checkBoxState = 1 Then ActiveCell.Value = benzDescr
    
    
    Sheets("����").Select
    Range("K8").Formula = "=SUM('������'!B2:B100)"
    If Range("K8").Value > Range("F12").Value Then Range("K8").Interior.Color = vbRed
    If Range("K8").Value < Range("F12").Value Then Range("K8").Interior.Pattern = xlNone

End Sub

Sub ������()
'
Dim prPrice As Variant
    prPrice = InputBox("��������� �������")
Dim prDescr As Variant
''Dim checkBoxState As Variant
   '' checkBoxState = ActiveSheet.CheckBoxes(1).Value
    ''If checkBoxState = 1 Then
    prDescr = InputBox("�������� �������")
    
   
    Sheets("������").Select
    Range("A2").Select
    Selection.EntireRow.Insert
    Range("A2").Select
    ActiveCell.Value = Date
    Range("B2").Select
    ActiveCell.Value = prPrice
    Range("C2").Select
    ''If checkBoxState = 1 Then
    ActiveCell.Value = prDescr
    
    Sheets("����").Select
    Range("K10").Formula = "=SUM('������'!B2:B100)"
    If Range("K10").Value > Range("F18").Value Then Range("K10").Interior.Color = vbRed
    If Range("K10").Value < Range("F18").Value Then Range("K10").Interior.Pattern = xlNone
    
End Sub
Sub ������4_������()
'
' ������4_������ ������
'

'
End Sub

Sub ��������()
'
Sheets("��������").Select
Range("A2:C100").Value = ""
Sheets("������").Select
Range("A2:C100").Value = ""

Sheets("������").Select
Range("A2:C100").Value = ""

Sheets("��� �� ������").Select
Range("A2:C100").Value = ""

Sheets("����").Select
Range("K7:K10").Interior.Pattern = xlNone

End Sub

