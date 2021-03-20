Attribute VB_Name = "Module1"
Sub Кнопка2_Клацніть()
Attribute Кнопка2_Клацніть.VB_Description = "Корегування"
Attribute Кнопка2_Клацніть.VB_ProcData.VB_Invoke_Func = "r\n14"
'
' Кнопка2_Клацніть Макрос
' Корегування
'
' Сполучення клавіш: Ctrl+r
'
Dim arrAct(), arrMms()
arrAct = Array(5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 22, 23, 25, 26, 28, 31, 32, 33, 34, 38, 39)
arrMms = Array(15, 16, 17, 18, 19, 20, 22, 23, 24, 25, 28, 29, 30, 31, 32, 34, 35, 36, 37, 40, 41)

For i = 0 To 20
'--- Копія целі корегування з акту
    Windows("DELTA_month.xlsx").Activate
    valueAct = Cells(arrAct(i), 7).Value

'--- Вставка целі корегування до генератора
    Windows("EAMD_NAEK_month.xlsm").Activate
    Sheets("Коригування").Select
    Range("B3").Value = valueAct

'--- Копія NEK_in
    Sheets("Дані").Select
    Cells(arrMms(i), 8).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy

'--- Вставка NEK_in у генератор
    Sheets("Коригування").Select
    Range("B8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'--- Копія результату корегування
    Range("B11").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Дані").Select

'--- Вставка скорегуваної строки на місце
    Cells(arrMms(i), 8).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Next i

For i = 1 To 744
'--- РАЕС
    Cells(13, 7 + i) = Cells(15, 7 + i) - Cells(16, 7 + i) + Cells(17, 7 + i) - Cells(18, 7 + i) + Cells(19, 7 + i) - Cells(20, 7 + i) - Cells(14, 7 + i)
'--- ЗАЕС
    Cells(21, 7 + i) = Cells(22, 7 + i) - Cells(23, 7 + i) + Cells(24, 7 + i) - Cells(25, 7 + i)
'--- ЮУАЕС
    Cells(27, 7 + i) = Cells(28, 7 + i) - Cells(29, 7 + i) + Cells(30, 7 + i) - Cells(31, 7 + i) + Cells(32, 7 + i) - Cells(34, 7 + i) - Cells(36, 7 + i) + Cells(35, 7 + i) + Cells(37, 7 + i) - Cells(26, 7 + i)
'--- ХАЕС
    Cells(38, 7 + i) = Cells(40, 7 + i) - Cells(41, 7 + i) - Cells(39, 7 + i)
Next i

End Sub


