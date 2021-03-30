Attribute VB_Name = "Module1"

Public  Sub Adicionar()
    ' encontra a próxima linha vazia na planilha Extrato
    Dim adc As Worksheet: Set adc = Sheets("Adicionar")
    Dim ext As Worksheet: Set ext = Sheets("Extrato")
    Dim cursor As Range
    Set cursor = ext.Range("A" & Rows.Count).End(xlUp).Offset(1)
    ext.Cells(cursor.Row, 2).Value = adc.Range("B1")
    cursor.Value = adc.Range("B2")
    ext.Cells(cursor.Row, 3).Value = adc.Range("B3")
    If (adc.Range("B4").Value < 0) Then
        ext.Cells(cursor.Row, 5).Value = adc.Range("B4").Value * -1
    Else ext.Cells(cursor.Row, 4).Value = adc.Range("B4").Value
    End If
    If (cursor.Row -1 <= 0) Then 
        ext.Cells(cursor.Row, 6).Value = ext.Cells(cursor.Row, 4).Value - ext.Cells(cursor.Row, 5).Value
    Else
        ext.Cells(cursor.Row, 6).Value = ext.Cells(cursor.Row, 4).Value - ext.Cells(cursor.Row, 5).Value + ext.Cells(cursor.Row - 1, 6)
    End If
End Sub

