Attribute VB_Name = "Module1"

Public  Sub Adicionar()
    ' encontra a próxima linha vazia na planilha Extrato
    Dim adc As Worksheet: Set adc = Sheets("Adicionar")
    Dim ext As Worksheet: Set ext = Sheets("Extrato")
    Dim cursor As Range

    Set cursor = ext.Range("A" & Rows.Count).End(xlUp).Offset(1)

    Dim data As Range: Set data = cursor
    Dim descr As Range: Set descr = ext.Cells(cursor.Row, 2)
    Dim catg As Range: Set catg = ext.Cells(cursor.Row, 3)
    Dim renda As Range: Set renda = ext.Cells(cursor.Row, 4)
    Dim despesa As Range: Set despesa = ext.Cells(cursor.Row, 5)
    Dim blncG As Range: Set blncG = ext.Cells(cursor.Row, 6)
    Dim prevBlncG As Range: Set prevBlncG = ext.Cells(cursor.Row - 1, 6)

    descr.Value = adc.Range("B1")
    data.Value = adc.Range("B2")
    catg.Value = adc.Range("B3")
    If (adc.Range("B4").Value < 0) Then
        despesa.Value = adc.Range("B4").Value * -1
        renda.Value = 0
    Else:
        renda.Value = adc.Range("B4").Value
        despesa.Value = 0
    End If
    blncG.Value = 0


    Dim G As Double, R As Double, D As Double, pG As Double
    R = renda.Value
    D = despesa.Value

    If (cursor.Row = 1) Then
        G = R - D
        blncG.Value = G
    Else
        pG = prevBlncG.Value
        G = pG + R - D
        blncG.Value = G
    End If
    adc.Range("B1:B4").Value = ""
End Sub

