Attribute VB_Name = "M�dulo1"
Sub Sort()

    Dim NDDPrint_Sum As Double
    
    ' Remove o cabe�alho
    ' Range("A1:A9").EntireRow.Delete
    
    ' Ordena toda a tabela, se baseando na coluna "S�rie"
    With ActiveWorkbook.Worksheets("Pr�-Faturamento").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Table26[[#All],[S�rie]]"), SortOn:= _
        xlSortOnValue, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A1:AI160") ' TO-DO = Alterar esta linha de forma que pegue altomaticamente a quantidade de linhas, pois desta forma est� fixo.
        .Header = xlYes
        .Apply
    End With

    ' Itera sobre cada c�lula
    For Each cell In Worksheets("Pr�-Faturamento").Range("A1:A160").Cells ' TO-DO = Alterar esta linha de forma que pegue altomaticamente a quantidade de linhas, pois desta forma est� fixo.
    
        'Se os 5 primeiros d�gitos da c�lula forem "S3096"
        If Left(cell.Value, 5) = "S3096" Then
            'MsgBox (Worksheets("Pr�-Faturamento").Range("V" & cell.Row))
            ' Soma todos os valores, gerando o total pago pelo software NDDPrint
            NDDPrint_Sum = NDDPrint_Sum + Worksheets("Pr�-Faturamento").Range("V" & cell.Row)
        ElseIf Left(cell.Value, 7) = "TOTAIS:" Then
            ' Remove a coluna "TOTAIS:", pois, ao organizar pela coluna "S�rie", ela fica no meio dos seriais.
            Range("V" & cell.Row).EntireRow.Delete
        End If
    
    Next cell
    
    'MsgBox (Round(NDDPrint_Sum, 2))

    myValue = InputBox("Aqui � a MSG", "T�tulo Teste", "1", 100, 100)

End Sub


