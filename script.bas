Attribute VB_Name = "Módulo1"
Sub Sort()

    Dim ws As Worksheet
    Dim NDDPrint_Sum As Double
    
    Workbooks.Open (ThisWorkbook.Path & "\prefaturamento")
    
    Application.DisplayAlerts = False

    Worksheets("Resumo").Delete
    
    Application.DisplayAlerts = True
    
    ' Remove o cabeçalho
    Range("A1:A9").EntireRow.Delete

    ' Ordena toda a tabela, se baseando na coluna "Série"
    With ActiveWorkbook.Worksheets("Pré-Faturamento").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Table2[[#All],[Série]]"), SortOn:= _
        xlSortOnValue, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A1:AI160") ' TO-DO = Alterar esta linha de forma que pegue altomaticamente a quantidade de linhas, pois desta forma está fixo.
        .Header = xlYes
        .Apply
    End With

    Set ws = Worksheets("Pré-Faturamento")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' acha a última linha com conteúdo na coluna A

    ' Itera sobre cada célula
    'For Each cell In Worksheets("Pré-Faturamento").Range("A1:A160").Cells ' TO-DO = Alterar esta linha de forma que pegue altomaticamente a quantidade de linhas, pois desta forma está fixo.
    For i = lastRow To 1 Step -1 ' de baixo para cima
    
        'Se os 5 primeiros dígitos da célula forem "S3096"
        If Left(ws.Cells(i, 1).Value, 5) = "S3096" Or Left(ws.Cells(i, 1).Value, 5) = "S0000" Then
        
            ' Soma todos os valores, gerando o total pago pelo software NDDPrint
            NDDPrint_Sum = NDDPrint_Sum + ws.Cells(i, 22).Value
            
            ws.Rows(i).Delete
        
        ElseIf Left(ws.Cells(i, 1).Value, 7) = "TOTAIS:" Then
            ' Remove a coluna "TOTAIS:", pois, ao organizar pela coluna "Série", ela fica no meio dos seriais.
            ws.Rows(i).Delete
        End If
    
    Next i
    
    MsgBox (Round(NDDPrint_Sum, 2))

    index = InputBox("Informe a linha inicial para começar o rateio:", "Título Teste", "1", 10, 10)

    Workbooks.Open (ThisWorkbook.Path & "\04_SIMPRESS - Outsourcing.xlsm")
    
    Worksheets("ALI").Range("A" & index).Value = "Teste"
    
    
End Sub

