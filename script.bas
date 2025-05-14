Attribute VB_Name = "Módulo1"
Sub Sort()

    Dim ws As Worksheet
    Dim NDDPrint_Sum As Double
    Dim prodPB_Sum As Long
    Dim prodColor_Sum As Long
    
    ' Variáveis para o rateio
    Dim filial As String
    Dim dept As String
    Dim equip As String
    Dim serie As String
        
    Dim prodPB As Long
    Dim ValorUnitPB As Double
        
    Dim prodColor As Long
    Dim ValorUnitColor As Double
    Dim locacao As Double
    Dim valorTotal As Double
    
    Dim cCusto As Long
    
    Dim dateValue As Variant
    
    ' Debug
    MsgBox (Date)

    ' Debug
    MsgBox (dateValue)
    
    ' Abre a planilha do pré faturamento
    Workbooks.Open (ThisWorkbook.Path & "\prefaturamento")
    
    ' Remove o alerta referente à exclusão de uma aba inteira
    Application.DisplayAlerts = False

    Worksheets("Resumo").Delete
    
    ' Adiciona novamente o alerta referente à exclusão de uma aba inteira
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
    
    'MsgBox (Round(NDDPrint_Sum, 2))

    'Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Copy _
    'Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento")

    index = InputBox("Informe a linha inicial para começar o rateio:", "Título Teste", "1", 10, 10)

    ' Abre a planilha principal
    Workbooks.Open (ThisWorkbook.Path & "\04_SIMPRESS - Outsourcing.xlsm")
    
    ' Teste
    
    MsgBox ("Breakpoint")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' acha a última linha com conteúdo na coluna A
    
    For i = 2 To lastRow Step 1 ' de cima para baixo
        
        ' Série
        serie = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("A" & i).Value
        
        ' Equipamento
        equip = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("D" & i).Value
        
        ' Filial
        filial = WorksheetFunction.XLookup(serie, Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("B:B"), Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("C:C"))
        
        ' Departamento
        dept = WorksheetFunction.XLookup(serie, Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("B:B"), Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("D:D"))
        
        ' Produção Preto e Branco
        prodPB = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("M" & i).Value
        
        ' Valor Unitário Preto e Branco
        ValorUnitPB = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("O" & i).Value
        
        ' Produção Colorido
        prodColor = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("N" & i).Value
        
        ' valor Unitário Preto e Branco
        ValorUnitColor = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("P" & i).Value
        
        ' Valor Locação
        locacao = Workbooks("prefaturamento.xlsx").Worksheets("Pré-Faturamento").Range("T" & i).Value
        
        ' Valor Total =
        valorTotal = (prodPB * ValorUnitPB) + (prodColor * ValorUnitColor) + locacao
        
        ' centro de custo
        cCusto = WorksheetFunction.XLookup(serie, Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("B:B"), Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("E:E"))
        
        ' == Atribuições ==
        
        Worksheets("ALI").Range("A" & index).Value = filial ' Filial
        Worksheets("ALI").Range("B" & index).Value = dept ' Departamento
        
        Worksheets("ALI").Range("C" & index).Value = equip ' Equip
        Worksheets("ALI").Range("D" & index).Value = serie ' Série
        
        Worksheets("ALI").Range("E" & index).Value = Date ' Data
        
        Worksheets("ALI").Range("F" & index).Value = prodPB ' Produção Preto e Branco
        Worksheets("ALI").Range("G" & index).Value = Format(ValorUnitPB, "#,##0.0000") ' Valor Unitário Preto e Branco
        Worksheets("ALI").Range("H" & index).Value = prodColor ' Produção Colorido
        Worksheets("ALI").Range("I" & index).Value = Format(ValorUnitColor, "#,##0.0000") ' valor Unitário Preto e Branco
        Worksheets("ALI").Range("J" & index).Value = Format(locacao, "#,##0.00") ' Valor Locação
        
        Worksheets("ALI").Range("K" & index).Value = Format(valorTotal, "#,##0.00") ' Valor Total
        
        Worksheets("ALI").Range("L" & index).Value = cCusto ' Centro de Custo
        
        prodPB_Sum = prodPB_Sum + prodPB
        prodColor_Sum = prodColor_Sum + prodColor
        
        index = index + 1
    
    Next i
    
        ' colorRange = ""
        Range("A" & index & ":L" & index).Interior.Color = RGB(192, 192, 192)
        
        ' Debugs
        MsgBox (prodPB_Sum)
        MsgBox (prodColor_Sum)
        
End Sub
