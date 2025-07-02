Attribute VB_Name = "Módulo1"
Sub Sort()
 
    ' Teste
    Dim index As Long
 
    Dim ws As Worksheet
    Dim NDDPrint_Sum As Double
    Dim prodPB_Sum As Long
    Dim prodColor_Sum As Long
    Dim totalProd_Sum As Double
    Dim totalLocacaoCusto As Double
    Dim totalRateio As Double
    
    Dim inicialIndex As Long
    Dim lastIndex As Long
    
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
    
    ' Debug
    MsgBox (Date)
    
    inicialLine = Worksheets("BASE").Range("G3").Value
    
    MsgBox (inicialLine)
    
    ' Abre a planilha do pré faturamento
    Workbooks.Open (ThisWorkbook.Path & "\prefaturamento")
    
    ' Remove o alerta referente à exclusão de uma aba inteira
    Application.DisplayAlerts = False
 
    Worksheets("Resumo").Delete
    
    ' Adiciona novamente o alerta referente à exclusão de uma aba inteira
    
    Application.DisplayAlerts = True
    
    ' Remove o cabeçalho
    Range("A1:A9").EntireRow.Delete
 
    Set ws = Worksheets("Pré-Faturamento")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' acha a última linha com conteúdo na coluna A
 
    With ActiveSheet
        .ListObjects(1).Name = "Table"
    End With
 
    ' Ordena toda a tabela, se baseando na coluna "Série"
    With ActiveWorkbook.Worksheets("Pré-Faturamento").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Table[[#All],[Série]]"), SortOn:= _
        xlSortOnValue, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange Range("A1:AI" & lastRow) ' TO-DO = Alterar esta linha de forma que pegue altomaticamente a quantidade de linhas, pois desta forma está fixo.
        .Header = xlYes
        .Apply
    End With
 
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
 
    'index = InputBox("Informe a linha inicial para começar o rateio:", "Título Teste", "1", 10, 10)
    index = inicialLine
    indexRateio = inicialLine
 
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

        
        ' Valor Total
        valorTotal = (prodPB * ValorUnitPB) + (prodColor * ValorUnitColor) + locacao
        
        ' centro de custo
        cCusto = WorksheetFunction.XLookup(serie, Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("B:B"), Workbooks("04_SIMPRESS - Outsourcing.xlsm").Worksheets("BASE").Range("E:E"))
        
        ' == Atribuições ==
        
        ' MsgBox ("Achei a serie: " + serie)
        
        Call InsertValues(index, filial, dept, equip, serie, prodPB, CDec(ValorUnitPB), prodColor, CDec(ValorUnitColor), CDec(locacao), CDec(valorTotal), cCusto)
        
        ' Range("A" & index & ":M" & index).Interior.Color = RGB(215, 215, 215)
        
        prodPB_Sum = prodPB_Sum + prodPB
        
        prodColor_Sum = prodColor_Sum + prodColor
        
        totalLocacaoCusto = totalLocacaoCusto + CDec(locacao)
        
        totalRateio = totalRateio + CDec(valorTotal)
        
        If serie = "0DKBB07K351PL3" Or serie = "BRCSSD609W" Then
            
            Call InsertValuesCaseDivision(index, prodPB / 2, prodColor / 2, CDec(locacao) / 2, CDec(valorTotal) / 2)
            
            index = index + 1

            Call InsertValues(index, filial, "Produção", "=", "=", prodPB / 2, CDec(ValorUnitPB), prodColor / 2, CDec(ValorUnitColor), CDec(locacao) / 2, CDec(valorTotal) / 2, 1130201)
        
        End If
        
        index = index + 1
        
    Next i

        Call InsertValues(index, "SP", "Financeiro", "NDDIGITAL", "S30960058180115", prodPB, CDec(ValorUnitPB), prodColor, CDec(ValorUnitColor), CDec(locacao) / 2, CDec(NDDPrint_Sum), 1310301)
        
        totalRateio = totalRateio + CDec(NDDPrint_Sum)
        
        index = index + 1
        
        ' Insere após a última linha as informações do NDDPrint
        ' Worksheets("ALI").Range("L" & index).Value = cCusto ' Centro de Custo
        
        ' Debugs
        lastIndex = index - 1
        
        MsgBox (lastIndex)
        
        ' MsgBox (prodPB_Sum)
        ' MsgBox (prodColor_Sum)
        
        totalProd_Sum = prodPB_Sum + prodColor_Sum
        
        ' Altera a cor da linha abaixo da última para cinza
        Range("A" & index & ":M" & index).Interior.Color = RGB(192, 192, 192)
        
        ' Insere o valor total da Produção Preto e Branco
        Range("F" & index) = prodPB_Sum
        
        Range("F" & index).Font.Bold = True
        ' Insere o valor total da Produção Colorido
        
        Range("H" & index) = prodColor_Sum
        Range("H" & index).Font.Bold = True

        ' Insere o valor total de ambas As Produções
        Range("I" & index) = totalProd_Sum
        Range("I" & index).Font.Bold = True
        
        ' Insere o valor total do custo por Locação
        Range("J" & index) = Round(totalLocacaoCusto, 2)
        Range("J" & index).Font.Bold = True
        Range("K" & index) = totalRateio
        Range("K" & index).Font.Bold = True
                
        Worksheets("BASE").Range("G3").Value = lastIndex + 3
                
        lastRowPlan2 = Worksheets("Plan2").Cells(Worksheets("Plan2").Rows.Count, "A").End(xlUp).Row
        
        ' Remove o alerta referente ao fechamento da planilha de pré faturamento
        Application.DisplayAlerts = False
            
        Workbooks("prefaturamento.xlsx").Close savechanges:=False '(or True)
        
        ' Adiciona novamente o alerta referente ao fechamento da planilha de pré faturamento
        Application.DisplayAlerts = True
        
        For i = 4 To lastRowPlan2
        
            sumAux = 0
        
            If Worksheets("Plan2").Range("A" & i).Value = "" Then
                GoTo checkValidation
            End If
            
            centroDeCusto = Worksheets("Plan2").Range("A" & i).Value
            
            For j = indexRateio To lastIndex
            
                If centroDeCusto = Worksheets("ALI").Cells(j, 12) Then
                
                    ' MsgBox (centroDeCusto & " - " & Worksheets("ALI").Cells(j, 12))
                    
                    sumAux = sumAux + Worksheets("ALI").Cells(j, 11)
                    'MsgBox (sumAux)
                
                End If
                    
                Worksheets("Plan2").Cells(i, 2) = sumAux
            
            Next j
                
        Next i

checkValidation:

    MsgBox ("Chegamos ao final sem erros!")
    
End Sub

Sub InsertValues(index As Long, filial As String, departamento As String, equipamento As String, serie As String, prodPB As Long, ValorUnitPB As Double, prodColor As Long, ValorUnitColor As Double, locacao As Double, NDDPrint_Sum As Double, cCusto As Long)

    Worksheets("ALI").Range("A" & index).Value = filial ' Filial
    Worksheets("ALI").Range("B" & index).Value = departamento ' Departamento
    Worksheets("ALI").Range("C" & index).Value = equipamento ' Equip
    Worksheets("ALI").Range("D" & index).Value = serie ' Série
    Worksheets("ALI").Range("E" & index).Value = Date ' Data
    Worksheets("ALI").Range("F" & index).Value = prodPB ' Produção Preto e Branco
    Worksheets("ALI").Range("G" & index).Value = ValorUnitPB ' Valor Unitário Preto e Branco
    Worksheets("ALI").Range("H" & index).Value = prodColor ' Produção Colorido
    Worksheets("ALI").Range("I" & index).Value = ValorUnitColor ' valor Unitário Preto e Branco
    Worksheets("ALI").Range("J" & index).Value = locacao ' Valor Locação
    Worksheets("ALI").Range("K" & index).Value = NDDPrint_Sum ' Valor Total
    Worksheets("ALI").Range("L" & index).Value = cCusto ' Centro de Custo
    Range("A" & index & ":M" & index).Interior.Color = RGB(215, 215, 215)

End Sub

Sub InsertValuesCaseDivision(index As Long, prodPB As Long, prodColor As Long, locacao As Double, idk As Double)

    Worksheets("ALI").Range("F" & index).Value = prodPB ' Produção Preto e Branco
    Worksheets("ALI").Range("H" & index).Value = prodColor ' Produção Colorido
    Worksheets("ALI").Range("J" & index).Value = locacao ' Valor Locação
    Worksheets("ALI").Range("K" & index).Value = idk ' Valor Total
    Range("A" & index & ":M" & index).Interior.Color = RGB(255, 255, 255)
    
End Sub
