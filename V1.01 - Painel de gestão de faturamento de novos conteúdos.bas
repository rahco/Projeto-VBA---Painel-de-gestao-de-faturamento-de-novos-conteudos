Attribute VB_Name = "Módulo1"
Sub Geral()

    Application.ScreenUpdating = False

    Call BD_Cre_2
    Call BD_bv_completa_2
    Call BV_completa
    Call Base_de_vendas_completa
    Call BD_resultados
    Call Base_de_resultados

    Sheets("MACROS").Select
    Range("B7").Select

    Application.ScreenUpdating = True

End Sub

Sub BD_Cre_2()
Attribute BD_Cre_2.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False

    Sheets("BD - CRE").Select
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BD - CRE (2)").Select
    Range("B2").Select
    ActiveSheet.Paste
    Range("B5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range(Selection, Selection.End(xlDown)).Select
    Range("B3").Select
    ActiveWorkbook.Worksheets("BD - CRE (2)").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BD - CRE (2)").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("E2:E100000"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BD - CRE (2)").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("BD - CRE (2)").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BD - CRE (2)").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("H2:H100000"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BD - CRE (2)").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range("$B$2:$V$100000").RemoveDuplicates Columns:=11, Header:= _
        xlYes
    Range("L3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("L3"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar _
        :="-", FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True
    Range("B3").Select
    
    Application.ScreenUpdating = True

End Sub


Sub BD_bv_completa_2()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BD - BV COMPLETA (2)").Range("C4").Value)
    final = Abs(Worksheets("BD - BV COMPLETA (2)").Range("B4").Value)
 
    Do While atual > final
        Sheets("BD - BV COMPLETA (2)").Select
        Range("B6").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B6").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BD - BV COMPLETA (2)").Range("C4").Value)
        final = Abs(Worksheets("BD - BV COMPLETA (2)").Range("B4").Value)
    Loop

    Sheets("BD - BV COMPLETA (2)").Select
    Range("B6").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C4").Value > 0 Then
        linhaf = linhai - Range("C4").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C4").Value < 0 Then
        linhaf = linhai + Range("C4").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B6").Select

    Sheets("BD - BV COMPLETA").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BD - BV COMPLETA (2)").Select
    Range("B6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$B$5:$E$50000").RemoveDuplicates Columns:=1, Header:= _
        xlYes
    Range("C2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B6").Select

    Application.ScreenUpdating = True

End Sub

Sub BV_completa()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BV COMPLETA").Range("C2").Value)
    final = Abs(Worksheets("BV COMPLETA").Range("B2").Value)
 
    Do While atual > final
        Sheets("BV COMPLETA").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BV COMPLETA").Range("C2").Value)
        final = Abs(Worksheets("BV COMPLETA").Range("B2").Value)
    Loop

    Sheets("BV COMPLETA").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select

    Sheets("BD - BV COMPLETA").Select
    Range("B6").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("B6").Select
    Sheets("BV COMPLETA").Select
    Range("B4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Range("H4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("H5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H4").Select
    Application.CutCopyMode = False
    Range("B4").Select

    Application.ScreenUpdating = True

End Sub

Sub Base_de_vendas_completa()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE DE VENDAS COMPLETA").Range("C1").Value)
    final = Abs(Worksheets("BASE DE VENDAS COMPLETA").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE DE VENDAS COMPLETA").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE DE VENDAS COMPLETA").Range("C1").Value)
        final = Abs(Worksheets("BASE DE VENDAS COMPLETA").Range("B1").Value)
    Loop

    Sheets("BASE DE VENDAS COMPLETA").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select

    Sheets("BV COMPLETA").Select
    Range("R3").Select
    ActiveSheet.Range("$B$3:$R$50000").AutoFilter Field:=17, Criteria1:="=1", _
        Operator:=xlAnd
    Range("H3:O3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE DE VENDAS COMPLETA").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BV COMPLETA").Select
    Range("R3").Select
    ActiveSheet.Range("$B$3:$R$50000").AutoFilter Field:=17
    Range("B4").Select
    Sheets("BASE DE VENDAS COMPLETA").Select
    Range("B4").Select
    ActiveWorkbook.RefreshAll

    Application.ScreenUpdating = True

End Sub

Sub BD_resultados()

    Application.ScreenUpdating = False

    'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BD - RESULTADOS").Range("C2").Value)
    final = Abs(Worksheets("BD - RESULTADOS").Range("B2").Value)
 
    Do While atual > final
        Sheets("BD - RESULTADOS").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BD - RESULTADOS").Range("C2").Value)
        final = Abs(Worksheets("BD - RESULTADOS").Range("B2").Value)
    Loop

    Sheets("BD - RESULTADOS").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C2").Value > 0 Then
        linhaf = linhai - Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C2").Value < 0 Then
        linhaf = linhai + Range("C2").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select

    Range("B4:C4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B4").Select
    Sheets("BD - BV COMPLETA (2)").Select
    Range("E5").Select
    ActiveSheet.Range("$B$5:$E$50000").AutoFilter Field:=4, Criteria1:="=0", _
        Operator:=xlAnd
    Range("B5:C5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BD - RESULTADOS").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Sheets("BD - BV COMPLETA (2)").Select
    Application.CutCopyMode = False
    Range("E5").Select
    ActiveSheet.Range("$B$5:$E$50000").AutoFilter Field:=4
    Range("B6").Select
    Sheets("BASE INICIAL").Select
    Range("B6:C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BD - RESULTADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Sheets("BASE INICIAL").Select
    Application.CutCopyMode = False
    Range("B6").Select
    Sheets("BD - RESULTADOS").Select
    Range("B4").Select
    Range("D4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("D5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("D4").Select
    Application.CutCopyMode = False
    Range("B4").Select
    ActiveWorkbook.RefreshAll

    Application.ScreenUpdating = True

End Sub

Sub Base_de_resultados()

    Application.ScreenUpdating = False
    
     'Tipo Var
    Dim atual As Double
    Dim final As Double
    Dim linhai As Double
    Dim linhaf As Double
    
    atual = Abs(Worksheets("BASE DE RESULTADOS").Range("C1").Value)
    final = Abs(Worksheets("BASE DE RESULTADOS").Range("B1").Value)
 
    Do While atual > final
        Sheets("BASE DE RESULTADOS").Select
        Range("B4").Select
        Selection.End(xlDown).Select
        linhai = ActiveCell.Row - 1
        linhaf = Range("B4").Row + 2
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
        atual = Abs(Worksheets("BASE DE RESULTADOS").Range("C1").Value)
        final = Abs(Worksheets("BASE DE RESULTADOS").Range("B1").Value)
    Loop

    Sheets("BASE DE RESULTADOS").Select
    Range("B4").Select
    Selection.End(xlDown).Select
    linhai = ActiveCell.Row - 1
    
    If Range("C1").Value > 0 Then
        linhaf = linhai - Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Copy
        Selection.Insert Shift:=xlDown
    ElseIf Range("C1").Value < 0 Then
        linhaf = linhai + Range("C1").Value + 1
        Rows(linhaf & ":" & linhai).Select
        Selection.Delete Shift:=xlUp
    Else
    End If
  
    Application.CutCopyMode = False
    Range("B4").Select
    
    Sheets("BD - RESULTADOS").Select
    Range("K3").Select
    ActiveSheet.Range("$B$3:$K$50000").AutoFilter Field:=10, Criteria1:="=1", _
        Operator:=xlAnd
    Range("D3:J3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("BASE DE RESULTADOS").Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B4").Select
    Application.CutCopyMode = False
    Sheets("BD - RESULTADOS").Select
    Range("K3").Select
    ActiveSheet.Range("$B$3:$K$50000").AutoFilter Field:=10
    Range("B4").Select
    Sheets("BASE DE RESULTADOS").Select
    Range("H3").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("H3"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("D3").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("D3"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E3").Select
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort.SortFields. _
        Add2 Key:=Range("E3"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BASE DE RESULTADOS").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B4").Select
    ActiveWorkbook.RefreshAll
      
    Application.ScreenUpdating = True

End Sub

Sub Arquivo_de_envio()

    Application.ScreenUpdating = False
    
    ActiveWorkbook.Save
    
    ActiveWorkbook.SaveAs Filename:= _
        ActiveWorkbook.Path & "\" & Worksheets("MACROS").Range("C14").Value & " - Gestão de Faturamento Novos Conteúdos - Dados até dia " & Worksheets("MACROS").Range("C15").Value & ".xlsm" _
        , FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False

    Sheets("QUADRO DE PERFORMANCE").Select
    Cells.Select
    Range("A2").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B6").Select
    Application.CutCopyMode = False
    ActiveWindow.DisplayHeadings = False
    Sheets("PERFORMANCE POR FORNECEDOR").Select
    ActiveSheet.Shapes.Range(Array("Supervisor RMV 2")).Select
    Range("B5").Select
    Sheets("TD").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("BASE DE RESULTADOS").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1:C1").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    Sheets("BASE DE VENDAS COMPLETA").Select
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B1:C1").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B4").Select
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.ScrollWorkbookTabs Sheets:=1
    Sheets(Array("TD - BV COMPLETA", "GRÁFICO DE ENVIO")).Select
    Sheets("GRÁFICO DE ENVIO").Activate
    Sheets(Array("BD - CRE (2)", "BD - BV COMPLETA (2)", "BV COMPLETA", _
        "BD - RESULTADOS", "TD - BV COMPLETA", "GRÁFICO DE ENVIO")).Select
    Sheets("GRÁFICO DE ENVIO").Activate
    Sheets(Array("MACROS", "BASE DIAS", "BD - BV COMPLETA", "BD - CRE", "BD - NOME F.", _
        "BASE INICIAL", "BD - CRE (2)", "BD - BV COMPLETA (2)", "BV COMPLETA", _
        "BD - RESULTADOS", "TD - BV COMPLETA", "GRÁFICO DE ENVIO")).Select
    Sheets("BD - CRE").Activate
    Sheets("HC").Visible = True
    Sheets("HC").Select
    Sheets("METAS").Visible = True
    Sheets("METAS").Select
    Sheets("TD").Visible = True
    Sheets("TD").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets(Array("TD - BV COMPLETA", "GRÁFICO DE ENVIO")).Select
    Sheets("GRÁFICO DE ENVIO").Activate
    Sheets(Array("BD - CRE (2)", "BD - BV COMPLETA (2)", "BV COMPLETA", _
        "BD - RESULTADOS", "TD - BV COMPLETA", "GRÁFICO DE ENVIO")).Select
    Sheets("GRÁFICO DE ENVIO").Activate
    Sheets(Array("HC", "METAS", "MACROS", "BASE DIAS", "BD - BV COMPLETA", "BD - CRE", _
        "BD - NOME F.", "BASE INICIAL", "BD - CRE (2)", "BD - BV COMPLETA (2)", _
        "BV COMPLETA", "BD - RESULTADOS", "TD - BV COMPLETA", "GRÁFICO DE ENVIO")).Select
    Sheets("HC").Activate
    ActiveWindow.SelectedSheets.Delete
    Sheets("QUADRO DE PERFORMANCE").Select
    ActiveWorkbook.Save

    Application.ScreenUpdating = True

End Sub
