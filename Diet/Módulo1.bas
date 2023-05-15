Attribute VB_Name = "Módulo1"
Sub DietaUFF()

Application.ScreenUpdating = False
With Application
.Calculation = xlManual
.MaxChange = 0.001
End With
ActiveWorkbook.PrecisionAsDisplayed = False

Dim Linha, UltLinDieta, UltLinPlan, z, Total As Long
Dim perc As Double

Worksheets("Planejamento").Select
Application.Cursor = xlWait

If Worksheets("Planejamento").Cells(4, 1).Value <> "" Then

UltiLinDieta = 2
While Worksheets("TabelaUFF").Cells(UltiLinDieta, 1).Value <> ""
UltiLinDieta = UltiLinDieta + 1
Wend
'UltiLinDieta = UltiLinDieta - 1

UltLinPlan = 4
While Worksheets("Planejamento").Cells(UltLinPlan, 1).Value <> ""
UltLinPlan = UltLinPlan + 1
Wend
UltLinPlan = UltLinPlan - 1

If Worksheets("Planejamento").Cells(UltLinPlan, 1).Value = "%" Then UltLinPlan = UltLinPlan - 1
If Worksheets("Planejamento").Cells(UltLinPlan, 1).Value = "Kcal:" Then UltLinPlan = UltLinPlan - 1
If Worksheets("Planejamento").Cells(UltLinPlan, 1).Value = "Totais:" Then UltLinPlan = UltLinPlan - 1

For i = 4 To UltLinPlan

'Worksheets("Planejamento").Rows(i).Select
Worksheets("Planejamento").Cells(i, 2).NumberFormat = "###0"
Worksheets("Planejamento").Cells(i, 2).Borders(xlEdgeLeft).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 2).Borders(xlEdgeTop).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 2).Borders(xlEdgeBottom).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 2).Borders(xlEdgeRight).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 2).Borders(xlInsideVertical).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 2).Borders(xlInsideHorizontal).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 1).Borders(xlEdgeLeft).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 1).Borders(xlEdgeTop).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 1).Borders(xlEdgeBottom).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 1).Borders(xlEdgeRight).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 1).Borders(xlInsideVertical).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, 1).Borders(xlInsideHorizontal).LineStyle = xlNone


Linha = 2
While Worksheets("Planejamento").Cells(i, 1).Value <> Worksheets("TabelaUFF").Cells(Linha, 1).Value And Worksheets("TabelaUFF").Cells(Linha, 1).Value <> ""
Linha = Linha + 1
Wend

If Linha = UltiLinDieta Then
MsgBox (Worksheets("Planejamento").Cells(i, 1).Value & " (Linha: " & i & "): Alimento não presente na Tabela UFF ou escrito incorretamente."), vbExclamation
Worksheets("Planejamento").Cells(i, 1).Select
Worksheets("Planejamento").Cells(i, 1).Interior.Color = RGB(253, 207, 207)
Worksheets("Planejamento").Cells(i, 1).Font.Bold = True
Else

If Worksheets("Planejamento").Cells(i, 2).Value = "" Or IsNumeric(Worksheets("Planejamento").Cells(i, 2).Value) = False Then

MsgBox (Worksheets("Planejamento").Cells(i, 1).Value & " (Linha: " & i & "): Preencher quantidade corretamente!"), vbExclamation
Worksheets("Planejamento").Cells(i, 2).Select
Worksheets("Planejamento").Cells(i, 2).Interior.Color = RGB(253, 207, 207)
Worksheets("Planejamento").Cells(i, 2).Font.Bold = True
Worksheets("Planejamento").Cells(i, 2).Value = "???"
Application.ScreenUpdating = True
While IsNumeric(Worksheets("Planejamento").Cells(i, 2).Value) = False Or Worksheets("Planejamento").Cells(i, 2).Value = ""
erro = InputBox("Corrigir quantidade de < " & Worksheets("Planejamento").Cells(i, 1).Value & " > para:")
Worksheets("Planejamento").Cells(i, 2).Value = erro
Worksheets("Planejamento").Cells(i, 2).Font.Bold = False
Worksheets("Planejamento").Cells(i, 2).Interior.Color = xlPatternNone
Wend
Application.ScreenUpdating = False
i = i - 1
Else
'escreve
Worksheets("Planejamento").Cells(i, 1).Interior.Color = xlPatternNone
Worksheets("Planejamento").Cells(i, 1).Font.Bold = False
Worksheets("Planejamento").Cells(i, 2).Interior.Color = xlPatternNone
Worksheets("Planejamento").Cells(i, 2).Font.Bold = False


For z = 3 To 20
'Worksheets("Planejamento").Cells(i, z) = Format(((Worksheets("Planejamento").Cells(i, 2) / 100) * Worksheets("TabelaUFF").Cells(Linha, z)), "###0.00")
Worksheets("Planejamento").Cells(i, z) = (Worksheets("Planejamento").Cells(i, 2) / 100) * Worksheets("TabelaUFF").Cells(Linha, z)
Worksheets("Planejamento").Cells(i, z).NumberFormat = "###0.00"
Worksheets("Planejamento").Cells(i, 2).HorizontalAlignment = xlCenter
Worksheets("Planejamento").Cells(i, z).Borders(xlEdgeLeft).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, z).Borders(xlEdgeTop).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, z).Borders(xlEdgeBottom).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, z).Borders(xlEdgeRight).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, z).Borders(xlInsideVertical).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, z).Borders(xlInsideHorizontal).LineStyle = xlNone
Worksheets("Planejamento").Cells(i, z).Font.Bold = False
Worksheets("Planejamento").Cells(i, z).Interior.Color = xlPatternNone
'brincadeira
If perc <> 0 Then
If CInt(perc) / perc = 1 Then
perc = (100 * i) / UltLinPlan
If perc <> 100 Then
Application.StatusBar = "Calculando... " & Fix(perc) & "% concluído."
Else
Application.StatusBar = "Cálculo " & perc & "% concluído."
End If
End If
End If
'fim da brincadeira
Next z

End If
End If
Next i

Worksheets("Planejamento").Cells(UltLinPlan + 1, 1).Value = "Totais:"
For y = 2 To 20
Total = 0
    For x = 4 To UltLinPlan
        If IsNumeric(Worksheets("Planejamento").Cells(x, y).Value) = True Then
        Total = Total + Worksheets("Planejamento").Cells(x, y).Value
        End If
    Next x
    Worksheets("Planejamento").Cells(UltLinPlan + 1, y).Value = Total
    Worksheets("Planejamento").Cells(UltLinPlan + 1, y).NumberFormat = "###0.00"
Next y
Worksheets("Planejamento").Cells(UltLinPlan + 1, 2).NumberFormat = "###0"
Worksheets("Planejamento").Cells(UltLinPlan + 1, 2).HorizontalAlignment = xlCenter
Worksheets("Planejamento").Range(Cells(UltLinPlan + 1, 1), Cells(UltLinPlan + 1, 20)).BorderAround Weight:=xlThin
Worksheets("Planejamento").Cells(UltLinPlan + 1, 1).Font.Bold = True
Worksheets("Planejamento").Rows(UltLinPlan + 1).Interior.Color = xlPatternNone

Worksheets("Planejamento").Cells(UltLinPlan + 2, 1).Value = "Kcal:"
Worksheets("Planejamento").Cells(UltLinPlan + 2, 3).Value = 4 * Worksheets("Planejamento").Cells(UltLinPlan + 1, 3).Value
Worksheets("Planejamento").Cells(UltLinPlan + 2, 4).Value = 4 * Worksheets("Planejamento").Cells(UltLinPlan + 1, 4).Value
Worksheets("Planejamento").Cells(UltLinPlan + 2, 5).Value = 9 * Worksheets("Planejamento").Cells(UltLinPlan + 1, 5).Value
Worksheets("Planejamento").Cells(UltLinPlan + 2, 2).Value = Worksheets("Planejamento").Cells(UltLinPlan + 2, 3).Value + Worksheets("Planejamento").Cells(UltLinPlan + 2, 4).Value + Worksheets("Planejamento").Cells(UltLinPlan + 2, 5).Value
Worksheets("Planejamento").Range(Cells(UltLinPlan + 2, 1), Cells(UltLinPlan + 2, 5)).BorderAround Weight:=xlThin
Worksheets("Planejamento").Range(Cells(UltLinPlan + 2, 1), Cells(UltLinPlan + 2, 5)).NumberFormat = "###0.00"
Worksheets("Planejamento").Cells(UltLinPlan + 2, 1).Font.Bold = True
Worksheets("Planejamento").Rows(UltLinPlan + 2).Interior.Color = xlPatternNone

Worksheets("Planejamento").Cells(UltLinPlan + 3, 1).Value = "%"
Worksheets("Planejamento").Cells(UltLinPlan + 3, 2).Value = "-"
'Worksheets("Planejamento").Cells(UltLinPlan + 3, 3).Value = (Worksheets("Planejamento").Cells(UltLinPlan + 2, 3).Value * 100) / Worksheets("Planejamento").Cells(UltLinPlan + 1, 3).Value
Worksheets("Planejamento").Cells(UltLinPlan + 3, 3).Value = (Worksheets("Planejamento").Cells(UltLinPlan + 2, 3).Value) / Worksheets("Planejamento").Cells(UltLinPlan + 1, 3).Value
Worksheets("Planejamento").Cells(UltLinPlan + 3, 4).Value = (Worksheets("Planejamento").Cells(UltLinPlan + 2, 4).Value) / Worksheets("Planejamento").Cells(UltLinPlan + 1, 4).Value
Worksheets("Planejamento").Cells(UltLinPlan + 3, 5).Value = (Worksheets("Planejamento").Cells(UltLinPlan + 2, 5).Value) / Worksheets("Planejamento").Cells(UltLinPlan + 1, 5).Value
Worksheets("Planejamento").Range(Cells(UltLinPlan + 3, 3), Cells(UltLinPlan + 3, 5)).NumberFormat = "0.00%"
Worksheets("Planejamento").Range(Cells(UltLinPlan + 3, 1), Cells(UltLinPlan + 3, 5)).BorderAround Weight:=xlThin
Worksheets("Planejamento").Cells(UltLinPlan + 3, 1).Font.Bold = True
Worksheets("Planejamento").Rows(UltLinPlan + 3).Interior.Color = xlPatternNone

Worksheets("Planejamento").Range(Cells(UltLinPlan + 1, 1), Cells(UltLinPlan + 1, 20)).Borders(xlEdgeTop).Weight = xlMedium
Worksheets("Planejamento").Range(Cells(UltLinPlan + 3, 1), Cells(UltLinPlan + 3, 5)).Borders(xlEdgeBottom).Weight = xlMedium
Worksheets("Planejamento").Range(Cells(UltLinPlan + 2, 5), Cells(UltLinPlan + 3, 5)).Borders(xlEdgeRight).Weight = xlMedium
Worksheets("Planejamento").Range(Cells(UltLinPlan + 1, 6), Cells(UltLinPlan + 1, 20)).Borders(xlEdgeBottom).Weight = xlMedium
Worksheets("Planejamento").Cells(UltLinPlan + 1, 20).Borders(xlEdgeRight).Weight = xlMedium

'R = Cells.SpecialCells(xlCellTypeLastCell).Row
R = UltLinPlan
'R = R - 4
If R Mod 2 = 0 Then R = R - 1
For h = R To 4 Step -2
Worksheets("Planejamento").Rows(h).Select
Worksheets("Planejamento").Rows(h).Interior.Color = RGB(210, 246, 254)
Next h

Worksheets("Planejamento").Cells(2, 7).Select


Application.ScreenUpdating = True
With Application
.Calculation = xlAutomatic
.MaxChange = 0.001
End With
ActiveWorkbook.PrecisionAsDisplayed = True
 
 
 Else
 
 MsgBox ("Nenhum alimento selecionado. Favor selecionar Alimentos na TabelaUFF e  respectivas Quantidade, antes de calcular.")
 

End If
Application.Cursor = xlDefault

End Sub




