Attribute VB_Name = "Module1"
Private Function HeatCell(TotOcorrencias As Long, NOcorr As Long)
'Fun��o que gera cor (long) de verde para vermelho crescente
'-------------------------------------------------------------
'Usar ".Interior.Color = Ncor" que dar� a cor diretamente.
'ex.:
'ActiveCell.Interior.Color = HeatCell(10, 10)
'-------------------------------------------------------------
Dim NCor, CorFinal As Long
NCor = (510 * NOcorr) / TotOcorrencias
If NCor <= 255 Then
CorFinal = RGB(NCor, 255, 5)
Else
If (NCor - 255) > 255 Then
CorFinal = RGB(255, 255, 5)
Else
CorFinal = RGB(255, 255 - (NCor - 255), 5)
End If
End If
HeatCell = CorFinal
End Function
