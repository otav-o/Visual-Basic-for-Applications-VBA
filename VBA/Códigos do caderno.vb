'002

Sub Aula2_ConfigurarPágina()
Dim Celulas As Range    
Range("B3").Select
Set Celulas = Selection.CurrentRegion
ActiveSheet.PageSetup.PrintArea = Celulas.Adress

'003
'Deslocamento de células
ActiveCell.Offset(2, 2).Select  'duas para a direita e duas para baixo
Range("B3").Offset(-1, 0).Select 
Cells(3, 2).Offset(n, 0).Select
