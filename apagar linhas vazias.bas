Attribute VB_Name = "M?dulo1"
Sub apagaLinha()
Columns("A:A").Select 'Adapte para a coluna que quiser
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
End Sub

