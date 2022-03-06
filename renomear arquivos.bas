Attribute VB_Name = "Módulo1"
Sub ren()
Dim arquivo As String
Dim plan As Worksheet
Dim caminho As String
Dim tp As String
Dim q As Long
q = 1
pasta = "C:\Renomear\"

For Each plan In Worksheets
    If plan.Name Like "arquivos" Then
        Sheets("arquivos").Delete
        Application.DisplayAlerts = False
        Exit For
    End If
Next plan

Cells(1, 6) = Left(Cells(1, 1), 4)
Cells(1, 7) = Right(Cells(1, 1), 6)
Cells(1, 8) = Cells(1, 7) + 19

For i = 2 To 51

    Cells(i, 7) = Cells(i - 1, 7) + 20
    Cells(i, 1) = Cells(1, 6) & Cells(i, 7)
    Cells(i, 8) = Cells(i - 1, 8) + 20
    Cells(i, 3) = Cells(1, 6) & Cells(i, 8)
    Cells(i, 4) = Cells(i, 1) & " A " & Cells(i, 3) & ".xls"

Next
Cells(1, 3) = Cells(1, 6) & (Cells(1, 7) + 19)
Cells(1, 4) = Cells(1, 1) & " A " & Cells(1, 3) & ".xls"
Columns("A:J").AutoFit
Rows(51).Delete
Columns("E:I").Delete
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "arquivos"
tp = "*.xls*"
caminho = Dir(pasta & tp, vbDirectory)
While caminho <> ""
ActiveSheet.Range("A" & q) = caminho
q = q + 1
caminho = Dir()
Wend


For e = 1 To 50
Cells(e, 5) = pasta & Cells(e, 1)
Next

Worksheets("RENOMEAR").Activate
Range("D1:D55").Copy
Worksheets("arquivos").Activate
Range("G1").Select
ActiveSheet.Paste

For a = 1 To 50
Cells(a, 9) = pasta & Cells(a, 7)
Next
Columns("A:D").Delete
Columns("C:D").Delete
Columns("A:E").AutoFit

For o = 1 To 50
Name Cells(o, 1) As Cells(o, 3)
Next

Worksheets("RENOMEAR").Activate
Rows("2:60").Delete
Columns("B:E").Delete

End Sub
