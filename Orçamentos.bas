Attribute VB_Name = "Módulo1"
Option Explicit

Public Appl, SapGuiAuto, Connection, session, WScript 'Application, SapGuiAuto, Connection, session, WScript'

Sub verificar_orcamento()

'Conexão do VBA com o SAP'

If Not IsObject(Appl) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Appl = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Appl.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

'Entrando com os numeros das peps'

Dim PEP As String
Dim i, j As Integer

Dim LastRow
    LastRow = Sheets("Indice").Cells(Rows.Count, 1).End(xlUp).Row

For i = 3 To LastRow

'Acessando transação ZPS_FLCX'

session.findById("wnd[0]/tbar[0]/okcd").Text = "/nZPS_FLCX"
session.findById("wnd[0]").sendVKey 0

PEP = Range("A" & i).Value

session.findById("wnd[0]/usr/ctxtS_PSPID-LOW").Text = PEP
session.findById("wnd[0]/tbar[1]/btn[8]").press

Application.Wait (Second(Now)) + 5

'Ajustando o peril e salvando os dados'
session.findById("wnd[0]/usr/cntlCCONTAINER1/shellcont/shell/shellcont[1]/shell[0]").pressContextButton "&LOAD"
session.findById("wnd[0]/usr/cntlCCONTAINER1/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItem "&LOAD"
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").currentCellRow = 17
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "17"
session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell
session.findById("wnd[0]/usr/cntlCCONTAINER1/shellcont/shell/shellcont[1]/shell[0]").pressContextButton "&PRINT_BACK"
session.findById("wnd[0]/usr/cntlCCONTAINER1/shellcont/shell/shellcont[1]/shell[0]").selectContextMenuItem "&PRINT_PREV_ALL"
session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[1]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press

'Colando os dados no excel'

Indice.Range("A50").Select
    ActiveSheet.Paste
    Rows("50:60").Select
    Selection.Delete Shift:=xlUp
    Rows("53:55").Select
    Selection.Delete Shift:=xlUp

'Editando os dados'

For j = 50 To 52
    Range("C" & j).FormulaR1C1 = "=TRIM(MID(R" & j & "C1,SEARCH("","",R" & j & "C1,1)+3,100))"
    Range("D" & j).FormulaR1C1 = "=RIGHT(R" & j & "C1,1)"
        If Range("D" & j) = "-" Then
            Range("E" & j).FormulaR1C1 = "=SEARCH(""-"",R" & j & "C3,1)"
            Range("F" & j).FormulaR1C1 = "=LEFT(RC[-3],RC[-1]-1)"
            Range("G" & j).FormulaR1C1 = "=CONCATENATE(RC[-3],RC[-1])"
        Else
            Range("G" & j).FormulaR1C1 = Range("C" & j).Value
    End If
    
Next
    
'Colando os novos valores'

Indice.Range("G50:G52").Copy
Indice.Range("C" & i & ":E" & i).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

Indice.Range("A50:H52").Clear
Next

'Convertendo os numeros'

Dim rngCelula As Range

Set rngCelula = Range("C3" & ":" & "E" & LastRow)
      With rngCelula
        .NumberFormat = "General"
        .FormulaLocal = rngCelula.Value
        .Style = "Currency"
      End With

End Sub
