Attribute VB_Name = "Módulo3"
Sub BNDES()

Dim relatorio_caminho As String
Dim relatorio As Object
Dim I
Dim R

Sheets("Base").Activate

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

caminho = "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\9. RETORNO BNDES.xlsx"
Workbooks.Open (caminho)


Windows("9. RETORNO BNDES").Activate
Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

    
'IMPORTANDO A PRIMIRA COLUNA
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Range("E2").Select
    Range(Selection, Selection.End(xlDown)).Copy

Windows("9. RETORNO BNDES").Activate
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A SEGUNDA COLUNA
Windows("6. Controle de fluxo MPME").Activate
    Range("V2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

Windows("9. RETORNO BNDES").Activate
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A TERCEIRA COLUNA

Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate

I = 2

Ultima_linha = Sheets("Base").Range("A:A").End(xlDown).Row

While I <= Ultima_linha
    Windows("6. Controle de fluxo MPME").Activate
    Sheets("Base").Activate
    If Cells(I, 22).Value = "EXPIRADA" Or Cells(I, 22).Value = "RECUSADA" Or Cells(I, 22).Value = "CANCELADA" Then
    
        Cells(I, 23).Copy
        Windows("9. RETORNO BNDES").Activate
        Cells(I, 3).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    End If
    
 I = I + 1
Wend

'IMPORTANDO A ULTIMA COLUNA

n = 2

While n <= Ultima_linha
    Windows("6. Controle de fluxo MPME").Activate
    Sheets("Base").Activate
    
    If Cells(n, 19).Value = "CONTRATADA" Then
        Windows("6. Controle de fluxo MPME").Activate
        Cells(n, 25).Copy
        Windows("9. RETORNO BNDES").Activate
        Cells(n, 4).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Else
        End If

n = n + 1
Wend

Windows("6. Controle de fluxo MPME").Activate
Sheets("Index").Activate

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

Windows("9. RETORNO BNDES").Activate
ActiveWorkbook.Save

End Sub





