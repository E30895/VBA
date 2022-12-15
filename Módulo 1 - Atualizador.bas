Attribute VB_Name = "Módulo0"
Sub atualizar()

escolha = MsgBox("Deseja Atualizar a base de dados?", vbYesNo)

If escolha <> "6" Then
    Exit Sub
End If
'Pegando o valor da ultima linha da base de dados

Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Sheets("Base").Activate
lin1 = Sheets("Base").Cells(2, 5).End(xlDown).Row
x1 = Cells(lin1, 5)

'Pegando o valor da ultima linha da planilha do alexandre

caminho = "I:\Desenvolvimento-GCO\GCO\ACOMPANHAMENTO\Canal MPME\Analise_06-17.xlsm"
Workbooks.Open (caminho)

'Application.Wait (Now + TimeValue("0:00:30"))

Workbooks("Analise_06-17").Activate
lin2 = Sheets("propostas").Cells(4, 1).End(xlDown).Row
x2 = Cells(lin2, 1)


'Enquanto o valor da ultima linha não forem os mesmos x2-1
calc = lin2

If x2 = x1 Then
    
    Workbooks("Analise_06-17").Activate
    Workbooks("Analise_06-17").Save
    Workbooks("Analise_06-17").Close
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox ("A planilha já está atualizada"), , AVISO
    Sheets("Index").Activate
    Exit Sub
End If

While x2 <> x1
    calc = calc - 1
    x2 = Cells(calc, 1)
Wend

calc = calc + 1
Y1 = lin2 - calc
y2 = lin2 - Y1

While y2 <= lin2
    'IMPORTANDO O PROTOCOLO
    Workbooks("Analise_06-17").Activate
    lin1 = lin1 + 1
    Cells(y2, 1).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 5).Select
    ActiveSheet.Paste
    
    'IMPORTANDO EMAIL AGENCIA
    Workbooks("Analise_06-17").Activate
    Cells(y2, 22).Copy
    
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 4).Select
    ActiveSheet.Paste
    
    'IMPORTANDO O CPF
    Workbooks("Analise_06-17").Activate
    Cells(y2, 3).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 6).Select
    ActiveSheet.Paste

    
    'IMPORTANDO O CNPJ
    Workbooks("Analise_06-17").Activate
    Cells(y2, 2).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 7).Select
    ActiveSheet.Paste
    
    'IMPORTANDO O NOME
    Workbooks("Analise_06-17").Activate
    Cells(y2, 4).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 8).Select
    ActiveSheet.Paste
    
    'IMPORTANDO E-MAIL
    Workbooks("Analise_06-17").Activate
    Cells(y2, 5).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 9).Select
    ActiveSheet.Paste
    
    'IMPORTANDO TELEFONE
    Workbooks("Analise_06-17").Activate
    Cells(y2, 6).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 10).Select
    ActiveSheet.Paste
    
    'IMPORTANDO VALOR
    Workbooks("Analise_06-17").Activate
    Cells(y2, 7).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 11).Select
    ActiveSheet.Paste
    
    'IMPORTANDO FINALIDADES
    Workbooks("Analise_06-17").Activate
    Cells(y2, 10).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 12).Select
    ActiveSheet.Paste
    
    'IMPORTANDO LINHA RECOMENDADA
    Workbooks("Analise_06-17").Activate
    Cells(y2, 13).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 13).Select
    ActiveSheet.Paste
    
    'IMPORTANDO ESTADO
    Workbooks("Analise_06-17").Activate
    Cells(y2, 14).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 14).Select
    ActiveSheet.Paste
    
    'IMPORTANDO STATUS
    Workbooks("Analise_06-17").Activate
    Cells(y2, 12).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 22).Select
    ActiveSheet.Paste
    
    'IMPORTANDO DATA DO PRIMEIRO EMAIL
    Workbooks("Analise_06-17").Activate
    Cells(y2, 28).Copy
    
    Workbooks("6. Controle de fluxo MPME").Activate
    Cells(lin1, 15).Select
    ActiveSheet.Paste
    
    Cells(lin1, 16).Value = Cells(lin1, 15).Value
    
    
    Cells(lin1, 1) = WorksheetFunction.VLookup(Cells(lin1, 4).Value, Sheets("Suregs").Range("A:D"), 2, FALSO)
    Cells(lin1, 2) = WorksheetFunction.VLookup(Cells(lin1, 4).Value, Sheets("Suregs").Range("A:D"), 3, FALSO)
    Cells(lin1, 3) = WorksheetFunction.VLookup(Cells(lin1, 4).Value, Sheets("Suregs").Range("A:D"), 4, FALSO)
    
    Cells(lin1, 22) = "EM_ANALISE"
    
    Range("F:F").NumberFormat = "000"".""000"".""000""-""00"
    Range("G:G").NumberFormat = "00"".""000"".""000""/""0000""-""00"
    Range("J:J").NumberFormat = "(##)"" ""00000""-""0000"
    Range("K:K").NumberFormat = "$ #,##0.00"
    
    y2 = y2 + 1
    
    
    Wend

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

MsgBox ("Operações atualizadas!"), vbInformation

Workbooks("Analise_06-17").Activate
Workbooks("Analise_06-17").Save
Workbooks("Analise_06-17").Close

End Sub


