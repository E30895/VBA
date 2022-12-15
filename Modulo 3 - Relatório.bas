Attribute VB_Name = "Módulo2"
Sub relatorio()

Dim base As Object
Dim relatorio_caminho As String
Dim relatorio As Object



'-------------------------Introdução-----------------------'

'Esse programa consiste em uma sequencia de passos lógicos:

'1. Procurar e abrir o caminho do relatorio;
'2. Apagar todos os valores já existentes;
'3. Importar todos os dados da base;
'4. Formatar a planilha de relatorio;


'-------------------------------------------------------------

'Desativando as atualizações de tela
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Sheets("Base").Activate

'LOCALIZANDO E ABRINDO O ARQUIVO
caminho = "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\11. Anexo relação semanal Canal MPME.xlsx"
Workbooks.Open (caminho)


'DELETANDO OS DADOS JÁ EXISTENTES E FORMATANDO
Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("A:ZA").Select
    Selection.Delete


'IMPORTANDO A PRIMIRA COLUNA (SUREG)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("A:A").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("A:A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A SEGUNDA COLUNA (COD AG)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("B:B").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("B:B").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A TERCEIRA COLUNA (AG)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("C:C").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("C:C").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A QUARTA COLUNA (PROTOCOLO)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("E:E").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("D:D").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A QUINTA COLUNA (CPF)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("F:F").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("E:E").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


''IMPORTANDO A SEXTA COLUNA (CNPJ)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("G:G").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("F:F").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

'IMPORTANDO A SETIMA COLUNA (NOME)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("H:H").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("G:G").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
          
          
'IMPORTANDO A OITAVA COLUNA (VALOR)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("K:K").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("H:H").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A NONVA COLUNA (DATA)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("O:O").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("I:I").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A DECIMA COLUNA (FINALIDADE)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("L:L").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("J:J").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A DECIMA PRIMEIRA COLUNA (LINHA RECOMENDADA)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("M:M").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("K:K").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'IMPORTANDO A DECIMA SEGUNDA COLUNA (STATUS)
Windows("6. Controle de fluxo MPME").Activate
Sheets("Base").Activate
    Columns("V:V").Select
    Selection.Copy

Windows("11. Anexo relação semanal Canal MPME").Activate
    Columns("L:L").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False


'FORMATANDO OS DADOS
    Range("A1:L1").Font.Bold = True 'Adicionando negrito'
    Range("A1:L1").Interior.ColorIndex = 23 'Pintar o cabecario de azul'
    Columns("A:L").AutoFit 'Ajustando o tamanho das células no relatorio'
    Columns("F:F").NumberFormat = "00"".""000"".""000""/""0000""-""00" 'Formatando CNPJ
    Columns("E:E").NumberFormat = "000"".""000"".""000""-""00" 'Formatando CPF
    Columns("I:I").NumberFormat = "dd/mm/yy" 'Formatando a data br'
    Columns("H:H").NumberFormat = "$ #,##0.00" 'formatando a moeda'
    
    'Aplicando filtros'
    Range("A1").Select
    Selection.AutoFilter
    
    'Ordenando de A-Z'
    ActiveWorkbook.Worksheets("MPME").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("MPME").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("MPME").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Alterando a cor do cabecalho'
    Range("A1:L1").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    
    'Centralizando o cabecalho
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

'Salvando e fechando
ActiveWorkbook.Save
ActiveWorkbook.Close


'Ativando a atualização de tela
caminho = Null

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

Sheets("Index").Activate
MsgBox ("Processo terminado. Planilha pronta")


    
End Sub
