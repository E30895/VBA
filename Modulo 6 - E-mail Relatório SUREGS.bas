Attribute VB_Name = "Módulo5"
Sub enviar_relatorio()

decisao = InputBox("Tem certeza que deseja enviar os e-mails? Digite 'SIM' caso queira enviar")
If decisao <> "SIM" Then
    Exit Sub
End If


Dim Outlook As Object, Novo_Email As Object
Set Outlook = CreateObject("Outlook.application")
Set Novo_Email = Outlook.createitem(0)

'Call relatorio

hora = Hour(Now)
Select Case hora
    Case Is <= 12
        saudacao = "Bom dia!"
    Case Is >= 12
        saudacao = "Boa tarde!"
End Select

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Sheets("Estatisticas").Activate

With Novo_Email
    .SentOnBehalfOfName = "Desenvolvimento_Canalmpme@banrisul.com.br"
    .display
    assinatura = Novo_Email.HTMLBody
    .To = "sureg_centro@banrisul.com.br; sureg_poacentro@banrisul.com.br; Sureg_PortoAlegre@banrisul.com.br; SUREG_Leste@banrisul.com.br; Sureg_Alto_Uruguai@banrisul.com.br; SurSerra@banrisul.com.br; sursul@banrisul.com.br; sureg_noroeste@banrisul.com.br; sureg_fronteira@banrisul.com.br; soestados@banrisul.com.br"
    .CC = "tiago_fernandes@banrisul.com.br; kelen_ferreira@banrisul.com.br; Ariel_Sturmer@banrisul.com.br; Vicente_Reis@banrisul.com.br; Carlos_Nunez@banrisul.com.br; katia_hansen@banrisul.com.br; Desenvolvimento_Analise@banrisul.com.br; Desenvolvimento_Acompanhamento@banrisul.com.br"
    .BCC = ""
    .Subject = "Relação semanal - Canal MPME e Cartão BNDES"
    .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\11. Anexo relação semanal Canal MPME.xlsx"
    .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\12. Anexo relação semanal Cartão BNDES.xlsx"
    .HTMLBody = saudacao _
    & "<br>" & "<br>" _
    & "<font size = ""3"" color=""#303E84""> Prezados (as), abaixo seguem três tabelas contendo as informações referentes ao Canal MPME: A primeira com as operações recebidas pelo CANAL MPME - BNDES na última semana, com o tipo de operação, valores, quantidade de propostas e ticket médio. A segunda com TODAS as operações recebidas desde o início do ano, e a última contendo as operações contratadas. Encaminharemos essas informações semanalmente para conhecimento. </font>" _
    & "<br>" & "<br>" _
    & "<font size = ""3"" color=""#303E84""> Segue em anexo, para consulta, as operações e seus respectivos status provenientes do canal MPME. </font>" _
    & "<br>" & "<br>" _
    & "<font size = ""3"" color=""#303E84""> Incluímos, também, em outro anexo, o arquivo importado diariamente do BNDES, de clientes e não clientes, com novas solicitações de Cartão BNDES que poderão ser analisadas pelas agências. </font>" _
    & "<br>" & "<br>" _
    & "<font size = ""3"" color=""#303E84""> Os novos pedidos podem ser pesquisados no sistema do Cartão BNDES > Solicitações > Consultar Solicitação de Cartão > Processo ou CNPJ ou Etapa 01 ou Etapa 02. </font>" _
    & "<br>" & "<br>" _
    & RangetoHTML(Range("A1:E10")) _
    & "<br>" & "<br>" _
    & RangetoHTML(Range("A13:H22")) _
    & "<br>" & "<br>" _
    & RangetoHTML(Range("A25:I37")) _
    & assinatura
    End With
    
Sheets("Index").Activate

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True


MsgBox ("Relatórios enviados com sucesso, verifique a parta 'enviados' na sua caixa de e-mail")

End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
