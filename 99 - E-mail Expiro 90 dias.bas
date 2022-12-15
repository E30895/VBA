Attribute VB_Name = "Módulo9"
Sub Expiro_90d()

decisao = MsgBox("Deseja Enviar os Emails e ENCERRAR as operações?", vbYesNo)

If decisao <> 6 Then
    Exit Sub
End If


Dim Outlook As Object, Novo_Email As Object
Set Outlook = CreateObject("Outlook.application")
Set Novo_Email = Outlook.createitem(0)

hora = Hour(Now)
Select Case hora
    Case Is <= 12
        saudacao = "Bom dia!"
    Case Is >= 12
        saudacao = "Boa tarde!"
End Select

Sheets("Base").Activate

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Ultima_linha = Sheets("Base").Range("E:E").End(xlDown).Row


Dim I
I = 2

While I <= Ultima_linha
    'EMAIL RURAL
    If Cells(I, 22) = "INICIO_RELACIONAMENTO_FORMAL" And Cells(I, 18).Value >= 90 And Cells(I, 20).Value = 1 Then
        With Novo_Email
            .SentOnBehalfOfName = "caixa de saída"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados(as), informamos que a solicitação supracitada expirou-se nesta data. </font>" _
            & "<br>" & "<br>" & saudacao _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a unidade possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME - BNDES.</font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> ID BNDES: </b> </font>" & Cells(I, 5).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Lead: </b> </font>" & Cells(I, 8).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CPF: </b> </font>" & Cells(I, 6).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> CNPJ: </b> </font>" & Cells(I, 7).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Telefone: </b> </font>" & Cells(I, 10) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> E-mail: </b> </font>" & Cells(I, 9).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Valor solicitado R$: </b> </font>" & Cells(I, 11).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Descrição do solicitado: </b> </font>" & Cells(I, 12) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Linha de Crédito Sugerida: </b> </font>" & Cells(I, 13).Value _
            & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - DESENVOLVIMENTO_CANALMPME@BANRISUL.COM.BR" _
            & "<br>" & assinatura
            .Send
            Set Novo_Email = Outlook.createitem(0)
            Cells(I, 16).Value = Date
            Cells(I, 22).Value = "EXPIRADA"
            Cells(I, 23).Value = "OUTROS"
        End With
    
    'EMAIL PJ
    ElseIf Cells(I, 22) = "INICIO_RELACIONAMENTO_FORMAL" And Cells(I, 18).Value >= 90 And Cells(I, 6).Value = "" Then
        With Novo_Email
            .SentOnBehalfOfName = "caixa de saída"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = "<font  size=""4"">Att Administração Agência" & "<br>" & "A/C Gerente Geral e/ou Gerente de Negócios" & "<br>" & "<br>" _
            & saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados(as), informamos que a solicitação supracitada expirou-se nesta data. </font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a agência possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME - BNDES.</font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> ID BNDES: </b> </font>" & Cells(I, 5).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Lead: </b> </font>" & Cells(I, 8).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CNPJ: </b> </font>" & Cells(I, 7).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Telefone: </b> </font>" & Cells(I, 10) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> E-mail: </b> </font>" & Cells(I, 9).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Valor solicitado R$: </b> </font>" & Cells(I, 11).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Descrição do solicitado: </b> </font>" & Cells(I, 12) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Linha de Crédito Sugerida: </b> </font>" & Cells(I, 13).Value _
            & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - DESENVOLVIMENTO_CANALMPME@BANRISUL.COM.BR" _
            & "<br>" & assinatura
            .Send
            Set Novo_Email = Outlook.createitem(0)
            Cells(I, 16).Value = Date
            Cells(I, 22).Value = "EXPIRADA"
            Cells(I, 23).Value = "OUTROS"
        End With
    
    'Email PF
    ElseIf Cells(I, 22) = "INICIO_RELACIONAMENTO_FORMAL" And Cells(I, 18).Value >= 90 And Cells(I, 6).Value <> "" Then
        With Novo_Email
            .SentOnBehalfOfName = "caixa de saída"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = "<font  size=""4"">Att Administração Agência" & "<br>" & "A/C Gerente Geral e/ou Gerente de Negócios" & "<br>" & "<br>" _
            & saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados(as), informamos que a solicitação supracitada expirou-se nesta data. </font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a agência possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME - BNDES.</font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> ID BNDES: </b> </font>" & Cells(I, 5).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Lead: </b> </font>" & Cells(I, 8).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CPF: </b> </font>" & Cells(I, 6).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Telefone: </b> </font>" & Cells(I, 10) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> E-mail: </b> </font>" & Cells(I, 9).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Valor solicitado R$: </b> </font>" & Cells(I, 11).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Descrição do solicitado: </b> </font>" & Cells(I, 12) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Linha de Crédito Sugerida: </b> </font>" & Cells(I, 13).Value _
            & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - DESENVOLVIMENTO_CANALMPME@BANRISUL.COM.BR" _
            & "<br>" & assinatura
            .Send
            Set Novo_Email = Outlook.createitem(0)
            Cells(I, 16).Value = Date
            Cells(I, 22).Value = "EXPIRADA"
            Cells(I, 23).Value = "OUTROS"
        End With
    End If
    

I = I + 1
Wend

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

Sheets("Index").Activate
MsgBox ("Emails enviados, operação finalizada.")
End Sub

