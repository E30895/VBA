Attribute VB_Name = "Módulo6"
Sub Enviar_Email_EM_ANALISE()


decisao = MsgBox("Deseja Enviar os Emails?", vbYesNo)

If decisao <> 6 Then
    Exit Sub
End If


Dim Outlook As Object, Novo_Email As Object
Dim data As Date
data = Date
Set Outlook = CreateObject("Outlook.application")
Set Novo_Email = Outlook.createitem(0)
Dim hora As Integer

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
        If Cells(I, 22).Value = "EM_ANALISE" And Cells(I, 6).Value <> "" And Cells(I, 17).Value = "X" And Cells(I, 20).Value = 1 Then
              With Novo_Email
                .SentOnBehalfOfName = "caixa de saída"
                .display
                assinatura = Novo_Email.HTMLBody
                .To = Cells(I, 4).Value
                .CC = ""
                .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
                .HTMLBody = saudacao & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4"">Prezados (as), encaminhamos um e-mail referente a solicitação supracitada no dia </font>" & Cells(I, 16).Value & "<font color=""#007FFF"" size=""4""> e informamos que ainda não recebemos um retorno referente ao contato do canal MPME. Favor informar se o Lead foi contatado e qual a situação atual desta solicitação. </font>" _
                & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4""> Em caso de ausência de retorno em até 60 dias após o primeiro envio, realizado na data </font>" & Cells(I, 15).Value & "<font color=""#007FFF"" size=""4""> a solicitação será automaticamente expirada. Atualmente restam </font>" & Cells(I, 19).Value & "<font color=""#007FFF"" size=""4""> dias. </font>" _
                & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a unidade possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME.</font>" _
                & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> ID BNDES: </b> </font>" & Cells(I, 5).Value & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> Lead: </b> </font>" & Cells(I, 8).Value _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CPF: </b> </font>" & Cells(I, 6).Value & "<br>" _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CNPJ: </b> </font>" & Cells(I, 7).Value & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> Telefone: </b> </font>" & Cells(I, 10) & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> E-mail: </b> </font>" & Cells(I, 9).Value _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Valor solicitado R$: </b> </font>" & Cells(I, 11).Value _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Descrição do solicitado: </b> </font>" & Cells(I, 12) & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> Linha de Crédito Sugerida: </b> </font>" & Cells(I, 13).Value _
                & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - " _
                & "<br>" & assinatura
                .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\Instrucoes.pdf"
                .Send
                Set Novo_Email = Outlook.createitem(0)
                Cells(I, 16).Value = Date
            End With
        
    'ENVIAR PJ
     ElseIf Cells(I, 22).Value = "EM_ANALISE" And Cells(I, 6).Value = "" And Cells(I, 17).Value = "X" Then
            With Novo_Email
            .SentOnBehalfOfName = "caixa de saída"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = "<font  size=""4"">Att Administração Agência" & "<br>" & "A/C Gerente Geral e/ou Gerente de Negócios" & "<br>" & "<br>" _
            & saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados (as), encaminhamos um e-mail referente a solicitação supracitada no dia </font>" & Cells(I, 16).Value & "<font color=""#007FFF"" size=""4""> e informamos que ainda não recebemos um retorno referente ao contato do canal MPME. Favor informar se o Lead foi contatado e qual a situação atual desta solicitação. </font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> Em caso de ausência de retorno em até 60 dias após o primeiro envio, realizado na data </font>" & Cells(I, 15).Value & "<font color=""#007FFF"" size=""4""> a solicitação será automaticamente expirada. Atualmente restam </font>" & Cells(I, 19).Value & "<font color=""#007FFF"" size=""4""> dias. </font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a agência possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME.</font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> ID BNDES: </b> </font>" & Cells(I, 5).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Empresa: </b> </font>" & Cells(I, 8).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CNPJ: </b> </font>" & Cells(I, 7).Value & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Telefone: </b> </font>" & Cells(I, 10) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> E-mail: </b> </font>" & Cells(I, 9).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Valor solicitado R$: </b> </font>" & Cells(I, 11).Value _
            & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Descrição do solicitado: </b> </font>" & Cells(I, 12) & "<br>" _
            & "<font color=""#007FFF"" size=""4""> <b> Linha de Crédito Sugerida: </b> </font>" & Cells(I, 13).Value _
            & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO -" _
            & "<br>" & assinatura
            .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\Instrucoes.pdf"
            .Send
            Set Novo_Email = Outlook.createitem(0)
            Cells(I, 16).Value = Date
            End With
           
        
        'EMAIL PF
        ElseIf Cells(I, 22).Value = "EM_ANALISE" And Cells(I, 6).Value <> "" And Cells(I, 17).Value = "X" Then
            With Novo_Email
                .SentOnBehalfOfName = "caixa de saída"
                .display
                assinatura = Novo_Email.HTMLBody
                .To = Cells(I, 4).Value
                .CC = ""
                .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
                .HTMLBody = "<font  size=""4"">Att Administração Agência" & "<br>" & "A/C Gerente Geral e/ou Gerente de Negócios" & "<br>" & "<br>" _
                & saudacao & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4"">Prezados (as), encaminhamos um e-mail referente a solicitação supracitada no dia </font>" & Cells(I, 16).Value & "<font color=""#007FFF"" size=""4""> e informamos que ainda não recebemos um retorno referente ao contato do canal MPME. Favor informar se o Lead foi contatado e qual a situação atual desta solicitação. </font>" _
                & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4""> Em caso de ausência de retorno em até 60 dias após o primeiro envio, realizado na data </font>" & Cells(I, 15).Value & "<font color=""#007FFF"" size=""4""> a solicitação será automaticamente expirada. Atualmente restam </font>" & Cells(I, 19).Value & "<font color=""#007FFF"" size=""4""> dias. </font>" _
                & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a agência possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME.</font>" _
                & "<br>" & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> ID BNDES: </b> </font>" & Cells(I, 5).Value & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> Lead: </b> </font>" & Cells(I, 8).Value _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> CPF: </b> </font>" & Cells(I, 6).Value & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> Telefone: </b> </font>" & Cells(I, 10) & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> E-mail: </b> </font>" & Cells(I, 9).Value _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Valor solicitado R$: </b> </font>" & Cells(I, 11).Value _
                & "<br>" & "<font color=""#007FFF"" size=""4""> <b> Descrição do solicitado: </b> </font>" & Cells(I, 12) & "<br>" _
                & "<font color=""#007FFF"" size=""4""> <b> Linha de Crédito Sugerida: </b> </font>" & Cells(I, 13).Value _
                & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - " _
                & "<br>" & assinatura
                .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\Instrucoes.pdf"
                Set Novo_Email = Outlook.createitem(0)
                .Send
                Cells(I, 16).Value = Date
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


