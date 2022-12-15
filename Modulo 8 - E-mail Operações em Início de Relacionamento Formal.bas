Attribute VB_Name = "Módulo7"
Sub Em_INICIO_RELACIONAMENTO_FORMAL()

decisao = MsgBox("Deseja Enviar os Emails?", vbYesNo)

If decisao <> 6 Then
    Exit Sub
End If


hora = Hour(Now)
Select Case hora
    Case Is <= 0
        saudacao = "Bom dia!"
    Case Is >= 12
        saudacao = "Boa tarde!"
End Select

Dim Outlook As Object, Novo_Email As Object
Dim data As Date

data = Date

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
    'Email RURAL
    If Cells(I, 22).Value = "INICIO_RELACIONAMENTO_FORMAL" And Cells(I, 17).Value = "X" And Cells(I, 20).Value = 1 Then
        With Novo_Email
            .SentOnBehalfOfName = "Desenvolvimento_Canalmpme@banrisul.com.br"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados (as), mediante o intervalo de 15 dias desde o último contato, solicitamos uma atualização do status referente à solicitação feita via canal MPME. Favor informar se o cliente foi contatado e qual a situação atual da proposta abaixo. </font>" _
            & "<br>" & "<br>" _
            & "<font color =""#007FFF"" size=""4""> Informamos que caso a operação exceda os 90 dias após o primeiro envio, ela será automaticamente expirada pelo BNDES. Atualmente restam </font>" & Cells(I, 19).Value & "<font color=""#007FFF"" size=""4""> dias. </font>" _
            & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4""> Para fins de esclarecimento, a unidade possui a liberdade para dar início, seguimento ou recusar a solicitação normalmente após o término do prazo estabelecido pelo Canal MPME.</font>" _
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
            .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\Instrucoes.pdf"
            .Send
            Set Novo_Email = Outlook.createitem(0)
            Cells(I, 16).Value = Date
        End With
        
    'EMAIL PJ
    ElseIf Cells(I, 22).Value = "INICIO_RELACIONAMENTO_FORMAL" And Cells(I, 6).Value = "" And Cells(I, 17).Value = "X" Then
        With Novo_Email
            .SentOnBehalfOfName = "Desenvolvimento_Canalmpme@banrisul.com.br"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = "<font  size=""4"">Att Administração Agência" & "<br>" & "A/C Gerente Geral e/ou Gerente de Negócios" & "<br>" & "<br>" _
            & saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados (as), mediante o intervalo de 15 dias desde o último contato, solicitamos uma atualização do status referente à solicitação feita via canal MPME. Favor informar se o cliente foi contatado e qual a situação atual da proposta abaixo. </font>" _
            & "<br>" & "<br>" _
            & "<font color =""#007FFF"" size=""4""> Informamos que caso a operação exceda os 90 dias após o primeiro envio, ela será automaticamente expirada pelo BNDES. Atualmente restam </font>" & Cells(I, 19).Value & "<font color=""#007FFF"" size=""4""> dias. </font>" _
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
            & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - DESENVOLVIMENTO_CANALMPME@BANRISUL.COM.BR" _
            & "<br>" & assinatura
            .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\Instrucoes.pdf"
            .Send
            Set Novo_Email = Outlook.createitem(0)
            Cells(I, 16).Value = Date
            End With
            
        
    'EMAIL PF
    ElseIf Cells(I, 22).Value = "INICIO_RELACIONAMENTO_FORMAL" And Cells(I, 6).Value <> "" And Cells(I, 17).Value = "X" Then
        With Novo_Email
            .SentOnBehalfOfName = "Desenvolvimento_Canalmpme@banrisul.com.br"
            .display
            assinatura = Novo_Email.HTMLBody
            .To = Cells(I, 4).Value
            .CC = ""
            .Subject = "CANAL MPME - BNDES PROTOCOLO: " & Cells(I, 5).Value
            .HTMLBody = "<font  size=""4"">Att Administração Agência" & "<br>" & "A/C Gerente Geral e/ou Gerente de Negócios" & "<br>" & "<br>" _
            & saudacao & "<br>" & "<br>" _
            & "<font color=""#007FFF"" size=""4"">Prezados (as), mediante o intervalo de 15 dias desde o último contato, solicitamos uma atualização do status referente à solicitação feita via canal MPME. Favor informar se o cliente foi contatado e qual a situação atual da proposta abaixo. </font>" _
            & "<br>" & "<br>" _
            & "<font color =""#007FFF"" size=""4""> Informamos que caso a operação exceda os 90 dias após o primeiro envio, ela será automaticamente expirada pelo BNDES. Atualmente restam </font>" & Cells(I, 19).Value & "<font color=""#007FFF"" size=""4""> dias. </font>" _
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
            & "<br>" & "<br>" & "<br>" & "UNIDADE DE DESENVOLVIMENTO - DESENVOLVIMENTO_CANALMPME@BANRISUL.COM.BR" _
            & "<br>" & assinatura
            .attachments.Add "I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\Instrucoes.pdf"
            .Send
            Set Novo_Email = Outlook.createitem(0)
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

