Attribute VB_Name = "Módulo4"
Sub Abrir_bndes()

Dim navegador As New ChromeDriver

'Call BNDES

login = InputBox("Digite seu usuário ou CPF", "Login")
senha = InputBox("Digite sua senha do ambiente operacional (no final do programa ele esvazia a variável, não vou roubar sua senha nem financiar um motorhome)", "Senha")

If login = "" Then
        MsgBox ("o login não pode ficar em branco")
        Exit Sub
End If

If senha = "" Then
    MsgBox ("A senha não pode ficar em branco")
    Exit Sub
End If
    

With navegador
    .Get ("https://portal.bndes.gov.br/prc/#/login?tipoPessoa=AF")
    .Wait (5000)
    .FindElementByXPath("//*[@id=""txt.CNPJ-ux4""]").SendKeys ("92.702.067/0001-96")
    .FindElementByXPath("//*[@id=""txt.userName-ux4""]").SendKeys (login)
    .FindElementByXPath("//*[@id=""txt.senha-ux4""]").SendKeys (senha)
    .FindElementByXPath("//*[@id=""btn.acessar-ux4""]").Click
    .Wait (5000)
    .Get ("https://ws.bndes.gov.br/pme/#/propostas/upload")
    .Wait (5000)
    .FindElementByXPath("//*[@id=""fileToUpload""]").SendKeys ("I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\10. RETORNO BNDES.csv")
    .Wait (2000)
    .FindElementByXPath("/html/body/pme/div/upload-propostas/button[2]").Click
    .Wait (15000)
    
    '"I:\Desenvolvimento-GAA\GAA\CANAL MPME - COMPLEMENTARES\10. RETORNO BNDES.csv"
End With

login = Null
senha = Null

End Sub
