VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Atualizações"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16035
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Enter()

With ComboBox1
    .AddItem "EM_ANALISE"
    .AddItem "INICIO_RELACIONAMENTO_FORMAL"
    .AddItem "CONTRATADA_COM_LINHAS_PROPRIAS"
    .AddItem "CONTRATADA_COM_LINHA_BNDES"
    .AddItem "CONTRATADA_BNDES_MICROCREDITO"
    .AddItem "EXPIRADA"
    .AddItem "RECUSADA"
    .AddItem "CANCELADA"
End With


End Sub

Private Sub ComboBox2_Enter()

With ComboBox2
    .AddItem "NEGATIVA_CREDITO"
    .AddItem "FALTA_DOCUMENTACAO_OU_CADASTRO"
    .AddItem "GARANTIAS_INSUFICIENTES"
    .AddItem "CONTRATADA_COM_LINHA_BNDES"
    .AddItem "OUTROS"
End With


End Sub

Private Sub CommandButton2_Click()

Dim n As Long, l As Long
Dim tabela As ListObject
Set tabela = Planilha1.ListObjects(1)

n = UserForm1.ListBox1.Value
l = tabela.Range.Columns().Find(n, , , xlWhole).Row


tabela.Range(l, 22).Value = UserForm1.ComboBox1.Value
tabela.Range(l, 23).Value = UserForm1.ComboBox2.Value
tabela.Range(l, 24).Value = UserForm1.TextBox2.Value
tabela.Range(l, 25).Value = UserForm1.TextBox3.Value

ListBox1.Clear
Call CommandButton3_Click

'Call preencherListBox

End Sub
Private Sub CommandButton3_Click()

'seleciona a primeira celula da planilha
Range("E1").Select
'poe o foco no TextBox1
TextBox1.SetFocus
'define duas variáveis para tratar a linha atual e o contador
Dim linhaAtual As Integer
Dim contador As Integer

'verifica se o TextBox1 é diferente (<>) de vazio
If TextBox1.Text <> "" Then

  'atribui o valor zero ao contador
  contador = 0
 
  'inicia um laço While verificando se o valor da célula é diferente do
  'TextBox1 e se o contador é menor que 20. Enquanto isso for verdade o
  'laço irá ser executado
  Do While ActiveCell.Value <> TextBox1.Text
     ActiveCell.Offset(1, 0).Select
    contador = contador + 1
  Loop

End If

'compara ambos os valores convertidos para maiusculas
'If UCase(ActiveCell.Value) = UCase(TextBox1.Value) Then
If ActiveCell.Value = TextBox1.Text Then
   'limpa o listbox
   ListBox1.Clear
   'atribuir o valor da célula ativa à linhaAtual
   linhaAtual = ActiveCell.Row
   
   'inclui o valor da linha atual no listbox
   ListBox1.AddItem Sheets("Base").Range("E" & linhaAtual)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 1) = Sheets("Base").Range("B" & linhaAtual)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 2) = Sheets("Base").Range("C" & linhaAtual)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 3) = Sheets("Base").Range("V" & linhaAtual)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 4) = Sheets("Base").Range("W" & linhaAtual)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 5) = Sheets("Base").Range("X" & linhaAtual)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 6) = Sheets("Base").Range("Y" & linhaAtual)
   
   
Else
   'o registro não foi encontrado
   MsgBox "Registro não encontrado", vbCritical, "Erro"
   TextBox1.SetFocus
End If


ComboBox1.Value = ""
ComboBox2.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""


End Sub
Private Sub CommandButton4_Click()

'limpa o listbox
ListBox1.Clear
'limpa o textbox
TextBox1.Text = ""
'atribui o foco ao textbox
TextBox1.SetFocus
'chama a macro preencherListBox

ComboBox1.Value = ""
ComboBox2.Value = ""
TextBox2.Value = ""
TextBox3.Value = ""

Call preencherListBox

End Sub
Private Sub ListBox1_Change()

Dim nlin As Integer
nlin = ListBox1.ListIndex
If nlin = -1 Then Exit Sub

ComboBox1.Value = ListBox1.List(nlin, 3)
ComboBox2.Value = ListBox1.List(nlin, 4)
TextBox2.Value = ListBox1.List(nlin, 5)
TextBox3.Value = ListBox1.List(nlin, 6)


End Sub
Private Sub UserForm_Initialize()
 
 'preenche o listbox
 Call preencherListBox
 
End Sub



