Attribute VB_Name = "Módulo1"
Sub preencherListBox()

Dim ultimaLinha As Long
Dim linha As Long

'retorna o valor da ultima linha preenchida da coluna
ultimaLinha = Sheets("Base").Range("E:E").End(xlDown).Row

'percorre a partir da linha 2 até a ultima linha
For linha = 2 To ultimaLinha
   UserForm1.ListBox1.AddItem Sheets("Base").Range("E" & linha)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 1) = Sheets("Base").Range("B" & linha)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 2) = Sheets("Base").Range("C" & linha)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 3) = Sheets("Base").Range("V" & linha)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 4) = Sheets("Base").Range("W" & linha)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 5) = Sheets("Base").Range("X" & linha)
   UserForm1.ListBox1.List(UserForm1.ListBox1.ListCount - 1, 6) = Sheets("Base").Range("Y" & linha)
   
Next
End Sub
Sub chamarFormulario()
    
    Application.Visible = False
    Worksheets("Base").Activate
    UserForm1.Show
    Worksheets("Index").Activate
    Application.Visible = True

End Sub

