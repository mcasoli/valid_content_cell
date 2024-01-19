VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Application.EnableEvents = False
 
    Dim User As String
    If Target.Cells.Count = 1 = True Then
        Debug.Print "Apenas uma celula alterada"
        User = Target.Value2
    Else
        Debug.Print "Multiplas linhas alteradas"
        User = Target.Cells(1, 1).Value2
    End If

    If User = "" Or Target.Column <> 15 Then
        Application.EnableEvents = True
        Exit Sub
    End If
    usuarios = pegar_usuarios()
    Dim achou As Boolean
    Dim password As String
    achou = False
    password = InputBox("Digite a senha para desbloquear a célula", "Senha")
    If password = "" Then
        Throw_Client_Error ("Senha vazia ou inválida.")
        Target.Locked = True
        Target.ClearContents
    Else
        For Each Item In usuarios
            email_linha = Item(0)
            password_linha = Item(1)
            Debug.Print (email_linha)
            Debug.Print (password_linha)
            If User = email_linha And password_linha = password Then
                Target.Locked = False
                achou = True
            End If
        Next
        If achou = False Then
            Throw_Client_Error "Usuário ou senha inválidos. Tente novamente."
            Target.ClearContents
            Target.Locked = True
        End If
    End If
    Application.EnableEvents = True
End Sub
Sub Throw_Client_Error(message As String)
    Debug.Print message
    MsgBox message
End Sub
