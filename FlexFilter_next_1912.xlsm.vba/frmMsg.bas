
Public Sub CommandButton1_Click()

dFunc.msgError = ""
Unload Me

End Sub

Private Sub lblMsg_Click()

End Sub

Private Sub UserForm_Initialize()

lblMsg.Caption = dFunc.msgError

End Sub
