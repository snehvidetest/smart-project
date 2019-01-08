
Public Sub cmdNo_Click()

dFunc.msgYesNo = "NEJ"
Me.Hide

End Sub
Public Sub cmdYes_Click()

dFunc.msgYesNo = "JA"
Me.Hide

End Sub
Private Sub UserForm_Initialize()

lblMsg.Caption = dFunc.msgYesNoTxt
dFunc.msgYesNo = ""

End Sub