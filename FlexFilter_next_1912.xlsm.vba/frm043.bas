Public Sub CommandButton1_Click()
Me.Hide

If frm004.ActiveControl Is Nothing Then
    ' ingen værdi
Else
        frm004.Hide
        SFunc.ShowFunc ("frm005")
        GoTo ending
End If

If frm002.ActiveControl Is Nothing Then
    ' ingen værdi
Else
    frm002.Hide
        If frm002.forkertData = False Then
            SFunc.ShowFunc ("frm003")
        Else
            SFunc.ShowFunc ("frm005")
        End If
        
        GoTo ending
End If











ending:
End Sub

Public Sub CommandButton2_Click()
Me.Hide
' frm002.txtModtStart.Value = ""
' frm002.txtModtSlut.Value = ""

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub