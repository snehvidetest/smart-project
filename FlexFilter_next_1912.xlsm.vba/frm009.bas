Public Sub OKButton_Click()
If OptionButton1.Value = False And OptionButton2.Value = False Then
    dFunc.msgError = "V�lg venligst et svar for at fors�tte"
    SFunc.ShowFunc ("frmMsg")
    GoTo ending
End If

Worksheets("SpmSvar").Range("C19:C19").Value = Controls("Label1").Caption

If OptionButton1.Value = True Then
    Worksheets("SpmSvar").Range("D19:D19").Value = "Ja"
    
    Worksheets("Gruppering").Range("C2:C2").Value = "NEJ"
    Worksheets("Gruppering").Range("C3:C3").Value = "JA"
    
    Worksheets("Population").Range("B16:B16").Value = "JA"
    Worksheets("Population").Range("B17:B17").Value = "NEJ"

ElseIf OptionButton1.Value = False Then
    
    Worksheets("SpmSvar").Range("D19:D19").Value = "Nej"

End If

If OptionButton2.Value = True Then
End If

Me.Hide

' Worksheets("Konfiguration").Activate
' Activate sheet

If OptionButton1 = True Then
    SFunc.ShowFunc ("frm039")
    'frm039.Show

ElseIf OptionButton2 = True Then
    SFunc.ShowFunc ("frm010")
    'frm010.Show

End If

ending:
End Sub




Private Sub OptionButton1_Click()

End Sub

Public Sub Tilbage_Click()
Me.Hide
SFunc.ShowFunc ("frm008")
'frm008.Show
End Sub

Private Sub UserForm_Initialize()

Image1.PictureSizeMode = fmPictureSizeModeStretch

' Indl�s tidligere svar 9a2
If Worksheets("SpmSvar").Range("D19:D19").Value = "Ja" Then
    OptionButton1.Value = True
ElseIf Worksheets("SpmSvar").Range("D19:D19").Value = "Nej" Then
    OptionButton2.Value = True
Else
    OptionButton1.Value = False
    OptionButton2.Value = False
End If



End Sub